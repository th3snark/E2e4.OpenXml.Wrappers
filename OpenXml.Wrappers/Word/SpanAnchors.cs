using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A place in the text of a story that stays with the text when other things are inserted or removed around it: a character offset in a text
/// element, or the start of a tab, break, paragraph mark (a <see cref="Paragraph"/>) or end of table row (a <see cref="TableRow"/>), or the
/// end of the story when <see cref="Element"/> is null.
/// </summary>
internal sealed class Anchor
{
    public Anchor(OpenXmlElement element, int offset)
    {
        Element = element;
        Offset = offset;
    }

    public OpenXmlElement Element { get; set; }

    public int Offset { get; set; }

    public Anchor Clone()
    {
        return new Anchor(Element, Offset);
    }

    public bool SameAs(Anchor other)
    {
        return other != null && ReferenceEquals(Element, other.Element) && Offset == other.Offset;
    }

    /// <summary>
    /// The anchor for a position in the map: the character at that position (or the end of the story when it is beyond the map).
    /// </summary>
    public static Anchor At(StoryMap map, int position)
    {
        int index = map.IndexAt(position);
        if (index < 0)
        {
            return map.EndAnchor();
        }
        Atom atom = map.Atoms[index];
        return new Anchor(atom.Element, position - atom.Start);
    }
}

/// <summary>
/// Remembers the ranges that are spans of characters in each story so that, when the library changes the text they point into, they can be
/// corrected: a text element that is split in two, text that is replaced in place, a paragraph that is split or joined, and text that is
/// removed. Ranges are held weakly.
/// </summary>
internal static class SpanRegistry
{
    private static readonly ConditionalWeakTable<OpenXmlElement, List<WeakReference<RangeWrapper>>> Stories = new ConditionalWeakTable<OpenXmlElement, List<WeakReference<RangeWrapper>>>();

    public static void Register(OpenXmlElement storyRoot, RangeWrapper range)
    {
        List<WeakReference<RangeWrapper>> list = Stories.GetOrCreateValue(storyRoot);
        lock (list)
        {
            list.Add(new WeakReference<RangeWrapper>(range));
        }
    }

    private static List<RangeWrapper> Live(OpenXmlElement element)
    {
        List<RangeWrapper> live = new List<RangeWrapper>();
        if (!Stories.TryGetValue(StoryMap.FindRoot(element), out List<WeakReference<RangeWrapper>> list))
        {
            return live;
        }
        lock (list)
        {
            list.RemoveAll(w => !w.TryGetTarget(out _));
            foreach (WeakReference<RangeWrapper> reference in list)
            {
                if (reference.TryGetTarget(out RangeWrapper range))
                {
                    live.Add(range);
                }
            }
        }
        return live;
    }

    /// <summary>
    /// A text element was split: what was after the first <paramref name="offset"/> characters of <paramref name="text"/> is now in <paramref name="suffix"/>.
    /// </summary>
    public static void TextSplit(Text text, Text suffix, int offset)
    {
        foreach (RangeWrapper range in Live(text))
        {
            foreach (Anchor anchor in range.Anchors())
            {
                if (ReferenceEquals(anchor.Element, text) && anchor.Offset > offset)
                {
                    anchor.Element = suffix;
                    anchor.Offset -= offset;
                }
            }
        }
    }

    /// <summary>
    /// <paramref name="oldLength"/> characters at <paramref name="start"/> in a text element were replaced in place by <paramref name="newLength"/>.
    /// </summary>
    public static void TextEdited(Text text, int start, int oldLength, int newLength)
    {
        foreach (RangeWrapper range in Live(text))
        {
            foreach (Anchor anchor in range.Anchors())
            {
                if (!ReferenceEquals(anchor.Element, text))
                {
                    continue;
                }
                if (anchor.Offset >= start + oldLength)
                {
                    anchor.Offset += newLength - oldLength;
                }
                else if (anchor.Offset > start)
                {
                    anchor.Offset = start;
                }
            }
        }
    }

    /// <summary>
    /// A paragraph was split in two: its paragraph mark, which is at the end of what is now the second paragraph, moved with it.
    /// </summary>
    public static void ParagraphSplit(Paragraph original, Paragraph second)
    {
        foreach (RangeWrapper range in Live(original))
        {
            foreach (Anchor anchor in range.Anchors())
            {
                if (ReferenceEquals(anchor.Element, original))
                {
                    anchor.Element = second;
                }
            }
        }
    }

    /// <summary>
    /// A paragraph was joined to the one before it, whose mark now ends both.
    /// </summary>
    public static void ParagraphMerged(Paragraph removed, Paragraph into)
    {
        foreach (RangeWrapper range in Live(into))
        {
            foreach (Anchor anchor in range.Anchors())
            {
                if (ReferenceEquals(anchor.Element, removed))
                {
                    anchor.Element = into;
                }
            }
        }
    }

    /// <summary>
    /// Text is about to be removed. Ranges that point into it are moved to <paramref name="after"/>, the place where it was.
    /// </summary>
    public static void Removing(OpenXmlElement inStory, HashSet<OpenXmlElement> removed, Anchor after)
    {
        foreach (RangeWrapper range in Live(inStory))
        {
            foreach (Anchor anchor in range.Anchors())
            {
                if (anchor.Element != null && removed.Contains(anchor.Element))
                {
                    anchor.Element = after.Element;
                    anchor.Offset = after.Offset;
                }
            }
        }
    }
}
