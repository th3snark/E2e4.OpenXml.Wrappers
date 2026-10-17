using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A place in the text of a story that stays with the text when other things are inserted or removed around it: a character offset in a text
/// element, or the start of a tab, break, paragraph mark (a <see cref="Paragraph"/>), end of table row (a <see cref="TableRow"/>) or other
/// piece of the text, or the end of the story when <see cref="Element"/> is null. An anchor never changes; a range that has to move one
/// replaces it with another.
/// </summary>
/// <remarks>
/// An anchor also remembers pieces of text near it, before and after (each of the first eight pieces away, then sixteen and thirty-two).
/// If the text it is on is removed without the library being told, the nearest of those that is still there says where the text was: exactly,
/// when no more than eight pieces went, and roughly (within the gap to the next one remembered) when more did.
/// </remarks>
internal sealed class Anchor
{
    //How far away the neighbours are, in pieces of text: close ones for small removals and far ones for large.
    private static readonly int[] Distances = { 1, 2, 3, 4, 5, 6, 7, 8, 16, 32 };

    private static readonly Anchor[] None = Array.Empty<Anchor>();

    public Anchor(OpenXmlElement element, int offset)
        : this(element, offset, None, None)
    {
    }

    private Anchor(OpenXmlElement element, int offset, Anchor[] before, Anchor[] after)
    {
        Element = element;
        Offset = offset;
        Before = before;
        After = after;
    }

    public OpenXmlElement Element { get; }

    /// <summary>The characters into the element. For a neighbour that is before, -1 is the end of the element.</summary>
    public int Offset { get; }

    /// <summary>The ends of the pieces of text before this one, nearest first.</summary>
    public Anchor[] Before { get; }

    /// <summary>The starts of the pieces of text after this one, nearest first.</summary>
    public Anchor[] After { get; }

    public bool SameAs(Anchor other)
    {
        return other != null && ReferenceEquals(Element, other.Element) && Offset == other.Offset;
    }

    /// <summary>The same place on other text, as when text moves to another element.</summary>
    public Anchor On(OpenXmlElement element, int offset)
    {
        return new Anchor(element, offset, Before, After);
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
        return new Anchor(atom.Element, position - atom.Start, Pick(map.ElementsBefore(index, Reach), -1), Pick(map.ElementsAfter(index, Reach), 0));
    }

    /// <summary>
    /// This anchor with the neighbours as they are now in the map, if the map has the text it is on, otherwise this anchor as it is.
    /// Neighbours the map cannot see (those beyond a map of one paragraph) are kept.
    /// </summary>
    public Anchor Refreshed(StoryMap map)
    {
        if (Element == null || !map.TryGetAtom(Element, out Atom atom))
        {
            return this;
        }

        Anchor[] before = Merge(Pick(map.ElementsBefore(atom.Index, Reach, false), -1), Before);
        Anchor[] after = Merge(Pick(map.ElementsAfter(atom.Index, Reach, false), 0), After);
        return Same(before, Before) && Same(after, After) ? this : new Anchor(Element, Offset, before, after);
    }

    /// <summary>
    /// This anchor with neighbours from the pieces of text walked past before it (the nearest last) and after it (the nearest first).
    /// </summary>
    public Anchor WithNeighbours(IReadOnlyList<StoryMap.AtomInfo> before, IReadOnlyList<StoryMap.AtomInfo> after)
    {
        List<Anchor> previous = new List<Anchor>();
        List<Anchor> next = new List<Anchor>();
        foreach (int distance in Distances)
        {
            if (distance <= before.Count)
            {
                previous.Add(new Anchor(before[before.Count - distance].Element, -1));
            }
            if (distance <= after.Count)
            {
                next.Add(new Anchor(after[distance - 1].Element, 0));
            }
        }
        return new Anchor(Element, Offset, previous.ToArray(), next.ToArray());
    }

    /// <summary>How many pieces of text are needed beyond an anchor to give all its neighbours.</summary>
    public static int Reach
    {
        get { return Distances[Distances.Length - 1]; }
    }

    //The neighbours at the distances kept, from the pieces of text on one side nearest first.
    private static Anchor[] Pick(List<OpenXmlElement> nearestFirst, int offset)
    {
        List<Anchor> list = new List<Anchor>();
        foreach (int distance in Distances)
        {
            if (distance <= nearestFirst.Count)
            {
                list.Add(new Anchor(nearestFirst[distance - 1], offset));
            }
        }
        return list.ToArray();
    }

    //The fresh neighbours, then those of the old that are not repeats, up to the number there can be.
    private static Anchor[] Merge(Anchor[] fresh, Anchor[] old)
    {
        if (fresh.Length >= Distances.Length)
        {
            return fresh;
        }
        List<Anchor> merged = new List<Anchor>(fresh);
        foreach (Anchor each in old)
        {
            if (merged.Count >= Distances.Length)
            {
                break;
            }
            if (!merged.Any(m => ReferenceEquals(m.Element, each.Element)))
            {
                merged.Add(each);
            }
        }
        return merged.ToArray();
    }

    private static bool Same(Anchor[] first, Anchor[] second)
    {
        if (first.Length != second.Length)
        {
            return false;
        }
        for (int i = 0; i < first.Length; i++)
        {
            if (!ReferenceEquals(first[i].Element, second[i].Element) || first[i].Offset != second[i].Offset)
            {
                return false;
            }
        }
        return true;
    }
}

/// <summary>
/// Remembers the ranges that are spans of characters in each story, by the text their ends are on, so that when the library changes that text
/// they can be corrected: a text element that is split in two, text that is replaced in place, a paragraph that is split or joined, and
/// text that is removed. Ranges are held weakly. It is safe to use from several threads, though the document itself, like the Open XML SDK,
/// is not safe to change from several at once.
/// </summary>
internal static class SpanRegistry
{
    private sealed class Story
    {
        public readonly object Lock = new object();

        //For each piece of text that a range has an end on, the ranges that have one.
        public readonly Dictionary<OpenXmlElement, List<WeakReference<RangeWrapper>>> ByElement = new Dictionary<OpenXmlElement, List<WeakReference<RangeWrapper>>>(ReferenceEqualityComparer.Instance);

        public int SinceTidy;
    }

    private static readonly ConditionalWeakTable<OpenXmlElement, Story> Stories = new ConditionalWeakTable<OpenXmlElement, Story>();

    /// <summary>
    /// Notes the pieces of text that the ends of a range are on, so that changes to them reach the range. Call it whenever a range gets new
    /// ends.
    /// </summary>
    public static void Index(OpenXmlElement storyRoot, RangeWrapper range)
    {
        Story story = Stories.GetOrCreateValue(storyRoot);
        lock (story.Lock)
        {
            foreach (OpenXmlElement element in range.EndElements())
            {
                Add(story, element, range);
            }

            //Ranges that have gone leave their entries behind, so now and then they are cleared out.
            if (++story.SinceTidy > 5000)
            {
                story.SinceTidy = 0;
                Tidy(story);
            }
        }
    }

    private static void Add(Story story, OpenXmlElement element, RangeWrapper range)
    {
        if (!story.ByElement.TryGetValue(element, out List<WeakReference<RangeWrapper>> list))
        {
            list = new List<WeakReference<RangeWrapper>>();
            story.ByElement[element] = list;
        }
        if (list.Count == 0 || !list[list.Count - 1].TryGetTarget(out RangeWrapper last) || !ReferenceEquals(last, range))
        {
            list.Add(new WeakReference<RangeWrapper>(range));
        }
    }

    private static void Tidy(Story story)
    {
        foreach (OpenXmlElement element in story.ByElement.Keys.ToList())
        {
            List<WeakReference<RangeWrapper>> list = story.ByElement[element];
            list.RemoveAll(w => !w.TryGetTarget(out RangeWrapper range) || !range.HasEndOn(element));
            if (list.Count == 0)
            {
                story.ByElement.Remove(element);
            }
        }
    }

    //The ranges that have an end on the element now, with those that no longer do dropped from the list.
    private static List<RangeWrapper> RangesOn(Story story, OpenXmlElement element)
    {
        List<RangeWrapper> ranges = new List<RangeWrapper>();
        if (story.ByElement.TryGetValue(element, out List<WeakReference<RangeWrapper>> list))
        {
            list.RemoveAll(w => !w.TryGetTarget(out RangeWrapper range) || !range.HasEndOn(element));
            //A range can be in the list more than once; each is corrected once.
            HashSet<RangeWrapper> seen = new HashSet<RangeWrapper>(ReferenceEqualityComparer.Instance);
            foreach (WeakReference<RangeWrapper> reference in list)
            {
                if (reference.TryGetTarget(out RangeWrapper range) && seen.Add(range))
                {
                    ranges.Add(range);
                }
            }
            if (list.Count == 0)
            {
                story.ByElement.Remove(element);
            }
        }
        return ranges;
    }

    private static void Change(OpenXmlElement inStory, OpenXmlElement element, Func<Anchor, Anchor> change)
    {
        if (!Stories.TryGetValue(StoryMap.FindRoot(inStory), out Story story))
        {
            return;
        }
        lock (story.Lock)
        {
            foreach (RangeWrapper range in RangesOn(story, element))
            {
                range.Remap(change);
                foreach (OpenXmlElement now in range.EndElements())
                {
                    Add(story, now, range);
                }
            }
        }
    }

    /// <summary>
    /// A text element was split: what was after the first <paramref name="offset"/> characters of <paramref name="text"/> is now in <paramref name="suffix"/>.
    /// </summary>
    public static void TextSplit(OpenXmlElement text, OpenXmlElement suffix, int offset)
    {
        Change(text, text, anchor => ReferenceEquals(anchor.Element, text) && anchor.Offset > offset ? anchor.On(suffix, anchor.Offset - offset) : anchor);
    }

    /// <summary>
    /// <paramref name="oldLength"/> characters at <paramref name="start"/> in a text element were replaced in place by <paramref name="newLength"/>.
    /// </summary>
    public static void TextEdited(Text text, int start, int oldLength, int newLength)
    {
        Change(text, text, anchor =>
        {
            if (!ReferenceEquals(anchor.Element, text))
            {
                return anchor;
            }
            if (anchor.Offset >= start + oldLength)
            {
                return anchor.On(text, anchor.Offset + newLength - oldLength);
            }
            return anchor.Offset > start ? anchor.On(text, start) : anchor;
        });
    }

    /// <summary>
    /// A paragraph was split in two: its paragraph mark, which is at the end of what is now the second paragraph, moved with it.
    /// </summary>
    public static void ParagraphSplit(Paragraph original, Paragraph second)
    {
        Change(original, original, anchor => ReferenceEquals(anchor.Element, original) ? anchor.On(second, anchor.Offset) : anchor);
    }

    /// <summary>
    /// A paragraph was joined to the one before it, whose mark now ends both.
    /// </summary>
    public static void ParagraphMerged(Paragraph removed, Paragraph into)
    {
        Change(into, removed, anchor => ReferenceEquals(anchor.Element, removed) ? anchor.On(into, anchor.Offset) : anchor);
    }

    /// <summary>
    /// Text is about to be removed. Ranges that have an end on it are moved to <paramref name="after"/>, the place where it was.
    /// </summary>
    public static void Removing(OpenXmlElement inStory, HashSet<OpenXmlElement> removed, Anchor after)
    {
        if (!Stories.TryGetValue(StoryMap.FindRoot(inStory), out Story story))
        {
            return;
        }
        lock (story.Lock)
        {
            foreach (OpenXmlElement element in removed)
            {
                foreach (RangeWrapper range in RangesOn(story, element))
                {
                    range.Remap(anchor => anchor.Element != null && removed.Contains(anchor.Element) ? after : anchor);
                    foreach (OpenXmlElement now in range.EndElements())
                    {
                        Add(story, now, range);
                    }
                }
                story.ByElement.Remove(element);
            }
        }
    }
}
