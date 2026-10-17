using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Moves a place in the text by characters, words, sentences or paragraphs one block at a time (a paragraph, or the end of a table row), so
/// that a move of a few units looks only at the paragraphs it crosses and never at the whole story. Units do not run across a block, so
/// moving across the story this way gives the same places as moving through the text of the whole story.
/// </summary>
internal static class UnitNavigator
{
    /// <summary>A paragraph (its text and its mark) or the end of a table row: the pieces the story is made of for moving about.</summary>
    private sealed class Block
    {
        public Block(Paragraph paragraph)
        {
            Element = paragraph;
            Map = StoryMap.Build(paragraph);
            Characters = Map.ToCharacters();
        }

        public Block(TableRow row)
        {
            Element = row;
            Characters = new[] { '\a' };
        }

        public OpenXmlElement Element { get; }

        public StoryMap Map { get; }

        public char[] Characters { get; }

        public int Length
        {
            get { return Characters.Length; }
        }

        private readonly Dictionary<TextUnit, List<int>> _boundaries = new Dictionary<TextUnit, List<int>>();

        public List<int> Boundaries(TextUnit unit)
        {
            if (!_boundaries.TryGetValue(unit, out List<int> list))
            {
                list = TextUnits.GetBoundaries(Characters, unit);
                _boundaries[unit] = list;
            }
            return list;
        }

        public Anchor AnchorAt(int position)
        {
            return Map == null ? new Anchor(Element, 0) : Anchor.At(Map, position);
        }
    }

    /// <summary>
    /// Moves an anchor by <paramref name="count"/> units (backwards when negative). False when the anchor is not on text that can be found
    /// in the story. The number of units moved is less than asked at the start or end of the story.
    /// </summary>
    public static bool TryMove(Anchor from, OpenXmlElement story, TextUnit unit, int count, out Anchor result, out int moved)
    {
        result = null;
        moved = 0;
        if (!TryBlockOf(from, story, out Block block, out int position))
        {
            return false;
        }

        int steps = Math.Abs(count);
        bool atEnd = false;
        int done = 0;
        while (done < steps)
        {
            if (count > 0)
            {
                if (position >= block.Length)
                {
                    break;
                }
                position = block.Boundaries(unit).First(b => b > position);
                done++;
                if (position >= block.Length)
                {
                    //The end of a block is the start of the next.
                    Block next = NextBlock(block, story);
                    if (next == null)
                    {
                        atEnd = true;
                        break;
                    }
                    block = next;
                    position = 0;
                }
            }
            else
            {
                if (position == 0)
                {
                    Block previous = PreviousBlock(block, story);
                    if (previous == null)
                    {
                        break;
                    }
                    block = previous;
                    position = block.Length;
                }
                List<int> boundaries = block.Boundaries(unit);
                position = boundaries[boundaries.FindLastIndex(b => b < position)];
                done++;
            }
        }

        result = atEnd ? new Anchor(null, 0) : block.AnchorAt(position);
        moved = count > 0 ? done : -done;
        return true;
    }

    /// <summary>
    /// Whether the first place is before (negative), the same as (zero) or after (positive) the second, or false if either cannot be found.
    /// </summary>
    public static bool TryCompare(Anchor first, Anchor second, OpenXmlElement story, out int order)
    {
        order = 0;
        if (first.Element == null || second.Element == null)
        {
            order = first.Element == null ? (second.Element == null ? 0 : 1) : -1;
            return true;
        }
        if (!TryBlockOf(first, story, out Block firstBlock, out int firstAt) || !TryBlockOf(second, story, out Block secondBlock, out int secondAt))
        {
            return false;
        }
        if (ReferenceEquals(firstBlock.Element, secondBlock.Element))
        {
            order = firstAt.CompareTo(secondAt);
            return true;
        }
        order = Precedes(firstBlock.Element, secondBlock.Element) ? -1 : 1;
        return true;
    }

    //Whether a block comes before another in the story. The end of a row comes after everything in the row.
    private static bool Precedes(OpenXmlElement first, OpenXmlElement second)
    {
        if (first is TableRow && second.Ancestors().Any(a => ReferenceEquals(a, first)))
        {
            return false;
        }
        if (second is TableRow && first.Ancestors().Any(a => ReferenceEquals(a, second)))
        {
            return true;
        }
        return first.IsBefore(second);
    }

    private static bool TryBlockOf(Anchor anchor, OpenXmlElement story, out Block block, out int position)
    {
        block = null;
        position = 0;
        if (anchor.Element == null)
        {
            block = LastBlock(story);
            position = block == null ? 0 : block.Length;
            return block != null;
        }
        if (anchor.Element is TableRow row)
        {
            block = new Block(row);
            return true;
        }

        Paragraph paragraph = anchor.Element as Paragraph ?? anchor.Element.Ancestors<Paragraph>().FirstOrDefault();
        if (paragraph == null || !ReferenceEquals(StoryMap.FindRoot(paragraph), story))
        {
            return false;
        }
        block = new Block(paragraph);
        position = RangeWrapper.ResolveLocal(block.Map, anchor);
        return position >= 0;
    }

    private static Block NextBlock(Block block, OpenXmlElement story)
    {
        Anchor next = StoryMap.AnchorAfter(block.Element, story);
        if (next.Element == null)
        {
            return null;
        }
        if (next.Element is TableRow row)
        {
            return new Block(row);
        }
        Paragraph paragraph = next.Element as Paragraph ?? next.Element.Ancestors<Paragraph>().FirstOrDefault();
        return paragraph == null ? null : new Block(paragraph);
    }

    private static Block PreviousBlock(Block block, OpenXmlElement story)
    {
        //What comes before the end of a table row is the last paragraph in the row.
        if (block.Element is TableRow row)
        {
            for (OpenXmlElement cell = row.LastChild; cell != null; cell = cell.PreviousSibling())
            {
                Block inside = LastBlockIn(cell);
                if (inside != null)
                {
                    return inside;
                }
            }
        }

        OpenXmlElement current = block.Element;
        while (current != null && !ReferenceEquals(current, story))
        {
            for (OpenXmlElement sibling = current.PreviousSibling(); sibling != null; sibling = sibling.PreviousSibling())
            {
                Block found = LastBlockIn(sibling);
                if (found != null)
                {
                    return found;
                }
            }
            current = current.Parent;
        }
        return null;
    }

    private static Block LastBlock(OpenXmlElement story)
    {
        return LastBlockIn(story);
    }

    //The last paragraph or end of a table row in an element, found from the end without looking at the rest.
    private static Block LastBlockIn(OpenXmlElement element)
    {
        if (element is Paragraph paragraph)
        {
            return new Block(paragraph);
        }
        if (element is TableRow row)
        {
            return new Block(row);
        }
        if (element is Run || element is SectionProperties || !(element is OpenXmlCompositeElement))
        {
            return null;
        }
        for (OpenXmlElement child = element.LastChild; child != null; child = child.PreviousSibling())
        {
            Block found = LastBlockIn(child);
            if (found != null)
            {
                return found;
            }
        }
        return null;
    }
}
