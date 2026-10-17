using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Character positions. A range can be made of whole elements (a paragraph, table, bookmark and so on) or of a span of characters between two
/// positions, as in Word. <see cref="Start"/> and <see cref="End"/> give the positions of either kind, and <see cref="SetRange"/>,
/// <see cref="Collapse"/>, <see cref="MoveStart"/> and the like make a range of a span of characters.
/// </summary>
public partial class RangeWrapper
{
    private bool _isSpan;
    private OpenXmlElement _root;
    private int _spanStart;
    private int _spanEnd;

    internal RangeWrapper(OpenXmlElement root, int start, int end)
    {
        _isSpan = true;
        _root = root;
        _spanStart = start;
        _spanEnd = end;
    }

    /// <summary>
    /// True when the range is a span of characters rather than made of elements.
    /// </summary>
    internal bool IsSpan
    {
        get { return _isSpan; }
    }

    /// <summary>
    /// The position of the first character of the range, counted from the start of the story (the body, header, footer or comment) it is in. A
    /// position counts each character of the text, each tab and break, each end of paragraph (which is also the end of a table cell for the
    /// last paragraph of a cell) and each end of table row. Pictures and field codes are not counted. Setting it moves the start of the range,
    /// and the end too when the start would pass it.
    /// </summary>
    /// <remarks>
    /// Positions are a snapshot: unlike in Word they are not adjusted when the document is changed by something other than the range itself.
    /// </remarks>
    public int Start
    {
        get
        {
            GetPositions(out _, out int start, out _);
            return start;
        }
        set
        {
            GetPositions(out _, out _, out int end);
            SetRange(value, Math.Max(value, end));
        }
    }

    /// <summary>
    /// The position after the last character of the range. Setting it moves the end, and the start too when the end would pass it.
    /// </summary>
    public int End
    {
        get
        {
            GetPositions(out _, out _, out int end);
            return end;
        }
        set
        {
            GetPositions(out _, out int start, out _);
            SetRange(Math.Min(start, value), value);
        }
    }

    /// <summary>
    /// The number of character positions in the story the range is in.
    /// </summary>
    public int StoryLength
    {
        get
        {
            GetPositions(out StoryMap map, out _, out _);
            return map.Length;
        }
    }

    /// <summary>
    /// Makes the range the span of characters from <paramref name="start"/> up to <paramref name="end"/> in the same story, as
    /// <c>Range.SetRange</c> in Word. The range is no longer made of elements; text formatted or replaced through it is split out of the
    /// runs that hold it.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The positions are not in order, or are beyond the end of the story.</exception>
    public void SetRange(int start, int end)
    {
        GetPositions(out StoryMap map, out _, out _);
        if (start < 0 || end < start || end > map.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(end), String.Format(CultureInfo.InvariantCulture, "The range {0} to {1} is not inside the story, which is {2} positions long.", start, end, map.Length));
        }
        MakeSpan(map.Root, start, end);
    }

    /// <summary>
    /// Collapses the range to its start or end, so that it is an insertion point, as <c>Range.Collapse</c> in Word.
    /// </summary>
    public void Collapse(CollapseDirection direction = CollapseDirection.Start)
    {
        GetPositions(out StoryMap map, out int start, out int end);
        int position = direction == CollapseDirection.Start ? start : end;
        MakeSpan(map.Root, position, position);
    }

    /// <summary>
    /// Moves the start of the range by <paramref name="count"/> units (backwards when negative). If it passes the end the range is
    /// collapsed there. Returns the number of units it moved, which is less than asked at the start or end of the story.
    /// </summary>
    public int MoveStart(TextUnit unit, int count = 1)
    {
        GetPositions(out StoryMap map, out int start, out int end);
        int moved = Shift(map, unit, start, count, out int position);
        MakeSpan(map.Root, position, Math.Max(position, end));
        return moved;
    }

    /// <summary>
    /// Moves the end of the range by <paramref name="count"/> units (backwards when negative). If it passes the start the range is
    /// collapsed there. Returns the number of units it moved.
    /// </summary>
    public int MoveEnd(TextUnit unit, int count = 1)
    {
        GetPositions(out StoryMap map, out int start, out int end);
        int moved = Shift(map, unit, end, count, out int position);
        MakeSpan(map.Root, Math.Min(start, position), position);
        return moved;
    }

    /// <summary>
    /// Collapses the range to its end and then moves that point by <paramref name="count"/> units, as <c>Range.Move</c> in Word. Returns the
    /// number of units it moved.
    /// </summary>
    public int Move(TextUnit unit, int count = 1)
    {
        GetPositions(out StoryMap map, out _, out int end);
        int moved = Shift(map, unit, end, count, out int position);
        MakeSpan(map.Root, position, position);
        return moved;
    }

    /// <summary>
    /// Widens the range to whole units, as <c>Range.Expand</c> in Word: to the start of the first unit it touches and the end of the last. A
    /// collapsed range grows to the unit it is in.
    /// </summary>
    public void Expand(TextUnit unit)
    {
        GetPositions(out StoryMap map, out int start, out int end);
        List<int> boundaries = TextUnits.GetBoundaries(map.ToCharacters(), unit);
        int newStart = boundaries.LastOrDefault(b => b <= start);
        int newEnd = end > start ? boundaries.FirstOrDefault(b => b >= end) : boundaries.FirstOrDefault(b => b > start);
        if (newEnd < end)
        {
            //There is no boundary after the position: it is at the end of the story.
            newEnd = map.Length;
        }
        MakeSpan(map.Root, newStart, Math.Max(newEnd, newStart));
    }

    /// <summary>
    /// The characters of the range, one range each, as <c>Range.Characters</c> in Word.
    /// </summary>
    public TextUnitCollection Characters
    {
        get { return new TextUnitCollection(this, TextUnit.Character); }
    }

    /// <summary>
    /// The words of the range, one range each (with the spaces after the word), as <c>Range.Words</c> in Word.
    /// </summary>
    public TextUnitCollection Words
    {
        get { return new TextUnitCollection(this, TextUnit.Word); }
    }

    /// <summary>
    /// The sentences of the range, one range each, as <c>Range.Sentences</c> in Word.
    /// </summary>
    public TextUnitCollection Sentences
    {
        get { return new TextUnitCollection(this, TextUnit.Sentence); }
    }

    #region Positions

    private void MakeSpan(OpenXmlElement root, int start, int end)
    {
        _isSpan = true;
        _root = root;
        _spanStart = start;
        _spanEnd = end;
        _startOpenXmlElement = null;
        _endOpenXmlElement = null;
    }

    /// <summary>
    /// A new range over the span of characters from <paramref name="start"/> to <paramref name="end"/> in the story of this range.
    /// </summary>
    internal RangeWrapper CreateSibling(int start, int end)
    {
        GetPositions(out StoryMap map, out _, out _);
        return new RangeWrapper(map.Root, start, end);
    }

    /// <summary>
    /// Builds the map of the story of the range and finds the positions of the start and end of the range in it.
    /// </summary>
    internal void GetPositions(out StoryMap map, out int start, out int end)
    {
        if (_isSpan)
        {
            map = StoryMap.Build(_root);
            start = Math.Min(_spanStart, map.Length);
            end = Math.Min(_spanEnd, map.Length);
            return;
        }

        map = StoryMap.Build(StoryMap.FindRoot(_startOpenXmlElement));
        start = map.Length;
        foreach (Atom atom in map.Atoms)
        {
            if (!Precedes(atom, _startOpenXmlElement))
            {
                start = atom.Start;
                break;
            }
        }

        end = 0;
        for (int i = map.Atoms.Count - 1; i >= 0; i--)
        {
            Atom atom = map.Atoms[i];
            if (IsInside(atom, _endOpenXmlElement) || Precedes(atom, _endOpenXmlElement))
            {
                end = atom.End;
                break;
            }
        }
        end = Math.Max(end, start);
    }

    private static bool IsInside(Atom atom, OpenXmlElement element)
    {
        return atom.Element == element || atom.Element.Ancestors().Any(a => a == element);
    }

    //True when the atom is wholly before the element in the document.
    private static bool Precedes(Atom atom, OpenXmlElement element)
    {
        if (IsInside(atom, element))
        {
            return false;
        }
        if (atom.IsMark)
        {
            //The marker comes at the end of its paragraph or row, so an element inside that paragraph or row comes before it.
            if (element.Ancestors().Any(a => a == atom.Element))
            {
                return false;
            }
            return atom.Element.IsBefore(element);
        }
        return atom.Element.IsBefore(element);
    }

    /// <summary>
    /// Moves a position by units. Returns how many units it moved.
    /// </summary>
    private static int Shift(StoryMap map, TextUnit unit, int position, int count, out int result)
    {
        if (count == 0)
        {
            result = position;
            return 0;
        }
        List<int> boundaries = TextUnits.GetBoundaries(map.ToCharacters(), unit);
        int moved = 0;
        int current = position;
        while (moved < Math.Abs(count))
        {
            int next;
            if (count > 0)
            {
                next = boundaries.FirstOrDefault(b => b > current);
                if (next <= current)
                {
                    break;
                }
            }
            else
            {
                int index = boundaries.FindLastIndex(b => b < current);
                if (index < 0)
                {
                    break;
                }
                next = boundaries[index];
            }
            current = next;
            moved++;
        }
        result = current;
        return count > 0 ? moved : -moved;
    }

    #endregion

    #region Elements of a span

    /// <summary>
    /// A range of elements for the same text, for the collections that work on elements (bookmarks, paragraphs and so on). Runs are split
    /// so that the first and last element hold only text in the range. A range of elements is its own element range.
    /// </summary>
    /// <exception cref="NotSupportedException">The range is collapsed, so it has no elements.</exception>
    internal RangeWrapper ElementRange()
    {
        if (!_isSpan)
        {
            return this;
        }

        StoryMap map = SplitAtBoundaries(out int start, out int end);
        List<Atom> atoms = map.Atoms.Where(a => a.Start >= start && a.End <= end).ToList();
        if (atoms.Count == 0)
        {
            throw new NotSupportedException("A range with no characters has no elements. Use a range with at least one character.");
        }
        return new RangeWrapper(ElementOf(atoms[0]), ElementOf(atoms[atoms.Count - 1]));
    }

    private static OpenXmlElement ElementOf(Atom atom)
    {
        return atom.IsMark ? atom.Element : atom.Element.Parent;
    }

    /// <summary>
    /// Splits the runs at the start and end of the span so that the range is made of whole runs, and returns the map of the new text.
    /// </summary>
    private StoryMap SplitAtBoundaries(out int start, out int end)
    {
        GetPositions(out StoryMap map, out start, out end);
        map = map.EnsureBoundary(end);
        map = map.EnsureBoundary(start);
        return map;
    }

    /// <summary>
    /// The runs that hold the characters of the span, grouped by their paragraph, after splitting the runs at the ends of the span.
    /// </summary>
    private IList<List<Run>> GetSpanRunsByParagraph()
    {
        StoryMap map = SplitAtBoundaries(out int start, out int end);
        return map.Atoms
            .Where(a => !a.IsMark && a.Start >= start && a.End <= end)
            .Select(a => (Run)a.Element.Parent)
            .Distinct()
            .GroupBy(r => r.Ancestors<Paragraph>().First())
            .Select(g => g.ToList())
            .ToList();
    }

    /// <summary>
    /// The paragraphs whose paragraph marks are in the span.
    /// </summary>
    private IList<Paragraph> GetSpanParagraphs()
    {
        GetPositions(out StoryMap map, out int start, out int end);
        return map.Atoms
            .Where(a => a.Kind == AtomKind.ParagraphMark && a.Start >= start && a.End <= end)
            .Select(a => (Paragraph)a.Element)
            .ToList();
    }

    /// <summary>
    /// The paragraphs that the range touches at all, which is where paragraph formatting applies. A collapsed range touches the paragraph it is in.
    /// </summary>
    internal IList<Paragraph> GetTouchedParagraphs()
    {
        GetPositions(out StoryMap map, out int start, out int end);
        List<Paragraph> paragraphs = new List<Paragraph>();
        foreach (Atom atom in map.Atoms)
        {
            bool touches = start == end ? (atom.Start <= start && start < atom.End) : (atom.Start < end && atom.End > start);
            if (!touches && !(start == end && start == map.Length && atom == map.Atoms.LastOrDefault()))
            {
                continue;
            }
            Paragraph paragraph = atom.Kind == AtomKind.ParagraphMark ? (Paragraph)atom.Element : atom.Kind == AtomKind.RowMark ? null : atom.Element.Ancestors<Paragraph>().First();
            if (paragraph != null && !paragraphs.Contains(paragraph))
            {
                paragraphs.Add(paragraph);
            }
        }
        return paragraphs;
    }

    #endregion
}
