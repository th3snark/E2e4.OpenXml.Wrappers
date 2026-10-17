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
/// <remarks>
/// A span does not remember integer positions: its two ends are anchored to the text itself (see <see cref="Anchor"/>), so it follows the
/// text it covers when other text is inserted or removed anywhere, whatever made the change. The integers are worked out when asked for. Most
/// operations on a span that lies inside one paragraph only look at that paragraph.
/// </remarks>
public partial class RangeWrapper
{
    private bool _isSpan;
    private OpenXmlElement _root;
    private Anchor _startAnchor;
    private Anchor _endAnchor;

    //The positions when they were last worked out. They are used only if the text a range was anchored to has gone.
    private int _lastStart;
    private int _lastEnd;

    internal RangeWrapper(StoryMap map, int start, int end)
    {
        MakeSpan(map, start, end);
    }

    /// <summary>
    /// True when the range is a span of characters rather than made of elements.
    /// </summary>
    internal bool IsSpan
    {
        get { return _isSpan; }
    }

    /// <summary>
    /// An element in the story of the range, for looking things up in the document it belongs to.
    /// </summary>
    internal OpenXmlElement AnchorElement
    {
        get { return _isSpan ? _root : _startOpenXmlElement; }
    }

    internal IEnumerable<Anchor> Anchors()
    {
        if (_isSpan)
        {
            yield return _startAnchor;
            yield return _endAnchor;
        }
    }

    internal Anchor StartAnchor
    {
        get { return _startAnchor; }
    }

    internal Anchor EndAnchor
    {
        get { return _endAnchor; }
    }

    /// <summary>
    /// The position of the first character of the range, counted from the start of the story (the body, header, footer or comment) it is in. A
    /// position counts each character of the text, each tab and break, each end of paragraph (which is also the end of a table cell for the
    /// last paragraph of a cell) and each end of table row. Pictures and field codes are not counted. Setting it moves the start of the range,
    /// and the end too when the start would pass it.
    /// </summary>
    /// <remarks>
    /// A range follows its text, so the positions change when text is added or removed before it. They are worked out each time they are read,
    /// by looking through the story, so keep a range (or its words) rather than calling this over and over on a very large document.
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
        MakeSpan(map, start, end);
    }

    /// <summary>
    /// Collapses the range to its start or end, so that it is an insertion point, as <c>Range.Collapse</c> in Word.
    /// </summary>
    public void Collapse(CollapseDirection direction = CollapseDirection.Start)
    {
        GetScoped(false, out StoryMap map, out int start, out int end);
        int position = direction == CollapseDirection.Start ? start : end;
        MakeSpan(map, position, position);
    }

    /// <summary>
    /// Moves the start of the range by <paramref name="count"/> units (backwards when negative). If it passes the end the range is
    /// collapsed there. Returns the number of units it moved, which is less than asked at the start or end of the story.
    /// </summary>
    public int MoveStart(TextUnit unit, int count = 1)
    {
        GetPositions(out StoryMap map, out int start, out int end);
        int moved = Shift(map, unit, start, count, out int position);
        MakeSpan(map, position, Math.Max(position, end));
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
        MakeSpan(map, Math.Min(start, position), position);
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
        MakeSpan(map, position, position);
        return moved;
    }

    /// <summary>
    /// Widens the range to whole units, as <c>Range.Expand</c> in Word: to the start of the first unit it touches and the end of the last. A
    /// collapsed range grows to the unit it is in.
    /// </summary>
    public void Expand(TextUnit unit)
    {
        //Words, sentences and characters of a range inside one paragraph are all inside that paragraph.
        GetScoped(unit == TextUnit.Paragraph, out StoryMap map, out int start, out int end);
        List<int> boundaries = TextUnits.GetBoundaries(map.ToCharacters(), unit);
        int newStart = boundaries.LastOrDefault(b => b <= start);
        int newEnd = end > start ? boundaries.FirstOrDefault(b => b >= end) : boundaries.FirstOrDefault(b => b > start);
        if (newEnd < end)
        {
            //There is no boundary after the position: it is at the end of the story.
            newEnd = map.Length;
        }
        MakeSpan(map, newStart, Math.Max(newEnd, newStart));
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

    /// <summary>
    /// Makes the range the span from <paramref name="start"/> to <paramref name="end"/> in the map, which can be the map of the whole story or
    /// of one paragraph in it.
    /// </summary>
    internal void MakeSpan(StoryMap map, int start, int end)
    {
        OpenXmlElement story = StoryMap.FindRoot(map.Root);
        Anchor startAnchor = Anchor.At(map, start);
        Anchor endAnchor = end == start ? startAnchor.Clone() : Anchor.At(map, end);

        bool register = !_isSpan || !ReferenceEquals(_root, story);
        _isSpan = true;
        _root = story;
        _startAnchor = startAnchor;
        _endAnchor = endAnchor;
        _startOpenXmlElement = null;
        _endOpenXmlElement = null;
        if (ReferenceEquals(map.Root, story))
        {
            _lastStart = start;
            _lastEnd = end;
        }
        if (register)
        {
            SpanRegistry.Register(story, this);
        }
    }

    /// <summary>
    /// A collapsed span at an anchor in a story. It follows the text like any other span, which makes it a way to remember a place.
    /// </summary>
    internal static RangeWrapper CreateAt(OpenXmlElement story, Anchor anchor)
    {
        RangeWrapper range = new RangeWrapper(story);
        range._isSpan = true;
        range._root = story;
        range._startAnchor = anchor.Clone();
        range._endAnchor = anchor.Clone();
        range._startOpenXmlElement = null;
        range._endOpenXmlElement = null;
        SpanRegistry.Register(story, range);
        return range;
    }

    /// <summary>
    /// A new range over the span of characters from <paramref name="start"/> to <paramref name="end"/> in the map.
    /// </summary>
    internal static RangeWrapper FromMap(StoryMap map, int start, int end)
    {
        return new RangeWrapper(map, start, end);
    }

    /// <summary>
    /// Builds the map of the story of the range and finds the positions of the start and end of the range in it.
    /// </summary>
    internal void GetPositions(out StoryMap map, out int start, out int end)
    {
        if (_isSpan)
        {
            map = StoryMap.Build(_root);
            start = ResolveStory(map, _startAnchor, _lastStart);
            end = Math.Max(start, ResolveStory(map, _endAnchor, _lastEnd));
            _lastStart = start;
            _lastEnd = end;
            return;
        }

        map = StoryMap.Build(StoryMap.FindRoot(_startOpenXmlElement));
        ElementPositions(map, out start, out end);
    }

    /// <summary>
    /// Works out the positions of the range in a map. For a range of elements, with a map that covers it; for a span with the map of the
    /// paragraph it lies in, or of the whole story, and the positions are those of that map.
    /// </summary>
    /// <param name="forceStory">True to use the map of the whole story even if the range is inside one paragraph.</param>
    /// <param name="map">The map the positions are in.</param>
    /// <param name="start">The position of the start of the range in the map.</param>
    /// <param name="end">The position of the end of the range in the map.</param>
    internal void GetScoped(bool forceStory, out StoryMap map, out int start, out int end)
    {
        if (!forceStory && TryGetLocal(out map, out start, out end))
        {
            return;
        }
        GetPositions(out map, out start, out end);
    }

    private bool TryGetLocal(out StoryMap map, out int start, out int end)
    {
        map = null;
        start = 0;
        end = 0;

        Paragraph paragraph;
        if (_isSpan)
        {
            paragraph = ParagraphOfAnchor(_startAnchor);
            if (paragraph == null || !ReferenceEquals(paragraph, ParagraphOfAnchor(_endAnchor)) || !ReferenceEquals(StoryMap.FindRoot(paragraph), _root))
            {
                return false;
            }
        }
        else if (_startOpenXmlElement is Paragraph && ReferenceEquals(_startOpenXmlElement, _endOpenXmlElement))
        {
            paragraph = (Paragraph)_startOpenXmlElement;
        }
        else if (_startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart)
        {
            paragraph = _startOpenXmlElement.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph == null || !ReferenceEquals(paragraph, _endOpenXmlElement.Ancestors<Paragraph>().FirstOrDefault()))
            {
                return false;
            }
        }
        else
        {
            return false;
        }

        StoryMap local = StoryMap.Build(paragraph);
        if (_isSpan)
        {
            start = ResolveLocal(local, _startAnchor);
            end = ResolveLocal(local, _endAnchor);
            if (start < 0 || end < start)
            {
                return false;
            }
        }
        else
        {
            ElementPositions(local, out start, out end);
        }
        map = local;
        return true;
    }

    private static Paragraph ParagraphOfAnchor(Anchor anchor)
    {
        if (anchor.Element == null || anchor.Element is TableRow)
        {
            return null;
        }
        return anchor.Element as Paragraph ?? anchor.Element.Ancestors<Paragraph>().FirstOrDefault();
    }

    private static int ResolveStory(StoryMap map, Anchor anchor, int fallback)
    {
        if (anchor.Element == null)
        {
            return map.Length;
        }
        int position = ResolveLocal(map, anchor);
        return position >= 0 ? position : Math.Min(fallback, map.Length);
    }

    internal static int ResolveLocal(StoryMap map, Anchor anchor)
    {
        if (anchor.Element != null && map.TryGetAtom(anchor.Element, out Atom atom))
        {
            return atom.Start + Math.Min(anchor.Offset, atom.Length);
        }
        return -1;
    }

    /// <summary>
    /// The positions of a range of elements: from the first character at or after its start element to the end of the last character in or
    /// before its end element.
    /// </summary>
    private void ElementPositions(StoryMap map, out int start, out int end)
    {
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
        List<Atom> atoms = map.AtomsIn(start, end).ToList();
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
        GetScoped(false, out StoryMap map, out start, out end);
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
        return map.AtomsIn(start, end)
            .Where(a => !a.IsMark)
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
        GetScoped(false, out StoryMap map, out int start, out int end);
        return map.AtomsIn(start, end)
            .Where(a => a.Kind == AtomKind.ParagraphMark)
            .Select(a => (Paragraph)a.Element)
            .ToList();
    }

    /// <summary>
    /// The paragraphs that the range touches at all, which is where paragraph formatting applies. A collapsed range touches the paragraph it is in.
    /// </summary>
    internal IList<Paragraph> GetTouchedParagraphs()
    {
        GetScoped(false, out StoryMap map, out int start, out int end);
        List<Paragraph> paragraphs = new List<Paragraph>();
        if (map.Atoms.Count == 0)
        {
            return paragraphs;
        }

        int first = map.IndexAt(start) < 0 ? map.Atoms.Count - 1 : map.IndexAt(start);
        for (int i = first; i < map.Atoms.Count; i++)
        {
            Atom atom = map.Atoms[i];
            if (atom.Start >= end && !(start == end && i == first))
            {
                break;
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
