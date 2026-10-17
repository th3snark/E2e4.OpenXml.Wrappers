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

    //The two anchors and the neighbours they remember, which a change of the document has to correct as well.
    internal IEnumerable<Anchor> AllAnchors()
    {
        if (_isSpan)
        {
            yield return _startAnchor;
            foreach (Anchor neighbour in _startAnchor.Neighbours)
            {
                yield return neighbour;
            }
            yield return _endAnchor;
            foreach (Anchor neighbour in _endAnchor.Neighbours)
            {
                yield return neighbour;
            }
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
    /// by counting through the story up to the range (nothing is built, but it is as far as the range from the start of the story), so keep
    /// a range (or its words) rather than calling this over and over on a very large document.
    /// </remarks>
    public int Start
    {
        get
        {
            if (_isSpan && StoryMap.TryPositionOf(_root, _startAnchor, out int position))
            {
                return position;
            }
            GetStoryPositions(out int start, out _);
            return start;
        }
        set
        {
            GetStoryPositions(out _, out int end);
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
            GetStoryPositions(out _, out int end);
            return end;
        }
        set
        {
            GetStoryPositions(out int start, out _);
            SetRange(Math.Min(start, value), value);
        }
    }

    /// <summary>
    /// The number of character positions in the story the range is in.
    /// </summary>
    public int StoryLength
    {
        get { return StoryMap.CountLength(StoryMap.FindRoot(AnchorElement)); }
    }

    /// <summary>
    /// Makes the range the span of characters from <paramref name="start"/> up to <paramref name="end"/> in the same story, as
    /// <c>Range.SetRange</c> in Word. The range is no longer made of elements; text formatted or replaced through it is split out of the
    /// runs that hold it.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The positions are not in order, or are beyond the end of the story.</exception>
    public void SetRange(int start, int end)
    {
        OpenXmlElement story = StoryMap.FindRoot(AnchorElement);
        if (start < 0 || end < start || !StoryMap.TryAnchorsFor(story, start, end, out Anchor startAnchor, out Anchor endAnchor))
        {
            throw new ArgumentOutOfRangeException(nameof(end), String.Format(CultureInfo.InvariantCulture, "The range {0} to {1} is not inside the story, which is {2} positions long.", start, end, StoryMap.CountLength(story)));
        }

        bool register = !_isSpan || !ReferenceEquals(_root, story);
        _isSpan = true;
        _root = story;
        _startAnchor = startAnchor;
        _endAnchor = endAnchor;
        _startOpenXmlElement = null;
        _endOpenXmlElement = null;
        _lastStart = start;
        _lastEnd = end;
        if (register)
        {
            SpanRegistry.Register(story, this);
        }
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
        if (TryShiftLocal(_startAnchor, unit, count, out Anchor moved, out int units))
        {
            //The start may have passed the end, which collapses the range there.
            Anchor end = _endAnchor;
            Paragraph oldParagraph = ParagraphOfAnchor(_startAnchor);
            if (ReferenceEquals(ParagraphOfAnchor(end), oldParagraph))
            {
                bool passed = !ReferenceEquals(ParagraphOfAnchor(moved), oldParagraph);
                if (!passed)
                {
                    StoryMap map = StoryMap.Build(oldParagraph);
                    passed = ResolveLocal(map, moved) > ResolveLocal(map, end);
                }
                if (passed)
                {
                    end = moved.Clone();
                }
            }
            MakeSpanAnchors(moved, end);
            return units;
        }

        GetPositions(out StoryMap whole, out int start, out int wholeEnd);
        int unitsMoved = Shift(whole, unit, start, count, out int position);
        MakeSpan(whole, position, Math.Max(position, wholeEnd));
        return unitsMoved;
    }

    /// <summary>
    /// Moves the end of the range by <paramref name="count"/> units (backwards when negative). If it passes the start the range is
    /// collapsed there. Returns the number of units it moved.
    /// </summary>
    public int MoveEnd(TextUnit unit, int count = 1)
    {
        if (TryShiftLocal(_endAnchor, unit, count, out Anchor moved, out int units))
        {
            //The end may have passed the start, which collapses the range there.
            Anchor start = _startAnchor;
            Paragraph oldParagraph = ParagraphOfAnchor(_endAnchor);
            if (ReferenceEquals(ParagraphOfAnchor(start), oldParagraph) && ReferenceEquals(ParagraphOfAnchor(moved), oldParagraph))
            {
                StoryMap map = StoryMap.Build(oldParagraph);
                if (ResolveLocal(map, moved) < ResolveLocal(map, start))
                {
                    start = moved.Clone();
                }
            }
            MakeSpanAnchors(start, moved);
            return units;
        }

        GetPositions(out StoryMap whole, out int wholeStart, out int end);
        int unitsMoved = Shift(whole, unit, end, count, out int position);
        MakeSpan(whole, Math.Min(wholeStart, position), position);
        return unitsMoved;
    }
    /// <summary>
    /// Collapses the range to its end and then moves that point by <paramref name="count"/> units, as <c>Range.Move</c> in Word. Returns the
    /// number of units it moved.
    /// </summary>
    public int Move(TextUnit unit, int count = 1)
    {
        if (TryShiftLocal(_isSpan ? _endAnchor : EndAnchorOfElements(), unit, count, out Anchor moved, out int units))
        {
            MakeSpanAnchors(moved, moved.Clone());
            return units;
        }

        GetPositions(out StoryMap whole, out _, out int end);
        int count1 = Shift(whole, unit, end, count, out int position);
        MakeSpan(whole, position, position);
        return count1;
    }

    //The end of a range of elements as an anchor, found in the map of the part of the story it is in.
    private Anchor EndAnchorOfElements()
    {
        GetScoped(false, out StoryMap map, out _, out int end);
        return Anchor.At(map, end);
    }

    /// <summary>
    /// Moves an anchor by units within its own paragraph. False when that cannot be done there (the anchor is not in a paragraph, or the
    /// move would go past the start of the paragraph or past its end), and a map of the story is needed.
    /// </summary>
    private bool TryShiftLocal(Anchor anchor, TextUnit unit, int count, out Anchor result, out int moved)
    {
        result = null;
        moved = 0;
        Paragraph paragraph = anchor == null ? null : ParagraphOfAnchor(anchor);
        if (paragraph == null || !ReferenceEquals(StoryMap.FindRoot(paragraph), StoryMap.FindRoot(AnchorElement)))
        {
            return false;
        }

        StoryMap map = StoryMap.Build(paragraph);
        int position = ResolveLocal(map, anchor);
        if (position < 0 || unit == TextUnit.Paragraph)
        {
            return false;
        }

        int units = Shift(map, unit, position, count, out int target);
        if (Math.Abs(units) < Math.Abs(count))
        {
            return false;
        }
        result = Anchor.At(map, target);
        moved = units;
        return true;
    }

    /// <summary>
    /// Widens the range to whole units, as <c>Range.Expand</c> in Word: to the start of the first unit it touches and the end of the last. A
    /// collapsed range grows to the unit it is in.
    /// </summary>
    public void Expand(TextUnit unit)
    {
        //Words, sentences and characters of a range inside one paragraph are all inside that paragraph.
        GetScoped(false, out StoryMap map, out int start, out int end);
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
    /// Makes the range the span between two anchors in the story of the range.
    /// </summary>
    internal void MakeSpanAnchors(Anchor startAnchor, Anchor endAnchor)
    {
        OpenXmlElement story = _isSpan ? _root : StoryMap.FindRoot(_startOpenXmlElement);
        bool register = !_isSpan;
        _isSpan = true;
        _root = story;
        _startAnchor = startAnchor;
        _endAnchor = endAnchor;
        _startOpenXmlElement = null;
        _endOpenXmlElement = null;
        if (register)
        {
            SpanRegistry.Register(story, this);
        }
    }

    /// <summary>
    /// Makes the range the span from <paramref name="start"/> to <paramref name="end"/> in the map, which can be the map of the whole story or
    /// of part of it.
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
    /// The positions of the start and end of the range in the story, found by counting through it without building a map.
    /// </summary>
    internal void GetStoryPositions(out int start, out int end)
    {
        if (_isSpan)
        {
            if (StoryMap.TryPositionOf(_root, _startAnchor, out start) && StoryMap.TryPositionOf(_root, _endAnchor, out end))
            {
                end = Math.Max(start, end);
                return;
            }
        }
        else if (TryGetLocal(out StoryMap local, out int localStart, out int localEnd))
        {
            int offset = StoryMap.OffsetOf(local.Root);
            start = offset + localStart;
            end = offset + localEnd;
            return;
        }

        //The text the range was on has gone, or it is a range of elements that is not inside one paragraph.
        GetPositions(out _, out start, out end);
    }

    /// <summary>
    /// Builds the map of the story of the range and finds the positions of the start and end of the range in it.
    /// </summary>
    internal void GetPositions(out StoryMap map, out int start, out int end)
    {
        if (_isSpan)
        {
            map = StoryMap.Build(_root);
            start = ResolveStory(map, ref _startAnchor, _lastStart);
            end = Math.Max(start, ResolveStory(map, ref _endAnchor, _lastEnd));
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
        if (!forceStory && (TryGetLocal(out map, out start, out end) || TryGetWindow(out map, out start, out end)))
        {
            return;
        }
        GetPositions(out map, out start, out end);
    }

    //A span from one paragraph to a later one that are among the same run of paragraphs (the paragraphs of a body or table cell with no table
    //between them) only needs a map of those paragraphs.
    private bool TryGetWindow(out StoryMap map, out int start, out int end)
    {
        map = null;
        start = 0;
        end = 0;
        if (!_isSpan)
        {
            return false;
        }

        Paragraph first = ParagraphOfAnchor(_startAnchor);
        Paragraph last = ParagraphOfAnchor(_endAnchor);
        if (first == null || last == null || ReferenceEquals(first, last) || !ReferenceEquals(first.Parent, last.Parent) || !ReferenceEquals(StoryMap.FindRoot(first), _root))
        {
            return false;
        }

        List<OpenXmlElement> blocks = new List<OpenXmlElement>();
        for (OpenXmlElement block = first; block != null; block = block.NextSibling())
        {
            if (!(block is Paragraph))
            {
                return false;
            }
            blocks.Add(block);
            if (ReferenceEquals(block, last))
            {
                break;
            }
        }
        if (!ReferenceEquals(blocks[blocks.Count - 1], last))
        {
            return false;
        }

        StoryMap window = StoryMap.BuildWindow(first.Parent, blocks);
        start = ResolveLocal(window, _startAnchor);
        end = ResolveLocal(window, _endAnchor);
        if (start < 0 || end < start)
        {
            return false;
        }
        _startAnchor.Refresh(window);
        _endAnchor.Refresh(window);
        map = window;
        return true;
    }

    private bool TryGetLocal(out StoryMap map, out int start, out int end)
    {
        map = null;
        start = 0;
        end = 0;

        OpenXmlElement scope;
        if (_isSpan)
        {
            Paragraph paragraph = ParagraphOfAnchor(_startAnchor);
            if (paragraph == null || !ReferenceEquals(paragraph, ParagraphOfAnchor(_endAnchor)) || !ReferenceEquals(StoryMap.FindRoot(paragraph), _root))
            {
                return false;
            }
            scope = paragraph;
        }
        else
        {
            scope = ElementScope();
            if (scope == null)
            {
                return false;
            }
        }

        StoryMap local = StoryMap.Build(scope);
        if (_isSpan)
        {
            start = ResolveLocal(local, _startAnchor);
            end = ResolveLocal(local, _endAnchor);
            if (start < 0 || end < start)
            {
                return false;
            }
            _startAnchor.Refresh(local);
            _endAnchor.Refresh(local);
        }
        else
        {
            ElementPositions(local, out start, out end);
        }
        map = local;
        return true;
    }

    /// <summary>
    /// The smallest part of the story that holds a range of elements and has a map of its own: the paragraph, table cell or table, or the
    /// table of a table row; the paragraph that holds a bookmark or comment range. Null for ranges that need the whole story.
    /// </summary>
    private OpenXmlElement ElementScope()
    {
        if (_startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart)
        {
            Paragraph paragraph = _startOpenXmlElement.Ancestors<Paragraph>().FirstOrDefault();
            return paragraph != null && ReferenceEquals(paragraph, _endOpenXmlElement.Ancestors<Paragraph>().FirstOrDefault()) ? paragraph : null;
        }
        if (!ReferenceEquals(_startOpenXmlElement, _endOpenXmlElement))
        {
            return null;
        }
        if (_startOpenXmlElement is TableRow row)
        {
            return row.Parent as Table;
        }
        return _startOpenXmlElement is Paragraph || _startOpenXmlElement is TableCell || _startOpenXmlElement is Table ? _startOpenXmlElement : null;
    }

    private static Paragraph ParagraphOfAnchor(Anchor anchor)
    {
        if (anchor.Element == null || anchor.Element is TableRow)
        {
            return null;
        }
        return anchor.Element as Paragraph ?? anchor.Element.Ancestors<Paragraph>().FirstOrDefault();
    }

    //The position of an anchor in the map of the whole story. If the text it was on has gone (without the library being told) the nearest
    //neighbour it remembers that is still there gives the place, and failing that the position it last had. The anchor is then moved to that
    //place so that it follows the text from there.
    private static int ResolveStory(StoryMap map, ref Anchor anchor, int fallback)
    {
        if (anchor.Element == null)
        {
            return map.Length;
        }
        int position = ResolveLocal(map, anchor);
        if (position >= 0)
        {
            anchor.Refresh(map);
            return position;
        }

        position = -1;
        foreach (Anchor neighbour in anchor.Neighbours)
        {
            if (map.TryGetAtom(neighbour.Element, out Atom atom))
            {
                position = atom.Start + Math.Min(neighbour.Offset, atom.Length);
                break;
            }
        }
        if (position < 0)
        {
            position = Math.Min(fallback, map.Length);
        }
        anchor = Anchor.At(map, position);
        return position;
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
