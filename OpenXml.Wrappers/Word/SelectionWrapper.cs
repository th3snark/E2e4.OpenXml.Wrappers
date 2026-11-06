using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a selection holds, as <c>WdSelectionType</c> in Word (with the same values; only the kinds that can be told here).
/// </summary>
public enum SelectionType
{
    /// <summary>An insertion point: nothing is selected.</summary>
    InsertionPoint = 1,

    /// <summary>A range of text.</summary>
    Normal = 2,

    /// <summary>A cell, or a column of cells.</summary>
    Column = 4,

    /// <summary>A row of a table.</summary>
    Row = 5,

    /// <summary>A single inline shape.</summary>
    InlineShape = 7
}

/// <summary>
/// The units <see cref="SelectionWrapper.HomeKey"/> and <see cref="SelectionWrapper.EndKey"/> move by, as <c>WdUnits</c> in Word (with the
/// same values; the line unit needs layout and is not offered).
/// </summary>
public enum KeyUnit
{
    /// <summary>The whole story.</summary>
    Story = 6,

    /// <summary>A column of a table.</summary>
    Column = 9,

    /// <summary>A row of a table.</summary>
    Row = 10
}

/// <summary>
/// The selection of a document, as <c>Selection</c> in Word: a place in a story (an insertion point or a range of text) that follows
/// the edits made to the story, and that typing, deleting and moving act on. There is no view here, so movement by lines, screens and
/// pages is not available.
/// </summary>
public partial class SelectionWrapper
{
    private readonly WordprocessingDocumentWrapper _document;
    private RangeWrapper _range;
    private bool _startIsActive;

    internal SelectionWrapper(WordprocessingDocumentWrapper document)
    {
        _document = document;
        _range = document.Content.Duplicate();
        _range.SetRange(0, 0);
    }

    //Makes the range the selection. A range of elements (a table, a row, a cell) becomes the span of the characters it covers.
    internal void Set(RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        if (range.IsSpan)
        {
            _range = range.Duplicate();
        }
        else
        {
            OpenXmlElement root = StoryMap.FindRoot(range.AnchorElement);
            StoryMap map = StoryMap.Build(root);
            range.GetStoryPositions(out int start, out int end);
            _range = RangeWrapper.FromMap(map, start, end);
        }
        _startIsActive = false;
    }

    /// <summary>The place the selection covers, as a range of its own, as <c>Selection.Range</c> in Word.</summary>
    public RangeWrapper Range
    {
        get { return _range.Duplicate(); }
    }

    /// <summary>The first position of the selection, as <c>Selection.Start</c> in Word.</summary>
    public int Start
    {
        get { return _range.Start; }
        set { SetRange(value, Math.Max(value, _range.End)); }
    }

    /// <summary>The position after the selection, as <c>Selection.End</c> in Word.</summary>
    public int End
    {
        get { return _range.End; }
        set { SetRange(Math.Min(value, _range.Start), value); }
    }

    /// <summary>Selects the characters between two positions, as <c>Selection.SetRange</c> in Word.</summary>
    public void SetRange(int start, int end)
    {
        _range.SetRange(start, end);
        _startIsActive = false;
    }

    /// <summary>
    /// The text selected, as <c>Selection.Text</c> in Word. An insertion point gives the character after it, as Word does.
    /// </summary>
    public string Text
    {
        get
        {
            if (_range.Start != _range.End)
            {
                return _range.Text;
            }
            RangeWrapper next = _range.Duplicate();
            int start = _range.Start;
            if (start >= _range.StoryLength)
            {
                return string.Empty;
            }
            next.SetRange(start, start + 1);
            return next.Text;
        }
    }

    /// <summary>What is selected, as <c>Selection.Type</c> in Word.</summary>
    public SelectionType Type
    {
        get
        {
            if (_range.Start == _range.End)
            {
                return SelectionType.InsertionPoint;
            }
            if (!_range.IsSpan)
            {
                return SelectionType.Normal;
            }
            //A selection that is exactly a row of a table, or exactly some of the cells of a table.
            RangeWrapper whole = new RangeWrapper(StoryMap.FindRoot(_range.AnchorElement));
            int start = _range.Start;
            int end = _range.End;
            List<RangeWrapper> rows = whole.Rows.Select(r => r.Range).ToList();
            if (rows.Any(r => r.Start == start) && rows.Any(r => r.End == end))
            {
                return SelectionType.Row;
            }
            List<RangeWrapper> cells = whole.Cells.Select(c => c.Range).ToList();
            if (cells.Any(c => c.Start == start) && cells.Any(c => c.End == end))
            {
                return SelectionType.Column;
            }
            if (_range.End - _range.Start == 1 && _range.InlineShapes.Count == 1)
            {
                return SelectionType.InlineShape;
            }
            return SelectionType.Normal;
        }
    }

    /// <summary>The kind of story the selection is in, as <c>Selection.StoryType</c> in Word.</summary>
    public StoryType StoryType
    {
        get { return _range.StoryType; }
    }

    /// <summary>True if the start is the end that moves when the selection is extended, as <c>Selection.StartIsActive</c> in Word.</summary>
    public bool StartIsActive
    {
        get { return _startIsActive; }
        set { _startIsActive = value; }
    }

    #region Collapsing and moving

    /// <summary>Makes the selection an insertion point at its start or end, as <c>Selection.Collapse</c> in Word.</summary>
    public void Collapse(CollapseDirection direction = CollapseDirection.Start)
    {
        _range.Collapse(direction);
        _startIsActive = false;
    }

    /// <summary>Extends the selection to the whole unit it is in, as <c>Selection.Expand</c> in Word; returns how much it grew.</summary>
    public int Expand(TextUnit unit = TextUnit.Word)
    {
        int before = _range.End - _range.Start;
        _range.Expand(unit);
        return _range.End - _range.Start - before;
    }

    /// <summary>Moves the start of the selection, as <c>Selection.MoveStart</c> in Word; returns how many units it moved.</summary>
    public int MoveStart(TextUnit unit = TextUnit.Character, int count = 1)
    {
        return _range.MoveStart(unit, count);
    }

    /// <summary>Moves the end of the selection, as <c>Selection.MoveEnd</c> in Word; returns how many units it moved.</summary>
    public int MoveEnd(TextUnit unit = TextUnit.Character, int count = 1)
    {
        return _range.MoveEnd(unit, count);
    }

    /// <summary>
    /// Collapses the selection to its end and moves it, as <c>Selection.Move</c> in Word, which counts the first unit as the collapse when
    /// something was selected. Returns how many units it moved.
    /// </summary>
    public int Move(TextUnit unit = TextUnit.Character, int count = 1)
    {
        return MovePoint(unit, count, Movement.Move);
    }

    private static void CheckUnit(TextUnit unit, bool allowParagraph)
    {
        if (!EnumCompat.IsDefined(unit) || (unit == TextUnit.Paragraph && !allowParagraph))
        {
            throw new ArgumentException("The unit cannot be used for this movement.", nameof(unit));
        }
    }

    /// <summary>
    /// Moves the insertion point to the right, or with <paramref name="extend"/> as <see cref="Movement.Extend"/> extends the selection
    /// by moving its active end, as <c>Selection.MoveRight</c> in Word. Returns how many units it moved.
    /// </summary>
    /// <exception cref="ArgumentException">The unit is a paragraph (use <see cref="MoveDown"/>).</exception>
    public int MoveRight(TextUnit unit = TextUnit.Character, int count = 1, Movement extend = Movement.Move)
    {
        CheckUnit(unit, false);
        return MovePoint(unit, count, Effective(extend));
    }

    /// <summary>
    /// Moves the insertion point to the left, or with <paramref name="extend"/> as <see cref="Movement.Extend"/> extends the selection
    /// by moving its active end, as <c>Selection.MoveLeft</c> in Word. Returns how many units it moved.
    /// </summary>
    /// <exception cref="ArgumentException">The unit is a paragraph (use <see cref="MoveUp"/>).</exception>
    public int MoveLeft(TextUnit unit = TextUnit.Character, int count = 1, Movement extend = Movement.Move)
    {
        CheckUnit(unit, false);
        return MovePoint(unit, -count, Effective(extend));
    }

    //Moves by characters, words or sentences. Without extending, a selection of something collapses first (to its end going right, to its
    //start going left) and that takes the place of the first unit of the move.
    private int MovePoint(TextUnit unit, int count, Movement extend)
    {
        if (!EnumCompat.IsDefined(extend))
        {
            throw new ArgumentOutOfRangeException(nameof(extend));
        }
        int start = _range.Start;
        int end = _range.End;
        if (extend == Movement.Move)
        {
            int from = count >= 0 ? end : start;
            int collapsing = 0;
            if (start != end && count != 0)
            {
                count += count > 0 ? -1 : 1;
                collapsing = 1;
            }
            int moved = MoveFrom(from, unit, count, out int to, true);
            _range.SetRange(to, to);
            _startIsActive = false;
            return moved + collapsing;
        }

        int active = _startIsActive ? start : end;
        int other = _startIsActive ? end : start;
        int result = MoveFrom(active, unit, count, out int target);
        if (target >= other)
        {
            _range.SetRange(other, target);
        }
        else
        {
            _range.SetRange(target, other);
            _startIsActive = !_startIsActive;
        }
        return result;
    }

    //Moves a point by units; returns how many units it moved (always positive) and where it got to.
    //An insertion point cannot go after the last paragraph mark of the story, so it stops before it.
    private int MoveFrom(int position, TextUnit unit, int count, out int target, bool insertionPoint = false)
    {
        int last = _range.StoryLength - 1;
        for (int attempt = count; ; attempt += attempt > 0 ? -1 : 1)
        {
            RangeWrapper point = _range.Duplicate();
            point.SetRange(position, position);
            int moved = Math.Abs(point.MoveEnd(unit, attempt));
            target = point.End;
            if (!insertionPoint || target <= last || attempt == 0)
            {
                return moved;
            }
        }
    }

    private IList<int> ParagraphStarts()
    {
        StoryMap map = StoryMap.Build(StoryMap.FindRoot(_range.AnchorElement));
        List<int> starts = new List<int> { 0 };
        foreach (Atom atom in map.Atoms)
        {
            if (atom.Kind == AtomKind.ParagraphMark)
            {
                starts.Add(atom.End);
            }
        }
        return starts;
    }

    /// <summary>
    /// Moves up by paragraphs, as <c>Selection.MoveUp</c> in Word: to the start of the paragraph, or of the one before when already at
    /// its start. Returns how many paragraphs it moved.
    /// </summary>
    /// <exception cref="NotSupportedException">The unit is not a paragraph: lines and screens need layout.</exception>
    public int MoveUp(TextUnit unit = TextUnit.Paragraph, int count = 1, Movement extend = Movement.Move)
    {
        return MoveParagraphs(unit, -count, Effective(extend));
    }

    /// <summary>
    /// Moves down by paragraphs, as <c>Selection.MoveDown</c> in Word: to the start of the next one, or the end of the story's last
    /// paragraph. Returns how many paragraphs it moved.
    /// </summary>
    /// <exception cref="NotSupportedException">The unit is not a paragraph: lines and screens need layout.</exception>
    public int MoveDown(TextUnit unit = TextUnit.Paragraph, int count = 1, Movement extend = Movement.Move)
    {
        return MoveParagraphs(unit, count, Effective(extend));
    }

    private int MoveParagraphs(TextUnit unit, int count, Movement extend)
    {
        if (unit != TextUnit.Paragraph)
        {
            throw new NotSupportedException("Only paragraphs can be moved by up and down: lines and screens need layout.");
        }
        if (!EnumCompat.IsDefined(extend))
        {
            throw new ArgumentOutOfRangeException(nameof(extend));
        }
        IList<int> starts = ParagraphStarts();
        int last = _range.StoryLength - 1;
        int active = extend == Movement.Move ? _range.End : _startIsActive ? _range.Start : _range.End;
        int position = active;
        int steps = 0;
        for (int i = 0; i < Math.Abs(count); i++)
        {
            int next;
            if (count < 0)
            {
                int current = starts.Where(s => s <= position).DefaultIfEmpty(0).Max();
                next = position > current ? current : starts.Where(s => s < current).DefaultIfEmpty(current).Max();
            }
            else
            {
                next = starts.Where(s => s > position).DefaultIfEmpty(last).Min();
                next = Math.Min(next, last);
            }
            if (next == position)
            {
                break;
            }
            position = next;
            steps++;
        }

        if (extend == Movement.Move)
        {
            _range.SetRange(position, position);
            _startIsActive = false;
        }
        else
        {
            int other = _startIsActive ? _range.End : _range.Start;
            if (position >= other)
            {
                _range.SetRange(other, position);
            }
            else
            {
                _range.SetRange(position, other);
                _startIsActive = !_startIsActive;
            }
        }
        return steps;
    }

    //The cell the start of the selection is in, or null outside a table.
    private TableCellWrapper CurrentCell()
    {
        int start = _range.Start;
        return new RangeWrapper(StoryMap.FindRoot(_range.AnchorElement)).Cells.LastOrDefault(c => c.Range.Start <= start && start < c.Range.End);
    }

    /// <summary>
    /// Goes to the start of the story, of the row or of the column of a table, or with <paramref name="extend"/> as
    /// <see cref="Movement.Extend"/> extends the selection to there, as <c>Selection.HomeKey</c> in Word. Returns how many characters it moved.
    /// </summary>
    /// <exception cref="InvalidOperationException">The unit is a row or column and the selection is not in a table.</exception>
    public int HomeKey(KeyUnit unit = KeyUnit.Story, Movement extend = Movement.Move)
    {
        return GoTo(unit, Effective(extend), true);
    }

    /// <summary>
    /// Goes to the end of the story (before its last paragraph mark), of the row or of the column of a table, or extends the selection
    /// to there, as <c>Selection.EndKey</c> in Word. Returns how many characters it moved.
    /// </summary>
    /// <exception cref="InvalidOperationException">The unit is a row or column and the selection is not in a table.</exception>
    public int EndKey(KeyUnit unit = KeyUnit.Story, Movement extend = Movement.Move)
    {
        return GoTo(unit, Effective(extend), false);
    }

    private int GoTo(KeyUnit unit, Movement extend, bool home)
    {
        if (!EnumCompat.IsDefined(unit) || !EnumCompat.IsDefined(extend))
        {
            throw new ArgumentOutOfRangeException(nameof(unit));
        }
        int target;
        if (unit == KeyUnit.Story)
        {
            target = home ? 0 : _range.StoryLength - 1;
        }
        else
        {
            TableCellWrapper cell = CurrentCell() ?? throw new InvalidOperationException("The selection is not in a table.");
            IEnumerable<TableCellWrapper> choice = unit == KeyUnit.Row ? cell.Row.Cells : (IEnumerable<TableCellWrapper>)cell.Column.Cells;
            List<TableCellWrapper> line = choice.ToList();
            target = (home ? line[0] : line[line.Count - 1]).Range.Start;
        }

        int from = extend == Movement.Move ? _range.End : _startIsActive ? _range.Start : _range.End;
        if (extend == Movement.Move)
        {
            _range.SetRange(target, target);
            _startIsActive = false;
        }
        else
        {
            int other = _startIsActive ? _range.End : _range.Start;
            if (target >= other)
            {
                _range.SetRange(other, target);
            }
            else
            {
                _range.SetRange(target, other);
                _startIsActive = !_startIsActive;
            }
        }
        return Math.Abs(target - from);
    }

    #endregion

    #region Typing and editing

    /// <summary>
    /// Types text at the insertion point, replacing what is selected, and leaves the insertion point after it, as
    /// <c>Selection.TypeText</c> in Word.
    /// </summary>
    public void TypeText(string text)
    {
        Guard.ThrowIfNull(text);
        if (text.Length == 0)
        {
            return;
        }
        ExtendMode = false;
        if (_range.Start != _range.End)
        {
            DeleteSelected();
        }
        int position = _range.Start;
        RangeWrapper inserted = _range.Duplicate();
        inserted.InsertAfter(text);
        int after = Math.Max(inserted.End, position + text.Length);
        _range.SetRange(after, after);
        _startIsActive = false;
    }

    /// <summary>Types a paragraph mark, as <c>Selection.TypeParagraph</c> in Word.</summary>
    public void TypeParagraph()
    {
        TypeText("\r");
    }

    /// <summary>
    /// Deletes the selection, or the character before the insertion point, as <c>Selection.TypeBackspace</c> in Word.
    /// </summary>
    public void TypeBackspace()
    {
        ExtendMode = false;
        if (_range.Start != _range.End)
        {
            DeleteSelected();
            return;
        }
        int at = _range.Start;
        if (at == 0)
        {
            return;
        }
        RangeWrapper before = _range.Duplicate();
        before.SetRange(at - 1, at);
        before.Delete();
        _range.SetRange(at - 1, at - 1);
    }

    private void DeleteSelected()
    {
        int start = _range.Start;
        RangeWrapper selected = _range.Duplicate();
        selected.Delete();
        _range.SetRange(start, start);
        _startIsActive = false;
    }

    /// <summary>
    /// Deletes the selection, or the units after the insertion point, as <c>Selection.Delete</c> in Word. Returns how many units it
    /// deleted (1 for a selection that was deleted).
    /// </summary>
    public int Delete(TextUnit unit = TextUnit.Character, int count = 1)
    {
        if (!EnumCompat.IsDefined(unit))
        {
            throw new ArgumentOutOfRangeException(nameof(unit));
        }
        if (_range.Start != _range.End)
        {
            DeleteSelected();
            return 1;
        }
        if (count <= 0)
        {
            return 0;
        }
        RangeWrapper ahead = _range.Duplicate();
        int start = _range.Start;
        if (start >= _range.StoryLength - 1)
        {
            return 0;
        }
        ahead.SetRange(start, start);
        int moved = ahead.MoveEnd(unit, count);
        if (ahead.End == ahead.Start)
        {
            return 0;
        }
        ahead.Delete();
        _range.SetRange(start, start);
        return Math.Abs(moved);
    }

    /// <summary>Inserts text after the selection and selects it too, as <c>Selection.InsertAfter</c> in Word.</summary>
    public void InsertAfter(string text)
    {
        _range.InsertAfter(text);
    }

    /// <summary>Inserts text before the selection and selects it too, as <c>Selection.InsertBefore</c> in Word.</summary>
    public void InsertBefore(string text)
    {
        _range.InsertBefore(text);
    }

    #endregion

    #region Tables

    /// <summary>Selects the row of the table the selection is in, as <c>Selection.SelectRow</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    public void SelectRow()
    {
        TableCellWrapper cell = CurrentCell() ?? throw new InvalidOperationException("The selection is not in a table.");
        Set(cell.Row.Range);
    }

    /// <summary>Selects the cell the selection is in, as <c>Selection.SelectCell</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    public void SelectCell()
    {
        Set((CurrentCell() ?? throw new InvalidOperationException("The selection is not in a table.")).Range);
    }

    #endregion

    #region What the selection holds

    /// <summary>The font of the selection, as <c>Selection.Font</c> in Word.</summary>
    public FontWrapper Font
    {
        get { return _range.Font; }
    }

    /// <summary>The paragraph formatting of the selection, as <c>Selection.ParagraphFormat</c> in Word.</summary>
    public ParagraphFormatWrapper ParagraphFormat
    {
        get { return _range.ParagraphFormat; }
    }

    /// <summary>The style of the selection, as <c>Selection.Style</c> in Word.</summary>
    public string Style
    {
        get { return _range.Style; }
        set { _range.Style = value; }
    }

    /// <summary>The paragraphs of the selection, as <c>Selection.Paragraphs</c> in Word.</summary>
    public ParagraphWrapperCollection Paragraphs
    {
        get { return _range.Paragraphs; }
    }

    /// <summary>The tables in the selection, as <c>Selection.Tables</c> in Word.</summary>
    public TableWrapperCollection Tables
    {
        get { return _range.Tables; }
    }

    /// <summary>The cells in the selection, as <c>Selection.Cells</c> in Word.</summary>
    public TableCellWrapperCollection Cells
    {
        get { return _range.Cells; }
    }

    /// <summary>The bookmarks in the selection, as <c>Selection.Bookmarks</c> in Word.</summary>
    public BookmarkWrapperCollection Bookmarks
    {
        get { return _range.Bookmarks; }
    }

    /// <summary>The fields in the selection, as <c>Selection.Fields</c> in Word.</summary>
    public FieldWrapperCollection Fields
    {
        get { return _range.Fields; }
    }

    /// <summary>The hyperlinks in the selection, as <c>Selection.Hyperlinks</c> in Word.</summary>
    public HyperlinkWrapperCollection Hyperlinks
    {
        get { return _range.Hyperlinks; }
    }

    /// <summary>The inline shapes in the selection, as <c>Selection.InlineShapes</c> in Word.</summary>
    public InlineShapeWrapperCollection InlineShapes
    {
        get { return _range.InlineShapes; }
    }

    /// <summary>The content controls in the selection, as <c>Selection.ContentControls</c> in Word.</summary>
    public ContentControlWrapperCollection ContentControls
    {
        get { return _range.ContentControls; }
    }

    /// <summary>The document the selection is in, as <c>Selection.Document</c> in Word.</summary>
    public WordprocessingDocumentWrapper Document
    {
        get { return _document; }
    }

    #endregion
}
