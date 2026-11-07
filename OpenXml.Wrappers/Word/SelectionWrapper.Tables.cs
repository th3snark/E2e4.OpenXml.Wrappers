using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class SelectionWrapper
{
    //The cells of a table that the selection touches: the one the insertion point is in, or those the selection overlaps. Those of the
    //table the selection starts in only.
    private List<TableCellWrapper> SelectedCells()
    {
        List<TableCellWrapper> cells = new RangeWrapper(StoryMap.FindRoot(_range.AnchorElement)).Cells
            .Where(c => _range.Start == _range.End ? c.Range.Start <= _range.Start && _range.Start < c.Range.End : c.Range.Start < _range.End && _range.Start < c.Range.End)
            .ToList();
        if (cells.Count == 0)
        {
            throw new InvalidOperationException("The selection is not in a table.");
        }
        object table = cells[0].TableCell.Parent.Parent;
        cells = cells.Where(c => ReferenceEquals(c.TableCell.Parent.Parent, table)).ToList();
        if (cells.Count < 2)
        {
            return cells;
        }

        //A selection that runs from one cell to another is the block of cells between them, as in Word.
        TableCell first = cells[0].TableCell;
        TableCell last = cells[cells.Count - 1].TableCell;
        int left = Math.Min(TableCellWrapper.GetGridStart(first), TableCellWrapper.GetGridStart(last));
        int right = Math.Max(TableCellWrapper.GetGridStart(first) + TableCellWrapper.Span(first), TableCellWrapper.GetGridStart(last) + TableCellWrapper.Span(last));
        return cells.Where(c => TableCellWrapper.GetGridStart(c.TableCell) < right && left < TableCellWrapper.GetGridStart(c.TableCell) + TableCellWrapper.Span(c.TableCell)).ToList();
    }

    //Selects what lies between the start of the first cell or row and the end of the last.
    private void SelectBetween(RangeWrapper first, RangeWrapper last)
    {
        SetRange(first.Start, last.End);
    }

    /// <summary>
    /// Splits the table in two before the row the insertion point or the start of the selection is in, as <c>Selection.SplitTable</c> in
    /// Word: that row becomes the first of the new table, and a paragraph is left between the two (before the table for its first row), and the insertion point is in that paragraph.
    /// </summary>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    public void SplitTable()
    {
        TableRowWrapper row = SelectedCells()[0].Row;
        Table table = (Table)row.TableRow.Parent;
        Paragraph separator;
        if (row.TableRow.PreviousSibling<TableRow>() == null)
        {
            //Before the first row Word puts an empty paragraph before the table.
            separator = new Paragraph();
            table.InsertBeforeSelf(separator);
        }
        else
        {
            separator = (Paragraph)new TableWrapper(table).Split(row).Table.PreviousSibling();
        }
        int at = new ParagraphWrapper(separator).Range.Start;
        SetRange(at, at);
    }

    /// <summary>
    /// Replaces the selection with a formula field, as <c>Selection.InsertFormula</c> in Word, and puts the insertion point after it. See
    /// <see cref="TableCellWrapper.Formula"/> for the results that are worked out here. Without a formula, in a cell, it is the sum of the
    /// numbers above or to the left; elsewhere the field has no formula, so that it shows an error, as in Word.
    /// </summary>
    /// <param name="formula">The formula, with or without the equals sign, or null.</param>
    /// <param name="numberFormat">A numeric picture such as <c>0.00</c> for the result, or null.</param>
    public void InsertFormula(string formula = null, string numberFormat = null)
    {
        string code;
        if (!string.IsNullOrWhiteSpace(formula))
        {
            code = formula.Trim()[0] == '=' ? formula.Trim() : "=" + formula.Trim();
        }
        else
        {
            TableCellWrapper cell = CurrentCell();
            code = cell == null ? "=" : TableFormulas.AutoSumCode(cell.TableCell);
        }
        if (!string.IsNullOrEmpty(numberFormat))
        {
            code += " \\# \"" + numberFormat + "\"";
        }
        ExtendMode = false;
        if (_range.Start != _range.End)
        {
            DeleteSelected();
        }
        RangeWrapper at = _range.Duplicate();
        FieldWrapper field = at.Fields.Add(at, code);
        int after = field.Range.End;
        _range.SetRange(after, after);
        _startIsActive = false;
    }
    /// <summary>
    /// Inserts cells at the selected cells, as <c>Selection.InsertCells</c> in Word, and selects the new ones. With
    /// <see cref="InsertCellsShift.Right"/> a cell is inserted in front of the selected ones in each row, which makes the row longer than
    /// the others; with <see cref="InsertCellsShift.Down"/> the content of the cells below moves down and a row is added at the bottom
    /// for what no longer fits; the other two insert as many rows or columns as the selection touches.
    /// </summary>
    /// <param name="shiftCells">Where the cells that are in the way go.</param>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    /// <exception cref="NotSupportedException">A shift down in a column with merged or missing cells.</exception>
    public void InsertCells(InsertCellsShift shiftCells = InsertCellsShift.Right)
    {
        if (!EnumCompat.IsDefined(shiftCells))
        {
            throw new ArgumentOutOfRangeException(nameof(shiftCells));
        }
        List<TableCellWrapper> cells = SelectedCells();
        if (shiftCells == InsertCellsShift.EntireRow)
        {
            InsertRowsAbove(cells.Select(c => c.Row).Distinct().Count());
        }
        else if (shiftCells == InsertCellsShift.EntireColumn)
        {
            InsertColumns();
        }
        else if (shiftCells == InsertCellsShift.Right)
        {
            InsertCellsRight(cells);
        }
        else
        {
            InsertCellsDown(cells);
        }
    }

    private void InsertCellsRight(List<TableCellWrapper> cells)
    {
        List<TableCell> added = new List<TableCell>();
        foreach (IGrouping<TableRow, TableCellWrapper> row in cells.GroupBy(c => c.Row.TableRow))
        {
            TableCell first = row.First().TableCell;
            for (int i = 0; i < row.Count(); i++)
            {
                TableCell cell = (TableCell)first.Clone();
                cell.RemoveAnchors();
                cell.ClearText();
                cell.SetGridSpan(1);
                first.InsertBeforeSelf(cell);
                added.Add(cell);
            }
        }
        MakeRowsRagged((Table)cells[0].TableCell.Parent.Parent);
        SelectBetween(new TableCellWrapper(added[0]).Range, new TableCellWrapper(added[added.Count - 1]).Range);
    }

    //A row with more cells than the table has columns gets a column of its own in the grid, and the rows that stay shorter say that the
    //columns after their last cell are not theirs (gridAfter), as Word writes it.
    private static void MakeRowsRagged(Table table)
    {
        TableGrid grid = table.GetFirstChild<TableGrid>();
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        int longest = rows.Max(r => r.Elements<TableCell>().Sum(c => TableCellWrapper.Span(c)));
        List<GridColumn> columns = grid.Elements<GridColumn>().ToList();
        for (int i = columns.Count; i < longest; i++)
        {
            GridColumn clone = (GridColumn)columns[columns.Count - 1].Clone();
            grid.Append(clone);
            columns.Add(clone);
        }
        foreach (TableRow row in rows)
        {
            int missing = longest - row.Elements<TableCell>().Sum(c => TableCellWrapper.Span(c));
            if (missing <= 0)
            {
                continue;
            }
            TableRowProperties properties = row.TableRowProperties;
            if (properties == null)
            {
                properties = new TableRowProperties();
                TablePropertyExceptions exceptions = row.GetFirstChild<TablePropertyExceptions>();
                if (exceptions != null)
                {
                    exceptions.InsertAfterSelf(properties);
                }
                else
                {
                    row.PrependChild(properties);
                }
            }
            properties.RemoveAllChildren<GridAfter>();
            properties.RemoveAllChildren<WidthAfterTableRow>();
            int width = columns.Skip(longest - missing).Sum(c => int.TryParse(c.Width?.Value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int w) ? w : 0);
            properties.AppendChild(new GridAfter() { Val = missing });
            properties.AppendChild(new WidthAfterTableRow() { Width = width.ToString(System.Globalization.CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
        }
        TableWrapper.ColumnsChanged(table);
    }

    private void InsertCellsDown(List<TableCellWrapper> cells)
    {
        Table table = (Table)cells[0].TableCell.Parent.Parent;
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        List<TableCell> selected = cells.Select(c => c.TableCell).ToList();
        if (selected.Any(c => TableCellWrapper.Span(c) != 1 || c.TableCellProperties?.VerticalMerge != null))
        {
            throw new NotSupportedException("Cells cannot be inserted with a shift down where cells are merged.");
        }

        //For each column the number of selected cells and the row of the first.
        List<IGrouping<int, TableCell>> columns = selected.GroupBy(c => TableCellWrapper.GetGridStart(c)).ToList();
        int newRows = columns.Max(g => g.Count());
        TableRow last = rows[rows.Count - 1];
        for (int i = 0; i < newRows; i++)
        {
            TableRow row = (TableRow)last.Clone();
            row.RemoveAnchors();
            row.ClearText();
            rows[rows.Count - 1].InsertAfterSelf(row);
            rows.Add(row);
        }

        List<TableCell> inserted = new List<TableCell>();
        foreach (IGrouping<int, TableCell> column in columns)
        {
            int count = column.Count();
            int top = rows.IndexOf((TableRow)column.First().Parent);
            List<TableCell> line = rows.Select(r => r.Elements<TableCell>().FirstOrDefault(c => TableCellWrapper.GetGridStart(c) == column.Key)).ToList();
            if (line.Skip(top).Any(c => c == null || TableCellWrapper.Span(c) != 1 || c.TableCellProperties?.VerticalMerge != null))
            {
                throw new NotSupportedException("Cells cannot be inserted with a shift down where cells are merged or missing.");
            }
            for (int i = line.Count - 1; i >= top + count; i--)
            {
                MoveContent(line[i - count], line[i]);
            }
            for (int i = top; i < top + count; i++)
            {
                MoveContent(null, line[i]);
                inserted.Add(line[i]);
            }
        }
        inserted = inserted.OrderBy(c => rows.IndexOf((TableRow)c.Parent)).ThenBy(c => TableCellWrapper.GetGridStart(c)).ToList();
        SelectBetween(new TableCellWrapper(inserted[0]).Range, new TableCellWrapper(inserted[inserted.Count - 1]).Range);
    }

    //Replaces what a cell holds with what another holds (or with an empty paragraph).
    private static void MoveContent(TableCell from, TableCell to)
    {
        foreach (OpenXmlElement child in to.ChildElements.Where(e => !(e is TableCellProperties)).ToList())
        {
            child.Remove();
        }
        if (from == null)
        {
            to.AppendChild(new Paragraph());
            return;
        }
        foreach (OpenXmlElement child in from.ChildElements.Where(e => !(e is TableCellProperties)).ToList())
        {
            child.Remove();
            to.AppendChild(child);
        }
    }
    /// <summary>
    /// The header or footer the selection is in, as <c>Selection.HeaderFooter</c> in Word, or null when it is in another story. A header
    /// that several sections share is the one of the first section that has it.
    /// </summary>
    public HeaderFooterWrapper HeaderFooter
    {
        get
        {
            StoryType story = _range.StoryType;
            if (story != StoryType.PrimaryHeader && story != StoryType.PrimaryFooter && story != StoryType.FirstPageHeader
                && story != StoryType.FirstPageFooter && story != StoryType.EvenPagesHeader && story != StoryType.EvenPagesFooter)
            {
                return null;
            }
            object root = StoryMap.FindRoot(_range.AnchorElement);
            foreach (SectionWrapper section in _document.Sections)
            {
                foreach (HeaderFooterWrapper candidate in section.Headers.Concat(section.Footers))
                {
                    if (candidate.Exists && ReferenceEquals(StoryMap.FindRoot(candidate.Range.AnchorElement), root))
                    {
                        return candidate;
                    }
                }
            }
            return null;
        }
    }
    /// <summary>Selects the column of the table the selection is in, as <c>Selection.SelectColumn</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    public void SelectColumn()
    {
        TableCellWrapper cell = CurrentCell() ?? throw new InvalidOperationException("The selection is not in a table.");
        List<TableCellWrapper> column = cell.Column.Cells.ToList();
        SelectBetween(column[0].Range, column[column.Count - 1].Range);
    }

    /// <summary>
    /// Inserts rows above the first row of the selection, as <c>Selection.InsertRows</c> in Word, and selects them.
    /// </summary>
    /// <param name="numRows">How many rows to insert.</param>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="numRows"/> is less than one.</exception>
    public void InsertRows(int numRows)
    {
        InsertRowsAbove(numRows);
    }

    /// <summary>
    /// Inserts a column to the left of the selected column for each column the selection touches, as <c>Selection.InsertColumns</c> in
    /// Word, and selects the new columns. The new columns have the width of the first selected column, so the table can grow wider than
    /// the page, as it does in Word. In a row where a merged cell spans the place of the new columns, that cell grows instead.
    /// </summary>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    public void InsertColumns()
    {
        InsertColumnsNextTo(false);
    }

    /// <summary>
    /// Inserts a column to the right of the selected column for each column the selection touches, as <c>Selection.InsertColumnsRight</c>
    /// in Word, and selects the new columns. See <see cref="InsertColumns"/>.
    /// </summary>
    /// <exception cref="InvalidOperationException">The selection is not in a table.</exception>
    public void InsertColumnsRight()
    {
        InsertColumnsNextTo(true);
    }

    private void InsertColumnsNextTo(bool right)
    {
        List<TableCellWrapper> cells = SelectedCells();
        Table table = (Table)cells[0].TableCell.Parent.Parent;
        //The columns are those of the cell the selection starts in and the cell it ends in, and those between them.
        TableCell startCell = cells[0].TableCell;
        TableCell endCell = cells[cells.Count - 1].TableCell;
        int first = Math.Min(TableCellWrapper.GetGridStart(startCell), TableCellWrapper.GetGridStart(endCell));
        int end = Math.Max(TableCellWrapper.GetGridStart(startCell) + TableCellWrapper.Span(startCell), TableCellWrapper.GetGridStart(endCell) + TableCellWrapper.Span(endCell));
        int count = end - first;

        List<TableCell> added = TableColumnWrapperCollection.InsertColumns(table, right ? end : first, count, right ? end - 1 : first);
        if (added.Count > 0)
        {
            SelectBetween(new TableCellWrapper(added[0]).Range, new TableCellWrapper(added[added.Count - 1]).Range);
        }
    }
}
