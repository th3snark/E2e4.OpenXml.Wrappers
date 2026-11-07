using System;
using System.Collections.Generic;
using System.Linq;

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
        return cells.Where(c => ReferenceEquals(c.TableCell.Parent.Parent, table)).ToList();
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
