using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class TableCellWrapper
{
    /// <summary>
    /// Merges the rectangle of cells between this cell and <paramref name="mergeTo"/> (either corner may come first) into one cell, as
    /// <c>Cell.Merge</c> in Word. The text of the other cells is added to the end of this rectangle's first cell.
    /// </summary>
    /// <exception cref="ArgumentException">The cells are not in the same table.</exception>
    /// <exception cref="NotSupportedException">The rectangle contains cells that are already vertically merged, or cuts through a merged cell.</exception>
    public void Merge(TableCellWrapper mergeTo)
    {
        ArgumentNullException.ThrowIfNull(mergeTo);
        Table table = _tableCell.Ancestors<Table>().FirstOrDefault();
        if (table == null || mergeTo._tableCell.Ancestors<Table>().FirstOrDefault() != table)
        {
            throw new ArgumentException("The cells must be in the same table.", nameof(mergeTo));
        }
        if (mergeTo._tableCell == _tableCell)
        {
            return;
        }

        List<TableRow> rows = table.Elements<TableRow>().ToList();
        int firstRow = Math.Min(rows.IndexOf((TableRow)_tableCell.Parent), rows.IndexOf((TableRow)mergeTo._tableCell.Parent));
        int lastRow = Math.Max(rows.IndexOf((TableRow)_tableCell.Parent), rows.IndexOf((TableRow)mergeTo._tableCell.Parent));
        GetGridRange(_tableCell, out int start1, out int end1);
        GetGridRange(mergeTo._tableCell, out int start2, out int end2);
        int firstColumn = Math.Min(start1, start2);
        int lastColumn = Math.Max(end1, end2);

        //Check the whole rectangle before changing anything.
        List<List<TableCell>> rectangle = new List<List<TableCell>>();
        for (int i = firstRow; i <= lastRow; i++)
        {
            List<TableCell> cells = new List<TableCell>();
            int covered = 0;
            foreach (TableCell cell in rows[i].Elements<TableCell>())
            {
                GetGridRange(cell, out int start, out int end);
                if (end < firstColumn || start > lastColumn)
                {
                    continue;
                }
                if (start < firstColumn || end > lastColumn)
                {
                    throw new NotSupportedException("The cells to merge cut through a cell that is already merged.");
                }
                if (cell.TableCellProperties?.VerticalMerge != null || cell.TableCellProperties?.HorizontalMerge != null)
                {
                    throw new NotSupportedException("Merging cells that are already vertically merged is not supported.");
                }
                cells.Add(cell);
                covered += end - start + 1;
            }
            if (covered != lastColumn - firstColumn + 1)
            {
                throw new NotSupportedException("The cells to merge are not all present in every row.");
            }
            rectangle.Add(cells);
        }

        //Merge each row of the rectangle sideways.
        List<TableCell> merged = new List<TableCell>();
        foreach (List<TableCell> cells in rectangle)
        {
            TableCell first = cells[0];
            if (cells.Count > 1)
            {
                int width = cells.Sum(c => new TableCellWrapper(c).CellWidth);
                int span = cells.Sum(Span);
                foreach (TableCell other in cells.Skip(1))
                {
                    MoveText(other, first);
                    other.Remove();
                }
                GetOrAddTableCellProperties(first).GridSpan = new GridSpan() { Val = span };
                new TableCellWrapper(first).CellWidth = width;
            }
            merged.Add(first);
        }

        //Then merge the rows down.
        if (merged.Count > 1)
        {
            GetOrAddTableCellProperties(merged[0]).VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
            foreach (TableCell below in merged.Skip(1))
            {
                MoveText(below, merged[0]);
                below.RemoveAllChildren<Paragraph>();
                below.AppendChild(new Paragraph());
                GetOrAddTableCellProperties(below).VerticalMerge = new VerticalMerge();
            }
        }
    }

    /// <summary>
    /// Splits the cell into <paramref name="numRows"/> rows and <paramref name="numColumns"/> columns, as <c>Cell.Split</c> in Word. The
    /// text stays in the top left cell. Splitting into columns divides the grid column of the cell, and the other cells of the table that
    /// span it grow to match; splitting into rows adds rows below and vertically merges the other cells of the row with them.
    /// </summary>
    /// <exception cref="NotSupportedException">The cell, or another cell of its row, is vertically merged, or the cell spans columns and
    /// the number of columns to split into does not divide the span.</exception>
    public void Split(int numRows, int numColumns)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(numRows, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(numColumns, 1);
        if (numRows == 1 && numColumns == 1)
        {
            return;
        }

        TableRow row = (TableRow)_tableCell.Parent;
        Table table = _tableCell.Ancestors<Table>().First();
        if (_tableCell.TableCellProperties?.VerticalMerge != null || _tableCell.TableCellProperties?.HorizontalMerge != null)
        {
            throw new NotSupportedException("Splitting a cell that is merged is not supported.");
        }
        int span = Span(_tableCell);
        if (numColumns > 1 && span > 1 && span % numColumns != 0)
        {
            throw new NotSupportedException("A cell that spans columns can only be split into a number of columns that divides the span.");
        }
        if (numRows > 1 && row.Elements<TableCell>().Any(c => c != _tableCell && (c.TableCellProperties?.VerticalMerge != null || c.TableCellProperties?.HorizontalMerge != null)))
        {
            throw new NotSupportedException("Splitting a cell into rows is not supported when another cell in its row is vertically merged.");
        }
        List<TableCell> pieces = new List<TableCell>() { _tableCell };
        if (numColumns > 1)
        {
            int width = CellWidth;
            int pieceSpan = span / numColumns;
            if (span == 1)
            {
                SplitGridColumn(table, GetGridStart(_tableCell), numColumns, _tableCell);
                pieceSpan = 1;
            }

            TableCell previous = _tableCell;
            for (int i = 0; i < numColumns; i++)
            {
                TableCell piece = i == 0 ? _tableCell : CreateLike(_tableCell);
                if (i > 0)
                {
                    previous.InsertAfterSelf(piece);
                    pieces.Add(piece);
                }
                TableCellProperties properties = GetOrAddTableCellProperties(piece);
                properties.GridSpan = pieceSpan > 1 ? new GridSpan() { Val = pieceSpan } : null;
                new TableCellWrapper(piece).CellWidth = i == numColumns - 1 ? width - (width / numColumns * (numColumns - 1)) : width / numColumns;
                previous = piece;
            }
        }

        if (numRows > 1)
        {
            foreach (TableCell other in row.Elements<TableCell>().Where(c => !pieces.Contains(c)))
            {
                GetOrAddTableCellProperties(other).VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
            }

            TableRow last = row;
            for (int i = 1; i < numRows; i++)
            {
                TableRow newRow = new TableRow();
                if (row.TableRowProperties != null)
                {
                    newRow.AppendChild(row.TableRowProperties.CloneNode(true));
                }
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    TableCell newCell = CreateLike(cell);
                    if (!pieces.Contains(cell))
                    {
                        GetOrAddTableCellProperties(newCell).VerticalMerge = new VerticalMerge();
                    }
                    newRow.AppendChild(newCell);
                }
                last.InsertAfterSelf(newRow);
                last = newRow;
            }
        }
    }

    private static int Span(TableCell cell)
    {
        return cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
    }

    private static int GetGridStart(TableCell cell)
    {
        TableRow row = (TableRow)cell.Parent;
        int start = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
        foreach (TableCell before in row.Elements<TableCell>().TakeWhile(c => c != cell))
        {
            start += Span(before);
        }
        return start;
    }

    private static void GetGridRange(TableCell cell, out int start, out int end)
    {
        start = GetGridStart(cell);
        end = start + Span(cell) - 1;
    }

    /// <summary>
    /// Adds the paragraphs of a cell that have content to the end of another cell.
    /// </summary>
    private static void MoveText(TableCell from, TableCell to)
    {
        foreach (Paragraph paragraph in from.Elements<Paragraph>().Where(p => p.Descendants<Run>().Any()).ToList())
        {
            to.AppendChild(paragraph.CloneNode(true));
        }
    }

    /// <summary>
    /// A new empty cell with the properties of the cell, except for any merge.
    /// </summary>
    private static TableCell CreateLike(TableCell cell)
    {
        TableCell newCell = new TableCell();
        if (cell.TableCellProperties != null)
        {
            TableCellProperties properties = (TableCellProperties)cell.TableCellProperties.CloneNode(true);
            properties.RemoveAllChildren<VerticalMerge>();
            properties.RemoveAllChildren<HorizontalMerge>();
            newCell.AppendChild(properties);
        }
        Paragraph paragraph = new Paragraph();
        ParagraphProperties paragraphProperties = cell.Elements<Paragraph>().FirstOrDefault()?.ParagraphProperties;
        if (paragraphProperties != null)
        {
            paragraph.AppendChild(paragraphProperties.CloneNode(true));
        }
        newCell.AppendChild(paragraph);
        return newCell;
    }

    /// <summary>
    /// Divides grid column <paramref name="index"/> into <paramref name="count"/> and widens the other cells that span it. The
    /// <paramref name="target"/> cell that is being split is the one whose span is not changed.
    /// </summary>
    private static void SplitGridColumn(Table table, int index, int count, TableCell target)
    {
        GridColumn gridColumn = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ElementAtOrDefault(index);
        if (gridColumn == null)
        {
            throw new InvalidOperationException("The cell is not in the grid of its table.");
        }

        int width = gridColumn.Width == null ? 0 : Convert.ToInt32(gridColumn.Width.Value, CultureInfo.InvariantCulture);
        int each = width / count;
        gridColumn.Width = each.ToString(CultureInfo.InvariantCulture);
        GridColumn previous = gridColumn;
        for (int i = 1; i < count; i++)
        {
            int newWidth = i == count - 1 ? width - (each * (count - 1)) : each;
            GridColumn added = new GridColumn() { Width = newWidth.ToString(CultureInfo.InvariantCulture) };
            previous.InsertAfterSelf(added);
            previous = added;
        }

        foreach (TableRow row in table.Elements<TableRow>())
        {
            int position = row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value ?? 0;
            if (index < position)
            {
                row.TableRowProperties.GetFirstChild<GridBefore>().Val += count - 1;
                continue;
            }

            bool found = false;
            foreach (TableCell cell in row.Elements<TableCell>())
            {
                int span = Span(cell);
                if (index >= position && index < position + span)
                {
                    //The cell being split gets its new neighbours later; every other cell that spans the column has to grow.
                    if (cell != target)
                    {
                        GetOrAddTableCellProperties(cell).GridSpan = new GridSpan() { Val = span + count - 1 };
                    }
                    found = true;
                    break;
                }
                position += span;
            }

            GridAfter gridAfter = row.TableRowProperties?.GetFirstChild<GridAfter>();
            if (!found && gridAfter != null)
            {
                gridAfter.Val += count - 1;
            }
        }
    }
}
