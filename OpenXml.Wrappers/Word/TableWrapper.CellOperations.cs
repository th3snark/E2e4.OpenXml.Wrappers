using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class TableCellWrapperCollection
{
    /// <summary>
    /// Merges all the cells into one, as <c>Cells.Merge</c> in Word: the rectangle of the table that holds all of them becomes one cell, with
    /// the text of the cells one after the other.
    /// </summary>
    /// <exception cref="ArgumentException">The cells are not all in one table.</exception>
    /// <exception cref="NotSupportedException">As <see cref="TableCellWrapper.Merge"/>.</exception>
    public void Merge()
    {
        IList<TableCellWrapper> cells = Items;
        if (cells.Count < 2)
        {
            return;
        }
        Table table = cells[0].TableCell.Ancestors<Table>().FirstOrDefault();
        if (table == null || cells.Any(c => c.TableCell.Ancestors<Table>().FirstOrDefault() != table))
        {
            throw new ArgumentException("The cells must be in the same table.");
        }

        List<TableRow> rows = table.Elements<TableRow>().ToList();
        int firstRow = cells.Min(c => rows.IndexOf((TableRow)c.TableCell.Parent));
        int lastRow = cells.Max(c => rows.IndexOf((TableRow)c.TableCell.Parent));
        int firstColumn = cells.Min(c => TableCellWrapper.GetGridStart(c.TableCell));
        int lastColumn = cells.Max(c => TableCellWrapper.GetGridStart(c.TableCell) + TableCellWrapper.Span(c.TableCell) - 1);
        TableCell topLeft = TableCellWrapper.CellAtGridColumn(rows[firstRow], firstColumn);
        TableCell bottomRight = TableCellWrapper.CellAtGridColumn(rows[lastRow], lastColumn);
        new TableCellWrapper(topLeft).Merge(new TableCellWrapper(bottomRight));
    }

    /// <summary>
    /// Splits every cell into <paramref name="numRows"/> rows and <paramref name="numColumns"/> columns, as <c>Cells.Split</c> in Word. With
    /// <paramref name="mergeBeforeSplit"/> the cells are first merged into one; the number of rows and columns is then that of the merged cell.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The number of rows or columns is less than 1.</exception>
    public void Split(int numRows, int numColumns, bool mergeBeforeSplit = false)
    {
        Guard.ThrowIfLessThan(numRows, 1);
        Guard.ThrowIfLessThan(numColumns, 1);

        IList<TableCellWrapper> cells = Items;
        if (mergeBeforeSplit && cells.Count > 1)
        {
            TableCell first = cells[0].TableCell;
            Merge();
            cells = new List<TableCellWrapper> { new TableCellWrapper(first) };
        }

        //Splitting whole rows into rows is adding rows, with no cell merged to the cells above or below.
        List<TableRow> wholeRows = WholeRows(cells);
        if (numColumns == 1 && numRows > 1 && wholeRows != null && !wholeRows.SelectMany(r => r.Elements<TableCell>()).Any(c => c.TableCellProperties?.VerticalMerge != null))
        {
            foreach (TableRow row in wholeRows)
            {
                //Word gives the row and the rows made from it the height it measured, one line of the default 11 point font.
                if (row.TableRowProperties?.GetFirstChild<TableRowHeight>() == null)
                {
                    new TableRowWrapper(row).GetOrAddProperties().AppendChild(new TableRowHeight() { Val = SplitRowHeight });
                }
                for (int added = 1; added < numRows; added++)
                {
                    TableRow newRow = (TableRow)row.CloneNode(false);
                    foreach (OpenXmlElement part in row.ChildElements.Where(e => e is TablePropertyExceptions || e is TableRowProperties))
                    {
                        newRow.AppendChild(part.CloneNode(true));
                    }
                    foreach (TableCell cell in row.Elements<TableCell>())
                    {
                        newRow.AppendChild(TableCellWrapper.NewCellLike(cell));
                    }
                    row.InsertAfterSelf(newRow);
                }
            }
            return;
        }

        foreach (TableCellWrapper cell in cells)
        {
            cell.Split(numRows, numColumns);
        }
    }

    private const uint SplitRowHeight = 252;

    //The rows that the cells are all the cells of, or null when they are not whole rows.
    private static List<TableRow> WholeRows(IList<TableCellWrapper> cells)
    {
        List<TableRow> rows = cells.Select(c => (TableRow)c.TableCell.Parent).Distinct().ToList();
        return rows.All(r => r.Elements<TableCell>().All(c => cells.Any(x => ReferenceEquals(x.TableCell, c)))) ? rows : null;
    }

    /// <summary>
    /// Adds a cell to the row, before <paramref name="beforeCell"/> or, without it, at the end, as <c>Cells.Add</c> in Word. The new cell is
    /// like the cell it is put before, or the last, and the grid grows by its width: the other rows are left short.
    /// </summary>
    /// <exception cref="InvalidOperationException">The cells are not those of a row.</exception>
    /// <exception cref="ArgumentException">The cell is not in the row.</exception>
    public TableCellWrapper Add(TableCellWrapper beforeCell = null)
    {
        if (!(_range.StartElement is TableRow row))
        {
            throw new InvalidOperationException("Cells can only be added to the cells of a row.");
        }
        if (beforeCell != null && !ReferenceEquals(beforeCell.TableCell.Parent, row))
        {
            throw new ArgumentException("The cell must be a cell of this row.", nameof(beforeCell));
        }

        TableCell template = beforeCell?.TableCell ?? row.Elements<TableCell>().LastOrDefault();
        if (template == null)
        {
            throw new InvalidOperationException("A cell cannot be added to a row that has no cells to use as a template.");
        }
        TableCell newCell = TableCellWrapper.NewCellLike(template);
        if (beforeCell == null)
        {
            row.AppendChild(newCell);
        }
        else
        {
            template.InsertBeforeSelf(newCell);
        }
        Table table = row.Ancestors<Table>().FirstOrDefault();
        TableWidths.FitGridToCells(table);
        if (table != null)
        {
            TableWrapper.ColumnsChanged(table);
        }
        return new TableCellWrapper(newCell);
    }

    /// <summary>
    /// Deletes all the cells, as <c>Cells.Delete</c> in Word. With <see cref="ShiftCells.Left"/> the cells after them move left; when they are
    /// all the cells of some columns the columns are deleted. With <see cref="ShiftCells.Up"/> the cells below move up.
    /// </summary>
    /// <exception cref="NotSupportedException">A cell that is part of a vertical merge.</exception>
    public void Delete(ShiftCells shiftCells = ShiftCells.Left)
    {
        if (!EnumCompat.IsDefined(shiftCells))
        {
            throw new ArgumentOutOfRangeException(nameof(shiftCells));
        }
        IList<TableCellWrapper> cells = Items;
        if (cells.Count == 0)
        {
            return;
        }

        if (shiftCells == ShiftCells.EntireRow)
        {
            foreach (TableRow row in cells.Select(c => (TableRow)c.TableCell.Parent).Distinct().Reverse().ToList())
            {
                new TableRowWrapper(row).Delete();
            }
        }
        else if (shiftCells == ShiftCells.EntireColumn)
        {
            DeleteColumns(cells.GroupBy(c => c.TableCell.Ancestors<Table>().First()).ToList());
        }
        else if (shiftCells == ShiftCells.Up)
        {
            TableCellWrapper.ShiftUp(cells);
        }
        else
        {
            List<IGrouping<Table, TableCellWrapper>> wholeColumns = cells.GroupBy(c => c.TableCell.Ancestors<Table>().First()).Where(g => CoversWholeColumns(g)).ToList();
            HashSet<Table> columnTables = new HashSet<Table>(wholeColumns.Select(g => g.Key));
            List<TableCellWrapper> others = cells.Where(c => !columnTables.Contains(c.TableCell.Ancestors<Table>().First())).Reverse().ToList();
            DeleteColumns(wholeColumns);
            foreach (TableCellWrapper cell in others)
            {
                cell.Delete(ShiftCells.Left);
            }
        }
    }

    //Whether the cells are all the cells of one or more grid columns of their table.
    private static bool CoversWholeColumns(IGrouping<Table, TableCellWrapper> group)
    {
        List<TableRow> rows = group.Key.Elements<TableRow>().ToList();
        if (group.Any(c => TableCellWrapper.Span(c.TableCell) != 1))
        {
            return false;
        }
        foreach (int column in group.Select(c => TableCellWrapper.GetGridStart(c.TableCell)).Distinct())
        {
            if (rows.Any(r => !group.Any(g => ReferenceEquals(g.TableCell, TableCellWrapper.CellAtGridColumn(r, column)))))
            {
                return false;
            }
        }
        return true;
    }

    private static void DeleteColumns(IList<IGrouping<Table, TableCellWrapper>> groups)
    {
        foreach (IGrouping<Table, TableCellWrapper> group in groups)
        {
            TableWrapper table = new TableWrapper(group.Key);
            foreach (int index in group.Select(c => TableCellWrapper.GetGridStart(c.TableCell)).Distinct().OrderByDescending(i => i).ToList())
            {
                table.Columns[index].Delete();
            }
        }
    }
}

public partial class TableCellWrapper
{
    //A new row's cell: empty, with the properties of the cell except for any merge.
    internal static TableCell NewCellLike(TableCell cell)
    {
        return CreateLike(cell);
    }

    //Deletes the cells and moves the content of the cells below them up in their columns; the cells at the bottom are left empty.
    internal static void ShiftUp(IList<TableCellWrapper> cells)
    {
        foreach (IGrouping<Table, TableCellWrapper> table in cells.GroupBy(c => c.TableCell.Ancestors<Table>().First()))
        {
            List<TableRow> rows = table.Key.Elements<TableRow>().ToList();
            if (table.Any(c => Span(c.TableCell) != 1))
            {
                throw new NotSupportedException("Shifting up cells that span columns is not supported.");
            }
            foreach (IGrouping<int, TableCellWrapper> selected in table.GroupBy(c => GetGridStart(c.TableCell)))
            {
                List<TableCell> column = rows.Select(r => CellAtGridColumn(r, selected.Key)).ToList();
                if (column.Any(c => c == null || Span(c) != 1 || c.TableCellProperties?.VerticalMerge != null))
                {
                    throw new NotSupportedException("Shifting cells up in a column with merged cells is not supported.");
                }

                List<List<OpenXmlElement>> contents = column
                    .Where(c => !selected.Any(s => ReferenceEquals(s.TableCell, c)))
                    .Select(c => c.ChildElements.Where(e => !(e is TableCellProperties)).ToList())
                    .ToList();
                RangeWrapper.DeleteElements(column.Where(c => selected.Any(s => ReferenceEquals(s.TableCell, c))).SelectMany(c => c.ChildElements.Where(e => !(e is TableCellProperties))).ToList());
                foreach (List<OpenXmlElement> content in contents)
                {
                    foreach (OpenXmlElement element in content)
                    {
                        element.Remove();
                    }
                }
                for (int index = 0; index < column.Count; index++)
                {
                    if (index < contents.Count)
                    {
                        foreach (OpenXmlElement element in contents[index])
                        {
                            column[index].AppendChild(element);
                        }
                    }
                    if (!column[index].Elements<Paragraph>().Any())
                    {
                        column[index].AppendChild(new Paragraph());
                    }
                }
            }
        }
    }
}
