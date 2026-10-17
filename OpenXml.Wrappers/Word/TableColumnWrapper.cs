using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableColumnWrapper : System.IEquatable<TableColumnWrapper>
{
    GridColumn _gridColumn;

    public TableColumnWrapper(GridColumn gridColumn)
    {
        _gridColumn = gridColumn;
    }

    public GridColumn GridColumn
    {
        get { return _gridColumn; }
    }

    /// <summary>
    /// The table the column belongs to.
    /// </summary>
    public TableWrapper Parent
    {
        get { return new TableWrapper(_gridColumn.Ancestors<Table>().First()); }
    }

    /// <summary>
    /// The 0-based position of the column in the table grid.
    /// </summary>
    public int Index
    {
        get { return _gridColumn.ElementsBefore().Count(e => e is GridColumn); }
    }

    /// <summary>
    /// Deletes the column and, in every row, the cell in that column. A cell that spans several columns loses one column of its span
    /// instead, and when the cell removed was the top of a vertical merge the next cell down takes its place. Deleting the only column
    /// deletes the table.
    /// </summary>
    /// <exception cref="NotSupportedException">The table uses the older horizontal merge (<c>hMerge</c>).</exception>
    public void Delete()
    {
        Table table = _gridColumn.Ancestors<Table>().First();
        if (table.Descendants<HorizontalMerge>().Any())
        {
            throw new NotSupportedException("Deleting a column of a table with horizontally merged cells of the older kind is not supported.");
        }

        List<GridColumn> grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().ToList();
        if (grid.Count == 1)
        {
            table.Remove();
            return;
        }

        int index = Index;
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        for (int r = 0; r < rows.Count; r++)
        {
            TableCell cell = TableCellWrapper.CellAtGridColumn(rows[r], index);
            if (cell == null)
            {
                continue;
            }
            int span = TableCellWrapper.Span(cell);
            if (span > 1)
            {
                cell.TableCellProperties.GridSpan = new GridSpan() { Val = span - 1 };
                TableCellWrapper wrapper = new TableCellWrapper(cell);
                wrapper.CellWidth = Math.Max(0, wrapper.CellWidth - Width);
                continue;
            }

            //The cell below a vertical merge that starts here becomes the start of what is left of it.
            if (cell.TableCellProperties?.VerticalMerge != null && cell.TableCellProperties.VerticalMerge.Val?.Value == MergedCellValues.Restart && r + 1 < rows.Count)
            {
                TableCell below = TableCellWrapper.CellAtGridColumn(rows[r + 1], index);
                if (below?.TableCellProperties?.VerticalMerge != null && below.TableCellProperties.VerticalMerge.Val?.Value != MergedCellValues.Restart)
                {
                    below.TableCellProperties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                }
            }
            cell.Remove();
        }
        _gridColumn.Remove();
    }

    public int Width
    {
        get { return _gridColumn.Width == null ? 0 : Convert.ToInt32(_gridColumn.Width.Value, CultureInfo.InvariantCulture); }
        set
        {
            _gridColumn.Width = value.ToString(CultureInfo.InvariantCulture);
            Table table = _gridColumn.Ancestors<Table>().First();
            List<GridColumn> grid = table.GetFirstChild<TableGrid>().Elements<GridColumn>().ToList();
            int index = Index;
            foreach (TableRow tableRow in table.Elements<TableRow>())
            {
                //The cell in this grid column is as wide as the grid columns it covers; rows that are too short for the grid are skipped.
                TableCell tableCell = TableCellWrapper.CellAtGridColumn(tableRow, index);
                if (tableCell != null)
                {
                    int start = TableCellWrapper.GetGridStart(tableCell);
                    int width = 0;
                    for (int i = start; i < start + TableCellWrapper.Span(tableCell) && i < grid.Count; i++)
                    {
                        width += new TableColumnWrapper(grid[i]).WidthOrZero;
                    }
                    tableCell.SetWidth(width, TableWidthUnitValues.Dxa);
                }
            }
        }
    }

    private int WidthOrZero
    {
        get { return _gridColumn.Width == null ? 0 : Convert.ToInt32(_gridColumn.Width.Value, CultureInfo.InvariantCulture); }
    }
    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as TableColumnWrapper);
    }

    public bool Equals(TableColumnWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_gridColumn, other._gridColumn);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_gridColumn);
    }

    public static bool operator ==(TableColumnWrapper left, TableColumnWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(TableColumnWrapper left, TableColumnWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
