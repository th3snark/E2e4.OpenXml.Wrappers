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
    /// The 0-based position of the column in the table grid.
    /// </summary>
    public int Index
    {
        get { return _gridColumn.ElementsBefore().Count(e => e is GridColumn); }
    }

    /// <summary>
    /// Deletes the column and the cell at its position in every row. Deleting the only column deletes the table.
    /// </summary>
    /// <exception cref="NotSupportedException">The table has merged cells, so cells do not line up with the grid.</exception>
    public void Delete()
    {
        Table table = _gridColumn.Ancestors<Table>().First();
        List<TableCell> cells = table.Elements<TableRow>().SelectMany(r => r.Elements<TableCell>()).ToList();
        if (cells.Any(c => c.TableCellProperties?.GridSpan?.Val?.Value > 1 || c.TableCellProperties?.HorizontalMerge != null))
        {
            throw new NotSupportedException("Deleting a column of a table with horizontally merged cells is not supported.");
        }

        if (table.GetFirstChild<TableGrid>().Elements<GridColumn>().Count() == 1)
        {
            table.Remove();
            return;
        }

        int index = Index;
        foreach (TableRow tableRow in table.Elements<TableRow>())
        {
            tableRow.Elements<TableCell>().ElementAtOrDefault(index)?.Remove();
        }
        _gridColumn.Remove();
    }

    public int Width
    {
        get { return _gridColumn.Width == null ? 0 : Convert.ToInt32(_gridColumn.Width.Value, CultureInfo.InvariantCulture); }
        set
        {
            string width = value.ToString(CultureInfo.InvariantCulture);
            _gridColumn.Width = width;
            int index = _gridColumn.ElementsBefore().Where(e => e is GridColumn).Count();
            Table table = _gridColumn.Ancestors<Table>().First();
            foreach (TableRow tableRow in table.Elements<TableRow>())
            {
                //Rows with fewer cells than the grid has columns (for example merged or short rows) are skipped.
                TableCell tableCell = tableRow.Elements<TableCell>().ElementAtOrDefault(index);
                if (tableCell != null)
                {
                    tableCell.SetWidth(value, TableWidthUnitValues.Dxa);
                }
            }
        }
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
