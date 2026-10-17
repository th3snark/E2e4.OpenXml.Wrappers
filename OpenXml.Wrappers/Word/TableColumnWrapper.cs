using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableColumnWrapper
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
}
