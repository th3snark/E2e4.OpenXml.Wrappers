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
        get { return Convert.ToInt32(_gridColumn.Width.Value, CultureInfo.InvariantCulture); }
        set
        {
            _gridColumn.Width.Value = value.ToString(CultureInfo.InvariantCulture);
            int index = _gridColumn.ElementsBefore().Where(e => e is GridColumn).Count();
            Table table = _gridColumn.Ancestors<Table>().First();
            foreach (TableRow tableRow in table.Elements<TableRow>())
            {
                TableCell tableCell = tableRow.Elements<TableCell>().ElementAt<TableCell>(index);
                tableCell.TableCellProperties.TableCellWidth.Width.Value = value.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}
