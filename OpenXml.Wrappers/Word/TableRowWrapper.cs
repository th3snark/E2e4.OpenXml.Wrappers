using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableRowWrapper
{
    TableRow _tableRow;

    public TableRowWrapper(TableRow tableRow)
    {
        _tableRow = tableRow;
    }

    public TableRow TableRow
    {
        get { return _tableRow; }
    }

    public TableWrapper Parent
    {
        get { return new TableWrapper((Table)_tableRow.Parent); }
    }

    public TableCellWrapperCollection Cells
    {
        get { return new TableCellWrapperCollection(new RangeWrapper(_tableRow)); }
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_tableRow); }
    }

    public int Index
    {
        get { return _tableRow.ElementsBefore().Where(element => element is TableRow).Count(); }
    }

    public void Delete()
    {
        _tableRow.Parent.RemoveChild(_tableRow);
    }

    public void ClearText()
    {
        _tableRow.ClearText();
    }

    public TableRowWrapper Copy()
    {
        return new TableRowWrapper((TableRow)_tableRow.Clone());
    }

    public bool KeepNext
    {
        set
        {
            foreach (TableCellWrapper cell in Cells)
            {
                cell.KeepNext = value;
            }
        }
    }

    public void SetTopBorder(bool visible)
    {
        foreach (TableCell tableCell in _tableRow.Elements<TableCell>())
        {
            if (tableCell.TableCellProperties == null)
            {
                tableCell.TableCellProperties = new TableCellProperties();
            }
            TableCellProperties tableCellProperties = tableCell.TableCellProperties;

            if (tableCellProperties.TableCellBorders == null)
            {
                tableCellProperties.TableCellBorders = new TableCellBorders();
            }
            TableCellBorders tableCellBorders = tableCellProperties.TableCellBorders;

            if (tableCellBorders.TopBorder == null)
            {
                tableCellBorders.TopBorder = new TopBorder();
            }
            TopBorder topBorder = tableCellBorders.TopBorder;

            topBorder.Val = visible ? BorderValues.Single : BorderValues.Nil;
        }
    }

    public void SetBottomBorder(bool visible)
    {
        foreach (TableCell tableCell in _tableRow.Elements<TableCell>())
        {
            if (tableCell.TableCellProperties == null)
            {
                tableCell.TableCellProperties = new TableCellProperties();
            }
            TableCellProperties tableCellProperties = tableCell.TableCellProperties;

            if (tableCellProperties.TableCellBorders == null)
            {
                tableCellProperties.TableCellBorders = new TableCellBorders();
            }
            TableCellBorders tableCellBorders = tableCellProperties.TableCellBorders;

            if (tableCellBorders.BottomBorder == null)
            {
                tableCellBorders.BottomBorder = new BottomBorder();
            }
            BottomBorder bottomBorder = tableCellBorders.BottomBorder;

            bottomBorder.Val = visible ? BorderValues.Single : BorderValues.Nil;
        }
    }
}
