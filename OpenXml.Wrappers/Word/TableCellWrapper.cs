using System.Globalization;

using System;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableCellWrapper
{
    TableCell _tableCell;

    public TableCellWrapper(TableCell tableCell)
    {
        _tableCell = tableCell;
    }

    public TableCell TableCell
    {
        get { return _tableCell; }
    }

    public int RowIndex
    {
        get { return Parent.Index; }
    }

    public int ColumnIndex
    {
        get { return _tableCell.GetIndex(); }
    }

    public TableRowWrapper Parent
    {
        get { return new TableRowWrapper((TableRow)_tableCell.Parent); }
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_tableCell); }
    }

    public int CellWidth
    {
        get
        {
            TableCellWidth tableCellWidth = _tableCell.TableCellProperties?.TableCellWidth;
            return tableCellWidth?.Width == null ? 0 : Convert.ToInt32(tableCellWidth.Width.Value, CultureInfo.InvariantCulture);
        }
        set
        {
            TableCellProperties tableCellProperties = GetOrAddTableCellProperties();
            if (tableCellProperties.TableCellWidth == null)
            {
                tableCellProperties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa };
            }
            tableCellProperties.TableCellWidth.Width = value.ToString(CultureInfo.InvariantCulture);
        }
    }

    private TableCellProperties GetOrAddTableCellProperties()
    {
        return GetOrAddTableCellProperties(_tableCell);
    }

    private static TableCellProperties GetOrAddTableCellProperties(TableCell tableCell)
    {
        TableCellProperties tableCellProperties = tableCell.TableCellProperties;
        if (tableCellProperties == null)
        {
            tableCellProperties = new TableCellProperties();
            tableCell.PrependChild(tableCellProperties);
        }
        return tableCellProperties;
    }

    public void MergeDown()
    {
        TableRow lastTableRow = _tableCell.Parent.NextSibling<TableRow>();
        if (lastTableRow != null)
        {
            TableCellWrapper lastCell = new TableCellWrapper(lastTableRow.Elements<TableCell>().ElementAt(ColumnIndex));
            MergeDown(lastCell);
        }
    }

    public void MergeDown(TableCellWrapper lastCell)
    {
        TableRow firstTableRow = (TableRow)_tableCell.Parent;
        TableRow lastTableRow = lastCell.Parent.TableRow;
        if (firstTableRow.IsBefore(lastTableRow))
        {
            int index = ColumnIndex;
            TableCell tableCell = null;
            //Start at the last row and stop when the first row is reached adding in a new vertcal merge to the appropriate table cell properties.
            for (TableRow tableRow = lastTableRow; tableRow != firstTableRow; tableRow = tableRow.PreviousSibling<TableRow>())
            {
                tableCell = tableRow.Elements<TableCell>().ElementAt(index);
                GetOrAddTableCellProperties(tableCell).VerticalMerge = new VerticalMerge(); //default val is 'continue'
            }

            //If the first cell in the vertical merge has no vertcal merge setting then set it up with a "restart" vertical merge
            tableCell = firstTableRow.Elements<TableCell>().ElementAt(index);
            TableCellProperties firstTableCellProperties = GetOrAddTableCellProperties(tableCell);
            if (firstTableCellProperties.VerticalMerge == null)
            {
                firstTableCellProperties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
            }
        }
    }

    public void MergeRight()
    {
        TableCell nextSiblingRightTableCell = _tableCell.NextSibling<TableCell>();
        if (nextSiblingRightTableCell != null)
        {
            TableCellWrapper rightCell = new TableCellWrapper(nextSiblingRightTableCell);

            //Add the width of the table cell to the right and add it to the with of this table cell.
            CellWidth += rightCell.CellWidth;

            //Add a grid span element if one does not already exist.
            TableCellProperties tableCellProperties = GetOrAddTableCellProperties();
            if (tableCellProperties.GridSpan == null)
            {
                tableCellProperties.GridSpan = new GridSpan();
                tableCellProperties.GridSpan.Val = 1;
            }

            //Increase the grid span attribute value by one.
            tableCellProperties.GridSpan.Val++;

            //Remove the table cell to the right.
            nextSiblingRightTableCell.Parent.RemoveChild(nextSiblingRightTableCell);
        }
    }

    public bool KeepNext
    {
        set
        {
            foreach (ParagraphWrapper paragraph in Range.Paragraphs)
            {
                paragraph.KeepNext = value;
            }
        }
    }

    public void SetShadingFillColor(string rgbColor)
    {
        TableCellProperties tableCellProperties = GetOrAddTableCellProperties();
        Shading shading = tableCellProperties.Elements<Shading>().FirstOrDefault();
        if (shading == null)
        {
            shading = new Shading();
            tableCellProperties.AppendChild(shading);
        }
        shading.Val = ShadingPatternValues.Clear;
        shading.Color = "auto";
        shading.Fill = rgbColor;
    }

    public void SetFontBold()
    {
        foreach (ParagraphWrapper paragraph in Range.Paragraphs)
        {
            paragraph.SetFontBold();
        }
    }

    public void SetFontColor(string color)
    {
        foreach (ParagraphWrapper paragraph in Range.Paragraphs)
        {
            paragraph.SetFontColor(color);
        }
    }
}
