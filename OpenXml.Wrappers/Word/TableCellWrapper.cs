using System.Globalization;

using System;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableCellWrapper : System.IEquatable<TableCellWrapper>
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

    /// <summary>
    /// The borders of the cell, as <c>Cell.Borders</c> in Word.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return new BordersWrapper(_tableCell, false); }
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
        Shading shading = tableCellProperties.Shading;
        if (shading == null)
        {
            //The typed property puts the element in schema order.
            shading = new Shading();
            tableCellProperties.Shading = shading;
        }
        shading.Val = ShadingPatternValues.Clear;
        shading.Color = "auto";
        shading.Fill = rgbColor;
    }

    /// <summary>
    /// The background colour of the cell as six hex digits, or null when it has no fill. Setting null removes the shading.
    /// </summary>
    public string BackgroundColor
    {
        get
        {
            string fill = _tableCell.TableCellProperties?.Shading?.Fill?.Value;
            return fill == null || String.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase) ? null : fill;
        }
        set
        {
            if (value == null)
            {
                _tableCell.TableCellProperties?.RemoveAllChildren<Shading>();
                return;
            }
            if (value.Length != 6 || !value.All(Uri.IsHexDigit))
            {
                throw new ArgumentException("The colour must be six hex digits.", nameof(value));
            }
            SetShadingFillColor(value);
        }
    }

    /// <summary>
    /// The vertical alignment of the text in the cell. The default is Top.
    /// </summary>
    public CellVerticalAlignment VerticalAlignment
    {
        get
        {
            TableVerticalAlignmentValues? alignment = _tableCell.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value;
            if (alignment == TableVerticalAlignmentValues.Center)
            {
                return CellVerticalAlignment.Center;
            }
            return alignment == TableVerticalAlignmentValues.Bottom ? CellVerticalAlignment.Bottom : CellVerticalAlignment.Top;
        }
        set
        {
            if (value == CellVerticalAlignment.Top)
            {
                _tableCell.TableCellProperties?.RemoveAllChildren<TableCellVerticalAlignment>();
                return;
            }
            GetOrAddTableCellProperties().TableCellVerticalAlignment = new TableCellVerticalAlignment()
            {
                Val = value == CellVerticalAlignment.Center ? TableVerticalAlignmentValues.Center : TableVerticalAlignmentValues.Bottom
            };
        }
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

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as TableCellWrapper);
    }

    public bool Equals(TableCellWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_tableCell, other._tableCell);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_tableCell);
    }

    public static bool operator ==(TableCellWrapper left, TableCellWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(TableCellWrapper left, TableCellWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
