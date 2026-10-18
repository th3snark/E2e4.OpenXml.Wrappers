using System.Globalization;

using System;
using OpenXml.Wrappers.Helpers;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class TableCellWrapper : System.IEquatable<TableCellWrapper>
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

    /// <summary>
    /// The next cell in the row, or the first cell of the next row; null for the last cell of the table (<c>Cell.Next</c> in Word).
    /// </summary>
    public TableCellWrapper Next()
    {
        TableCell cell = _tableCell.NextSibling<TableCell>();
        if (cell == null)
        {
            cell = ((TableRow)_tableCell.Parent).NextSibling<TableRow>()?.Elements<TableCell>().FirstOrDefault();
        }
        return cell == null ? null : new TableCellWrapper(cell);
    }

    /// <summary>
    /// The previous cell in the row, or the last cell of the previous row; null for the first cell of the table (<c>Cell.Previous</c> in Word).
    /// </summary>
    public TableCellWrapper Previous()
    {
        TableCell cell = _tableCell.PreviousSibling<TableCell>();
        if (cell == null)
        {
            cell = ((TableRow)_tableCell.Parent).PreviousSibling<TableRow>()?.Elements<TableCell>().LastOrDefault();
        }
        return cell == null ? null : new TableCellWrapper(cell);
    }

    /// <summary>
    /// The preferred width of the cell in points, 0 when none is set.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number or more than 31680 points.</exception>
    public double Width
    {
        get { return ConversionHelper.TwipsToPoints(WidthTwips); }
        set { WidthTwips = ConversionHelper.PointsToTwips(value, nameof(value)); }
    }

    internal int WidthTwips
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

    /// <summary>
    /// Whether every paragraph of the cell is kept on the same page as the next one; false when there are none or only some are.
    /// </summary>
    public bool KeepWithNext
    {
        get
        {
            List<ParagraphWrapper> paragraphs = Range.Paragraphs.ToList();
            return paragraphs.Count > 0 && paragraphs.All(p => p.KeepWithNext == true);
        }
        set
        {
            foreach (ParagraphWrapper paragraph in Range.Paragraphs)
            {
                paragraph.KeepWithNext = value;
            }
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
