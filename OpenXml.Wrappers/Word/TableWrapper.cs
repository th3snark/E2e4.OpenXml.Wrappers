using System;
using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableWrapper : System.IEquatable<TableWrapper>
{
    Table _table;

    public TableWrapper(Table table)
    {
        _table = table;
    }

    public Table Table
    {
        get { return _table; }
    }

    /// <summary>
    /// The name of the table style, as <c>Table.Style</c> in Word. Setting null removes the table style. See <see cref="RangeWrapper.Style"/>
    /// for how names are matched.
    /// </summary>
    /// <exception cref="ArgumentException">There is no table style with the name.</exception>
    /// <exception cref="InvalidOperationException">The document does not have any styles.</exception>
    public string Style
    {
        get
        {
            string styleId = _table.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
            return styleId == null ? StyleHelper.GetDefaultStyleName(_table, StyleValues.Table) : StyleHelper.GetStyleName(_table, styleId);
        }
        set
        {
            if (value == null)
            {
                _table.GetFirstChild<TableProperties>()?.RemoveAllChildren<TableStyle>();
                return;
            }

            string styleId = StyleHelper.FindStyleId(_table, value, out StyleValues type);
            if (type != StyleValues.Table)
            {
                throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The style '{0}' is not a table style.", value), nameof(value));
            }

            TableProperties tableProperties = _table.GetFirstChild<TableProperties>();
            if (tableProperties == null)
            {
                tableProperties = new TableProperties();
                _table.PrependChild(tableProperties);
            }
            tableProperties.TableStyle = new TableStyle() { Val = styleId };
        }
    }

    /// <summary>
    /// The borders of the table, as <c>Table.Borders</c> in Word. Diagonal borders are not supported.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return new BordersWrapper(_table, true); }
    }

    private TableProperties GetOrAddTableProperties()
    {
        TableProperties tableProperties = _table.GetFirstChild<TableProperties>();
        if (tableProperties == null)
        {
            tableProperties = new TableProperties();
            _table.PrependChild(tableProperties);
        }
        return tableProperties;
    }

    /// <summary>
    /// The horizontal alignment of the table on the page, as <c>Table.Rows.Alignment</c> in Word. The default is Left.
    /// </summary>
    public RowAlignment Alignment
    {
        get
        {
            TableRowAlignmentValues? alignment = _table.GetFirstChild<TableProperties>()?.TableJustification?.Val?.Value;
            if (alignment == TableRowAlignmentValues.Center)
            {
                return RowAlignment.Center;
            }
            return alignment == TableRowAlignmentValues.Right ? RowAlignment.Right : RowAlignment.Left;
        }
        set
        {
            if (value == RowAlignment.Left)
            {
                _table.GetFirstChild<TableProperties>()?.RemoveAllChildren<TableJustification>();
                return;
            }
            GetOrAddTableProperties().TableJustification = new TableJustification()
            {
                Val = value == RowAlignment.Center ? TableRowAlignmentValues.Center : TableRowAlignmentValues.Right
            };
        }
    }

    /// <summary>
    /// The preferred width of the table in points, or null when it is not a fixed width (automatic or a percentage). Setting null
    /// removes the preferred width.
    /// </summary>
    public double? PreferredWidth
    {
        get
        {
            TableWidth width = _table.GetFirstChild<TableProperties>()?.TableWidth;
            if (width?.Type == null || !width.Type.HasValue || width.Type.Value != TableWidthUnitValues.Dxa || width.Width == null)
            {
                return null;
            }
            return Double.TryParse(width.Width.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips) ? twips / 20.0 : null;
        }
        set
        {
            if (value == null)
            {
                _table.GetFirstChild<TableProperties>()?.RemoveAllChildren<TableWidth>();
                return;
            }
            if (Double.IsNaN(value.Value) || value.Value < 0 || value.Value > 1584)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The width must be between 0 and 1584 points.");
            }
            GetOrAddTableProperties().TableWidth = new TableWidth()
            {
                Width = Math.Round(value.Value * 20, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture),
                Type = TableWidthUnitValues.Dxa
            };
        }
    }

    public TableWrapper Parent
    {
        get
        {
            Table table = _table.ParentAtRelativeNestingLevel<Table>(-1);
            return table != null ? new TableWrapper(table) : null;
        }
    }

    public int NestingLevel
    {
        get { return _table.NestingLevel<Table>(); }
    }

    public TableWrapperCollection Tables
    {
        get { return new TableWrapperCollection(new RangeWrapper(_table)); }
    }

    public TableRowWrapperCollection Rows
    {
        get { return new TableRowWrapperCollection(new RangeWrapper(_table)); }
    }

    public TableColumnWrapperCollection Columns
    {
        get { return new TableColumnWrapperCollection(new RangeWrapper(_table)); }
    }

    public TableCellWrapper Cell(int rowIndex, int columnIndex)
    {
        return Rows[rowIndex].Cells[columnIndex];
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_table); }
    }

    public void Delete()
    {
        Range.Delete();
    }

    public void AppendRow(TableRowWrapper row)
    {
        _table.AppendChild(row.TableRow);
    }

    public void MergeTable(TableWrapper reportTable)
    {
        _table.MergeTable(reportTable.Table);
    }

    public ParagraphWrapper InsertParagraphAfterSelf()
    {
        return Range.InsertParagraphAfter();
    }
    public int Width
    {
        get { return Convert.ToInt32(_table.GetTableWidth()); }
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as TableWrapper);
    }

    public bool Equals(TableWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_table, other._table);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_table);
    }

    public static bool operator ==(TableWrapper left, TableWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(TableWrapper left, TableWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
