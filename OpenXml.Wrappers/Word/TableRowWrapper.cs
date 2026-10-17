using System;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableRowWrapper : System.IEquatable<TableRowWrapper>
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

    /// <summary>
    /// The row after this one in the table, or null for the last row (<c>Row.Next</c> in Word).
    /// </summary>
    public TableRowWrapper Next()
    {
        TableRow row = _tableRow.NextSibling<TableRow>();
        return row == null ? null : new TableRowWrapper(row);
    }

    /// <summary>
    /// The row before this one in the table, or null for the first row (<c>Row.Previous</c> in Word).
    /// </summary>
    public TableRowWrapper Previous()
    {
        TableRow row = _tableRow.PreviousSibling<TableRow>();
        return row == null ? null : new TableRowWrapper(row);
    }

    public bool IsFirst
    {
        get { return _tableRow.PreviousSibling<TableRow>() == null; }
    }

    public bool IsLast
    {
        get { return _tableRow.NextSibling<TableRow>() == null; }
    }

    public int Index
    {
        get { return _tableRow.ElementsBefore().Where(element => element is TableRow).Count(); }
    }

    public void Delete()
    {
        Range.Delete();
    }

    public void ClearText()
    {
        _tableRow.ClearText();
    }

    public TableRowWrapper Copy()
    {
        return new TableRowWrapper((TableRow)_tableRow.Clone());
    }

    private TableRowProperties GetOrAddTableRowProperties()
    {
        TableRowProperties properties = _tableRow.TableRowProperties;
        if (properties == null)
        {
            properties = new TableRowProperties();
            TablePropertyExceptions exceptions = _tableRow.GetFirstChild<TablePropertyExceptions>();
            if (exceptions != null)
            {
                exceptions.InsertAfterSelf(properties);
            }
            else
            {
                _tableRow.PrependChild(properties);
            }
        }
        return properties;
    }

    /// <summary>
    /// The height of the row in points, or 0 when the row has no explicit height (<see cref="HeightRule"/> is then Auto).
    /// </summary>
    public double Height
    {
        get
        {
            TableRowHeight height = _tableRow.TableRowProperties?.GetFirstChild<TableRowHeight>();
            return height?.Val == null ? 0 : height.Val.Value / 20.0;
        }
        set
        {
            if (Double.IsNaN(value) || value < 0 || value > 1584)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The height must be between 0 and 1584 points.");
            }
            TableRowProperties properties = GetOrAddTableRowProperties();
            TableRowHeight height = properties.GetFirstChild<TableRowHeight>();
            if (height == null)
            {
                height = new TableRowHeight() { HeightType = HeightRuleValues.AtLeast };
                properties.AppendChild(height);
            }
            height.Val = (uint)Math.Round(value * 20, MidpointRounding.AwayFromZero);
        }
    }

    /// <summary>
    /// How <see cref="Height"/> is applied. Setting a rule on a row without a height uses a height of 0.
    /// </summary>
    public RowHeightRule HeightRule
    {
        get
        {
            TableRowHeight height = _tableRow.TableRowProperties?.GetFirstChild<TableRowHeight>();
            if (height?.HeightType == null || !height.HeightType.HasValue)
            {
                //Without an explicit rule, Word treats the height as a minimum.
                return height == null ? RowHeightRule.Auto : RowHeightRule.AtLeast;
            }
            if (height.HeightType.Value == HeightRuleValues.Exact)
            {
                return RowHeightRule.Exactly;
            }
            return height.HeightType.Value == HeightRuleValues.AtLeast ? RowHeightRule.AtLeast : RowHeightRule.Auto;
        }
        set
        {
            if (value == RowHeightRule.Auto)
            {
                _tableRow.TableRowProperties?.RemoveAllChildren<TableRowHeight>();
                return;
            }
            TableRowProperties properties = GetOrAddTableRowProperties();
            TableRowHeight height = properties.GetFirstChild<TableRowHeight>();
            if (height == null)
            {
                height = new TableRowHeight() { Val = 0 };
                properties.AppendChild(height);
            }
            height.HeightType = value == RowHeightRule.Exactly ? HeightRuleValues.Exact : HeightRuleValues.AtLeast;
        }
    }

    /// <summary>
    /// True when the row is repeated at the top of each page, as <c>Row.HeadingFormat</c> in Word.
    /// </summary>
    public bool HeadingFormat
    {
        get
        {
            TableHeader header = _tableRow.TableRowProperties?.GetFirstChild<TableHeader>();
            return header != null && IsOn(header.Val);
        }
        set { SetToggle<TableHeader>(value, () => new TableHeader()); }
    }

    /// <summary>
    /// True when the row may be split across two pages, as <c>Row.AllowBreakAcrossPages</c> in Word. This is the default.
    /// </summary>
    public bool AllowBreakAcrossPages
    {
        get
        {
            CantSplit cantSplit = _tableRow.TableRowProperties?.GetFirstChild<CantSplit>();
            return cantSplit == null || !IsOn(cantSplit.Val);
        }
        set { SetToggle<CantSplit>(!value, () => new CantSplit()); }
    }

    private static bool IsOn(EnumValue<OnOffOnlyValues> value)
    {
        return value == null || !value.HasValue || value.Value == OnOffOnlyValues.On;
    }

    private void SetToggle<T>(bool on, Func<T> create) where T : OpenXmlElement
    {
        if (on)
        {
            if (_tableRow.TableRowProperties?.GetFirstChild<T>() == null)
            {
                GetOrAddTableRowProperties().AppendChild(create());
            }
        }
        else
        {
            _tableRow.TableRowProperties?.RemoveAllChildren<T>();
        }
    }

    /// <summary>
    /// Whether every cell of the row keeps its paragraphs with the next one; false when there are no cells or only some do.
    /// </summary>
    public bool KeepNext
    {
        get
        {
            List<TableCellWrapper> cells = Cells.ToList();
            return cells.Count > 0 && cells.All(c => c.KeepNext);
        }
        set
        {
            foreach (TableCellWrapper cell in Cells)
            {
                cell.KeepNext = value;
            }
        }
    }

    /// <summary>
    /// Shows or hides the top border of the row; <c>Borders[BorderPosition.Top].LineStyle</c> does the same and can set more.
    /// </summary>
    public void SetTopBorder(bool visible)
    {
        Borders[BorderPosition.Top].LineStyle = visible ? BorderLineStyle.SingleLine : BorderLineStyle.None;
    }
    /// <summary>
    /// The borders of the row, as <c>Row.Borders</c> in Word. They are stored as the borders of the cells of the row. There is no horizontal or
    /// diagonal border.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return new BordersWrapper(_tableRow, false); }
    }

    /// <summary>
    /// Shows or hides the bottom border of the row; <c>Borders[BorderPosition.Bottom].LineStyle</c> does the same and can set more.
    /// </summary>
    public void SetBottomBorder(bool visible)
    {
        Borders[BorderPosition.Bottom].LineStyle = visible ? BorderLineStyle.SingleLine : BorderLineStyle.None;
    }
    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as TableRowWrapper);
    }

    public bool Equals(TableRowWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_tableRow, other._tableRow);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_tableRow);
    }

    public static bool operator ==(TableRowWrapper left, TableRowWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(TableRowWrapper left, TableRowWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
