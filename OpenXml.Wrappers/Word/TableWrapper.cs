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
        _table.Remove();
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
        Paragraph paragraph = new Paragraph();
        _table.InsertAfterSelf(paragraph);
        return new ParagraphWrapper(paragraph);
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
