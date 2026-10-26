using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What happens to the cells around a deleted cell, as <c>WdDeleteCells</c> in Word (with the same values).
/// </summary>
public enum ShiftCells
{
    /// <summary>The cells to the right move left, leaving a gap at the end of the row.</summary>
    Left = 0,

    /// <summary>The cells below move up. This is not supported.</summary>
    Up = 1,

    /// <summary>The whole row is deleted.</summary>
    EntireRow = 2,

    /// <summary>The whole column is deleted.</summary>
    EntireColumn = 3
}

/// <summary>
/// Reads and writes the widths of tables, columns and cells, which are a number and what it means.
/// </summary>
internal static class PreferredWidths
{
    public static PreferredWidthType GetKind(string width, TableWidthUnitValues? type)
    {
        if (type == TableWidthUnitValues.Pct)
        {
            return PreferredWidthType.Percent;
        }
        if (type == TableWidthUnitValues.Auto || type == TableWidthUnitValues.Nil || width == null)
        {
            return PreferredWidthType.Auto;
        }
        return PreferredWidthType.Points;
    }

    public static double GetValue(string width, TableWidthUnitValues? type)
    {
        PreferredWidthType kind = GetKind(width, type);
        if (kind == PreferredWidthType.Auto || string.IsNullOrEmpty(width))
        {
            return 0;
        }

        bool percent = width[width.Length - 1] == '%';
        if (!double.TryParse(percent ? width.Substring(0, width.Length - 1) : width, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
        {
            return 0;
        }
        if (kind == PreferredWidthType.Percent)
        {
            //Percentages are stored in fiftieths of a percent, or (in newer files) with a percent sign.
            return percent ? number : number / 50;
        }
        return number / 20;
    }

    public static void Format(PreferredWidthType kind, double value, out string width, out TableWidthUnitValues type)
    {
        if (!EnumCompat.IsDefined(kind))
        {
            throw new ArgumentOutOfRangeException(nameof(kind));
        }
        if (double.IsNaN(value) || value < 0 || value > (kind == PreferredWidthType.Percent ? 600 : 31680))
        {
            throw new ArgumentOutOfRangeException(nameof(value), kind == PreferredWidthType.Percent ? "The width must be between 0 and 600 percent." : "The width must be between 0 and 31680 points.");
        }

        if (kind == PreferredWidthType.Auto)
        {
            width = "0";
            type = TableWidthUnitValues.Auto;
        }
        else if (kind == PreferredWidthType.Percent)
        {
            width = Math.Round(value * 50, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
            type = TableWidthUnitValues.Pct;
        }
        else
        {
            width = Math.Round(value * 20, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
            type = TableWidthUnitValues.Dxa;
        }
    }

    //The twips in a margin or spacing element whatever the type of its width.
    public static double ToPoints(object width)
    {
        if (width == null)
        {
            return 0;
        }
        return double.TryParse(Convert.ToString(width, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out double twips) ? twips / 20 : 0;
    }

    public static string FromPoints(double points, string name)
    {
        if (double.IsNaN(points) || points < 0 || points > 1584)
        {
            throw new ArgumentOutOfRangeException(name, "The value must be between 0 and 1584 points.");
        }
        return Math.Round(points * 20, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
    }
}

public partial class TableWrapper
{
    private TableProperties ExistingProperties
    {
        get { return _table.GetFirstChild<TableProperties>(); }
    }

    /// <summary>
    /// What <see cref="PreferredWidth"/> means: points, a percentage or automatic. Changing it keeps the number.
    /// </summary>
    public PreferredWidthType PreferredWidthType
    {
        get
        {
            TableWidth width = ExistingProperties?.TableWidth;
            return PreferredWidths.GetKind(width?.Width?.Value, width?.Type?.Value);
        }
        set
        {
            double current = PreferredWidth;
            PreferredWidths.Format(value, current, out string width, out TableWidthUnitValues type);
            GetOrAddTableProperties().TableWidth = new TableWidth() { Width = width, Type = type };
        }
    }

    /// <summary>
    /// The preferred width of the table, in points or in percent according to <see cref="PreferredWidthType"/>; 0 for an automatic width.
    /// Setting it makes the width a width in points, unless it is a percentage already.
    /// </summary>
    public double PreferredWidth
    {
        get
        {
            TableWidth width = ExistingProperties?.TableWidth;
            return PreferredWidths.GetValue(width?.Width?.Value, width?.Type?.Value);
        }
        set
        {
            PreferredWidths.Format(PreferredWidthType == PreferredWidthType.Percent ? PreferredWidthType.Percent : PreferredWidthType.Points, value, out string width, out TableWidthUnitValues type);
            GetOrAddTableProperties().TableWidth = new TableWidth() { Width = width, Type = type };
        }
    }

    /// <summary>
    /// Whether the columns can resize to fit their content (<c>Table.AllowAutoFit</c> in Word). False gives fixed column widths.
    /// </summary>
    public bool AllowAutoFit
    {
        get
        {
            TableLayoutValues? layout = ExistingProperties?.TableLayout?.Type?.Value;
            return layout != TableLayoutValues.Fixed;
        }
        set
        {
            if (value)
            {
                ExistingProperties?.RemoveAllChildren<TableLayout>();
            }
            else
            {
                GetOrAddTableProperties().TableLayout = new TableLayout() { Type = TableLayoutValues.Fixed };
            }
        }
    }

    /// <summary>
    /// Sizes the table, as <c>Table.AutoFitBehavior</c> in Word: to fixed column widths, to the content, or to the width between the margins.
    /// </summary>
    public void AutoFitBehavior(AutoFitBehavior behavior)
    {
        if (!EnumCompat.IsDefined(behavior))
        {
            throw new ArgumentOutOfRangeException(nameof(behavior));
        }

        List<TableCell> cells = _table.Descendants<TableCell>().Where(c => c.Ancestors<Table>().First() == _table).ToList();
        if (behavior == Word.AutoFitBehavior.Fixed)
        {
            double total = _table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Sum(c => PreferredWidths.ToPoints(c.Width?.Value)) ?? 0;
            AllowAutoFit = false;
            PreferredWidthType = PreferredWidthType.Points;
            if (total > 0)
            {
                PreferredWidth = total;
            }
            return;
        }

        AllowAutoFit = true;
        if (behavior == Word.AutoFitBehavior.Content)
        {
            PreferredWidthType = PreferredWidthType.Auto;
            foreach (TableCell cell in cells)
            {
                cell.TableCellProperties?.RemoveAllChildren<TableCellWidth>();
            }
            return;
        }

        PreferredWidthType = PreferredWidthType.Percent;
        PreferredWidth = 100;
        double totalWidth = cells.Sum(c => TableCellWrapper.Span(c) == 0 ? 0 : new TableCellWrapper(c).Width);
        foreach (TableRow row in _table.Elements<TableRow>())
        {
            double rowWidth = row.Elements<TableCell>().Sum(c => new TableCellWrapper(c).Width);
            foreach (TableCell cell in row.Elements<TableCell>())
            {
                double share = rowWidth > 0 ? new TableCellWrapper(cell).Width / rowWidth * 100 : 100.0 / row.Elements<TableCell>().Count();
                new TableCellWrapper(cell).SetPreferredWidth(PreferredWidthType.Percent, share);
            }
        }
        GC.KeepAlive(totalWidth);
    }

    /// <summary>The distance in points of the left edge of the table from the margin (<c>Rows.LeftIndent</c> in Word).</summary>
    public double LeftIndent
    {
        get { return PreferredWidths.ToPoints(ExistingProperties?.TableIndentation?.Width?.Value); }
        set
        {
            if (double.IsNaN(value) || value < -1584 || value > 1584)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The indent must be between -1584 and 1584 points.");
            }
            GetOrAddTableProperties().TableIndentation = new TableIndentation() { Width = (int)Math.Round(value * 20, MidpointRounding.AwayFromZero), Type = TableWidthUnitValues.Dxa };
        }
    }

    /// <summary>The space between the cells in points (<c>Table.Spacing</c> in Word); 0 when there is none.</summary>
    public double Spacing
    {
        get { return PreferredWidths.ToPoints(ExistingProperties?.TableCellSpacing?.Width?.Value); }
        set
        {
            string twips = PreferredWidths.FromPoints(value, nameof(value));
            if (value == 0)
            {
                ExistingProperties?.RemoveAllChildren<TableCellSpacing>();
                return;
            }
            GetOrAddTableProperties().TableCellSpacing = new TableCellSpacing() { Width = twips, Type = TableWidthUnitValues.Dxa };
        }
    }

    private double GetPadding(Func<TableCellMarginDefault, object> read, double defaultPoints)
    {
        TableCellMarginDefault margins = ExistingProperties?.TableCellMarginDefault;
        object width = margins == null ? null : read(margins);
        return width == null ? defaultPoints : PreferredWidths.ToPoints(width);
    }

    private void SetPadding(double value, Action<TableCellMarginDefault, string> write)
    {
        string twips = PreferredWidths.FromPoints(value, nameof(value));
        TableProperties properties = GetOrAddTableProperties();
        if (properties.TableCellMarginDefault == null)
        {
            properties.TableCellMarginDefault = new TableCellMarginDefault();
        }
        write(properties.TableCellMarginDefault, twips);
    }

    /// <summary>The space in points between the text and the top of the cells (<c>Table.TopPadding</c> in Word); 0 unless set.</summary>
    public double TopPadding
    {
        get { return GetPadding(m => m.TopMargin?.Width?.Value, 0); }
        set { SetPadding(value, (m, twips) => m.TopMargin = new TopMargin() { Width = twips, Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space in points between the text and the bottom of the cells; 0 unless set.</summary>
    public double BottomPadding
    {
        get { return GetPadding(m => m.BottomMargin?.Width?.Value, 0); }
        set { SetPadding(value, (m, twips) => m.BottomMargin = new BottomMargin() { Width = twips, Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space in points between the text and the left of the cells; 5.4 unless set.</summary>
    public double LeftPadding
    {
        get { return GetPadding(m => (object)m.TableCellLeftMargin?.Width?.Value ?? m.StartMargin?.Width?.Value, 5.4); }
        set
        {
            SetPadding(value, (m, twips) =>
            {
                m.StartMargin = null;
                m.TableCellLeftMargin = new TableCellLeftMargin() { Width = short.Parse(twips, CultureInfo.InvariantCulture), Type = TableWidthValues.Dxa };
            });
        }
    }

    /// <summary>The space in points between the text and the right of the cells; 5.4 unless set.</summary>
    public double RightPadding
    {
        get { return GetPadding(m => (object)m.TableCellRightMargin?.Width?.Value ?? m.EndMargin?.Width?.Value, 5.4); }
        set
        {
            SetPadding(value, (m, twips) =>
            {
                m.EndMargin = null;
                m.TableCellRightMargin = new TableCellRightMargin() { Width = short.Parse(twips, CultureInfo.InvariantCulture), Type = TableWidthValues.Dxa };
            });
        }
    }

    private bool GetLook(int mask, Func<TableLook, OnOffValue> attribute)
    {
        TableLook look = ExistingProperties?.TableLook;
        if (look == null)
        {
            return false;
        }
        OnOffValue value = attribute(look);
        if (value != null && value.HasValue)
        {
            return value.Value;
        }
        string hex = look.Val?.Value;
        return hex != null && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int bits) && (bits & mask) != 0;
    }

    private void SetLook(bool value, int mask, Action<TableLook, bool> write)
    {
        TableProperties properties = GetOrAddTableProperties();
        if (properties.TableLook == null)
        {
            properties.TableLook = new TableLook() { Val = "0000" };
        }
        TableLook look = properties.TableLook;
        write(look, value);

        //The hexadecimal value that older programs read is kept in step with the switches.
        int bits = 0;
        bits |= (look.FirstRow?.Value ?? false) ? 0x0020 : 0;
        bits |= (look.LastRow?.Value ?? false) ? 0x0040 : 0;
        bits |= (look.FirstColumn?.Value ?? false) ? 0x0080 : 0;
        bits |= (look.LastColumn?.Value ?? false) ? 0x0100 : 0;
        bits |= (look.NoHorizontalBand?.Value ?? false) ? 0x0200 : 0;
        bits |= (look.NoVerticalBand?.Value ?? false) ? 0x0400 : 0;
        look.Val = bits.ToString("X4", CultureInfo.InvariantCulture);
    }

    /// <summary>Whether the first row gets the formatting that the table style gives it (<c>Table.ApplyStyleHeadingRows</c> in Word).</summary>
    public bool ApplyStyleHeadingRows
    {
        get { return GetLook(0x0020, l => l.FirstRow); }
        set { SetLook(value, 0x0020, (l, v) => l.FirstRow = v); }
    }

    /// <summary>Whether the last row gets the formatting that the table style gives it.</summary>
    public bool ApplyStyleLastRow
    {
        get { return GetLook(0x0040, l => l.LastRow); }
        set { SetLook(value, 0x0040, (l, v) => l.LastRow = v); }
    }

    /// <summary>Whether the first column gets the formatting that the table style gives it.</summary>
    public bool ApplyStyleFirstColumn
    {
        get { return GetLook(0x0080, l => l.FirstColumn); }
        set { SetLook(value, 0x0080, (l, v) => l.FirstColumn = v); }
    }

    /// <summary>Whether the last column gets the formatting that the table style gives it.</summary>
    public bool ApplyStyleLastColumn
    {
        get { return GetLook(0x0100, l => l.LastColumn); }
        set { SetLook(value, 0x0100, (l, v) => l.LastColumn = v); }
    }

    /// <summary>Whether the rows alternate in formatting as the table style says; true unless banding has been switched off.</summary>
    public bool ApplyStyleRowBands
    {
        get { return ExistingProperties?.TableLook == null || !GetLook(0x0200, l => l.NoHorizontalBand); }
        set { SetLook(!value, 0x0200, (l, v) => l.NoHorizontalBand = v); }
    }

    /// <summary>Whether the columns alternate in formatting as the table style says; true unless banding has been switched off.</summary>
    public bool ApplyStyleColumnBands
    {
        get { return ExistingProperties?.TableLook == null || !GetLook(0x0400, l => l.NoVerticalBand); }
        set { SetLook(!value, 0x0400, (l, v) => l.NoVerticalBand = v); }
    }

    /// <summary>The title of the table for people who cannot see it (<c>Table.Title</c> in Word).</summary>
    public string Title
    {
        get { return ExistingProperties?.TableCaption?.Val?.Value ?? string.Empty; }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                ExistingProperties?.RemoveAllChildren<TableCaption>();
                return;
            }
            GetOrAddTableProperties().TableCaption = new TableCaption() { Val = value };
        }
    }

    /// <summary>The description of the table for people who cannot see it (<c>Table.Descr</c> in Word).</summary>
    public string Descr
    {
        get { return ExistingProperties?.TableDescription?.Val?.Value ?? string.Empty; }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                ExistingProperties?.RemoveAllChildren<TableDescription>();
                return;
            }
            GetOrAddTableProperties().TableDescription = new TableDescription() { Val = value };
        }
    }

    /// <summary>True if every row has the same number of cells and none of them is merged (<c>Table.Uniform</c> in Word).</summary>
    public bool Uniform
    {
        get
        {
            List<TableRow> rows = _table.Elements<TableRow>().ToList();
            if (rows.Count == 0)
            {
                return true;
            }
            int count = rows[0].Elements<TableCell>().Count();
            return rows.All(r => r.Elements<TableCell>().Count() == count
                && r.Elements<TableCell>().All(c => TableCellWrapper.Span(c) == 1 && c.TableCellProperties?.VerticalMerge == null)
                && r.TableRowProperties?.GetFirstChild<GridBefore>() == null && r.TableRowProperties?.GetFirstChild<GridAfter>() == null);
        }
    }

    /// <summary>The shading of the whole table, as <c>Table.Shading</c> in Word.</summary>
    public ShadingWrapper Shading
    {
        get
        {
            return new ShadingWrapper(create =>
            {
                TableProperties properties = create ? GetOrAddTableProperties() : ExistingProperties;
                Shading shading = properties?.GetFirstChild<Shading>();
                if (shading == null && create)
                {
                    shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    properties.Shading = shading;
                }
                return new List<Shading> { shading };
            });
        }
    }

    /// <summary>
    /// Splits the table in two before a row, as <c>Table.Split</c> in Word, and returns the new table (the rows from the row on). A paragraph
    /// is put between the tables, as Word does.
    /// </summary>
    /// <exception cref="ArgumentException">The row is not a row of this table, or is its first row.</exception>
    public TableWrapper Split(TableRowWrapper beforeRow)
    {
        Guard.ThrowIfNull(beforeRow);
        TableRow row = beforeRow.TableRow;
        if (row.Parent != _table)
        {
            throw new ArgumentException("The row is not a row of the table.", nameof(beforeRow));
        }
        if (row.PreviousSibling<TableRow>() == null)
        {
            throw new ArgumentException("A table cannot be split before its first row.", nameof(beforeRow));
        }

        Table second = new Table();
        foreach (OpenXmlElement child in _table.ChildElements.Where(e => e is TableProperties || e is TableGrid))
        {
            second.AppendChild(child.CloneNode(true));
        }
        foreach (TableRow moved in _table.Elements<TableRow>().SkipWhile(r => r != row).ToList())
        {
            moved.Remove();
            second.AppendChild(moved);
        }

        Paragraph separator = new Paragraph();
        _table.InsertAfterSelf(separator);
        separator.InsertAfterSelf(second);
        return new TableWrapper(second);
    }
}

public partial class TableRowWrapper
{
    /// <summary>The horizontal alignment of the row on the page (<c>Row.Alignment</c> in Word). The default is Left.</summary>
    public RowAlignment Alignment
    {
        get
        {
            TableRowAlignmentValues? alignment = _tableRow.TableRowProperties?.GetFirstChild<TableJustification>()?.Val?.Value;
            if (alignment == TableRowAlignmentValues.Center)
            {
                return RowAlignment.Center;
            }
            return alignment == TableRowAlignmentValues.Right ? RowAlignment.Right : RowAlignment.Left;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            TableRowProperties properties = GetOrAddTableRowProperties();
            properties.RemoveAllChildren<TableJustification>();
            if (value != RowAlignment.Left)
            {
                properties.AppendChild(new TableJustification() { Val = value == RowAlignment.Center ? TableRowAlignmentValues.Center : TableRowAlignmentValues.Right });
            }
        }
    }

    /// <summary>Sets the height and its rule in one go, as <c>Row.SetHeight</c> in Word.</summary>
    public void SetHeight(double rowHeight, RowHeightRule heightRule)
    {
        Height = rowHeight;
        HeightRule = heightRule;
    }

    /// <summary>The shading of the cells of the row, as <c>Row.Shading</c> in Word.</summary>
    public ShadingWrapper Shading
    {
        get { return new ShadingWrapper(create => _tableRow.Elements<TableCell>().Select(c => new TableCellWrapper(c).GetShading(create)).ToList()); }
    }
}

public partial class TableRowWrapperCollection
{
    private IList<TableRowWrapper> Items
    {
        get { return this.ToList(); }
    }

    private T? ReadAll<T>(Func<TableRowWrapper, T> read) where T : struct
    {
        IList<TableRowWrapper> rows = Items;
        if (rows.Count == 0)
        {
            return null;
        }
        T first = read(rows[0]);
        return rows.All(r => EqualityComparer<T>.Default.Equals(read(r), first)) ? first : (T?)null;
    }

    private void WriteAll(Action<TableRowWrapper> write)
    {
        foreach (TableRowWrapper row in Items)
        {
            write(row);
        }
    }

    /// <summary>The first row, or null if there are none.</summary>
    public TableRowWrapper First
    {
        get { return Items.FirstOrDefault(); }
    }

    /// <summary>The last row, or null if there are none.</summary>
    public TableRowWrapper Last
    {
        get { return Items.LastOrDefault(); }
    }

    /// <summary>The alignment of the table on the page, as <c>Rows.Alignment</c> in Word.</summary>
    public RowAlignment Alignment
    {
        get { return Parent.Alignment; }
        set { Parent.Alignment = value; }
    }

    /// <summary>The distance in points of the left edge of the table from the margin, as <c>Rows.LeftIndent</c> in Word.</summary>
    public double LeftIndent
    {
        get { return Parent.LeftIndent; }
        set { Parent.LeftIndent = value; }
    }

    /// <summary>The height of the rows in points; null when they differ. Setting it sets every row.</summary>
    public double? Height
    {
        get { return ReadAll(r => r.Height); }
        set
        {
            if (value.HasValue)
            {
                WriteAll(r => r.Height = value.Value);
            }
        }
    }

    /// <summary>How the height of the rows is applied; null when they differ.</summary>
    public RowHeightRule? HeightRule
    {
        get { return ReadAll(r => r.HeightRule); }
        set
        {
            if (value.HasValue)
            {
                WriteAll(r => r.HeightRule = value.Value);
            }
        }
    }

    /// <summary>Whether the rows repeat at the top of each page; null when they differ.</summary>
    public bool? HeadingFormat
    {
        get { return ReadAll(r => r.HeadingFormat); }
        set
        {
            if (value.HasValue)
            {
                WriteAll(r => r.HeadingFormat = value.Value);
            }
        }
    }

    /// <summary>Whether the rows can break across pages; null when they differ.</summary>
    public bool? AllowBreakAcrossPages
    {
        get { return ReadAll(r => r.AllowBreakAcrossPages); }
        set
        {
            if (value.HasValue)
            {
                WriteAll(r => r.AllowBreakAcrossPages = value.Value);
            }
        }
    }

    /// <summary>The shading of the cells of all the rows.</summary>
    public ShadingWrapper Shading
    {
        get { return new ShadingWrapper(create => Items.SelectMany(r => r.TableRow.Elements<TableCell>()).Select(c => new TableCellWrapper(c).GetShading(create)).ToList()); }
    }
}

public partial class TableCellWrapper
{
    internal Shading GetShading(bool create)
    {
        TableCellProperties properties = create ? GetOrAddTableCellProperties() : _tableCell.TableCellProperties;
        Shading shading = properties?.Shading;
        if (shading == null && create)
        {
            shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
            properties.Shading = shading;
        }
        return shading;
    }

    /// <summary>The shading of the cell, as <c>Cell.Shading</c> in Word.</summary>
    public ShadingWrapper Shading
    {
        get { return new ShadingWrapper(create => new List<Shading> { GetShading(create) }); }
    }

    /// <summary>The row the cell is in (<c>Cell.Row</c> in Word).</summary>
    public TableRowWrapper Row
    {
        get { return Parent; }
    }

    /// <summary>The column of the table grid where the cell starts (<c>Cell.Column</c> in Word).</summary>
    public TableColumnWrapper Column
    {
        get
        {
            Table table = _tableCell.Ancestors<Table>().First();
            GridColumn column = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ElementAtOrDefault(GetGridStart(_tableCell));
            return column == null ? null : new TableColumnWrapper(column);
        }
    }

    /// <summary>The tables inside the cell (<c>Cell.Tables</c> in Word).</summary>
    public TableWrapperCollection Tables
    {
        get { return new TableWrapperCollection(new RangeWrapper(_tableCell)); }
    }

    /// <summary>The height of the row of the cell in points (<c>Cell.Height</c> in Word); the same as <c>Row.Height</c>.</summary>
    public double Height
    {
        get { return Parent.Height; }
        set { Parent.Height = value; }
    }

    /// <summary>How the height of the row of the cell is applied (<c>Cell.HeightRule</c> in Word).</summary>
    public RowHeightRule HeightRule
    {
        get { return Parent.HeightRule; }
        set { Parent.HeightRule = value; }
    }

    /// <summary>Whether text wraps in the cell; true unless it has been switched off (<c>Cell.WordWrap</c> in Word).</summary>
    public bool WordWrap
    {
        get { return _tableCell.TableCellProperties?.NoWrap == null; }
        set
        {
            if (value)
            {
                _tableCell.TableCellProperties?.RemoveAllChildren<NoWrap>();
            }
            else
            {
                GetOrAddTableCellProperties().NoWrap = new NoWrap();
            }
        }
    }

    /// <summary>Whether the text is squeezed to fit the width of the cell (<c>Cell.FitText</c> in Word).</summary>
    public bool FitText
    {
        get { return _tableCell.TableCellProperties?.TableCellFitText != null; }
        set
        {
            if (value)
            {
                GetOrAddTableCellProperties().TableCellFitText = new TableCellFitText();
            }
            else
            {
                _tableCell.TableCellProperties?.RemoveAllChildren<TableCellFitText>();
            }
        }
    }

    /// <summary>What <see cref="PreferredWidth"/> means: points, a percentage or automatic.</summary>
    public PreferredWidthType PreferredWidthType
    {
        get
        {
            TableCellWidth width = _tableCell.TableCellProperties?.TableCellWidth;
            return PreferredWidths.GetKind(width?.Width?.Value, width?.Type?.Value);
        }
        set { SetPreferredWidth(value, PreferredWidth); }
    }

    /// <summary>The preferred width of the cell, in points or in percent according to <see cref="PreferredWidthType"/>; 0 when automatic.</summary>
    public double PreferredWidth
    {
        get
        {
            TableCellWidth width = _tableCell.TableCellProperties?.TableCellWidth;
            return PreferredWidths.GetValue(width?.Width?.Value, width?.Type?.Value);
        }
        set { SetPreferredWidth(PreferredWidthType == PreferredWidthType.Percent ? PreferredWidthType.Percent : PreferredWidthType.Points, value); }
    }

    internal void SetPreferredWidth(PreferredWidthType kind, double value)
    {
        PreferredWidths.Format(kind, value, out string width, out TableWidthUnitValues type);
        GetOrAddTableCellProperties().TableCellWidth = new TableCellWidth() { Width = width, Type = type };
    }

    private double GetPadding(Func<TableCellMargin, object> read, Func<TableWrapper, double> inherited)
    {
        TableCellMargin margins = _tableCell.TableCellProperties?.TableCellMargin;
        object width = margins == null ? null : read(margins);
        return width == null ? inherited(Parent.Parent) : PreferredWidths.ToPoints(width);
    }

    private void SetPadding(double value, Action<TableCellMargin, string> write)
    {
        string twips = PreferredWidths.FromPoints(value, nameof(value));
        TableCellProperties properties = GetOrAddTableCellProperties();
        if (properties.TableCellMargin == null)
        {
            properties.TableCellMargin = new TableCellMargin();
        }
        write(properties.TableCellMargin, twips);
    }

    /// <summary>The space in points between the text and the top of the cell; that of the table unless set.</summary>
    public double TopPadding
    {
        get { return GetPadding(m => m.TopMargin?.Width?.Value, t => t.TopPadding); }
        set { SetPadding(value, (m, twips) => m.TopMargin = new TopMargin() { Width = twips, Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space in points between the text and the bottom of the cell; that of the table unless set.</summary>
    public double BottomPadding
    {
        get { return GetPadding(m => m.BottomMargin?.Width?.Value, t => t.BottomPadding); }
        set { SetPadding(value, (m, twips) => m.BottomMargin = new BottomMargin() { Width = twips, Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space in points between the text and the left of the cell; that of the table unless set.</summary>
    public double LeftPadding
    {
        get { return GetPadding(m => (object)m.LeftMargin?.Width?.Value ?? m.StartMargin?.Width?.Value, t => t.LeftPadding); }
        set
        {
            SetPadding(value, (m, twips) =>
            {
                m.StartMargin = null;
                m.LeftMargin = new LeftMargin() { Width = twips, Type = TableWidthUnitValues.Dxa };
            });
        }
    }

    /// <summary>The space in points between the text and the right of the cell; that of the table unless set.</summary>
    public double RightPadding
    {
        get { return GetPadding(m => (object)m.RightMargin?.Width?.Value ?? m.EndMargin?.Width?.Value, t => t.RightPadding); }
        set
        {
            SetPadding(value, (m, twips) =>
            {
                m.EndMargin = null;
                m.RightMargin = new RightMargin() { Width = twips, Type = TableWidthUnitValues.Dxa };
            });
        }
    }

    /// <summary>
    /// Deletes the cell, as <c>Cell.Delete</c> in Word. With <see cref="ShiftCells.Left"/> the cells after it in the row move left and the row is
    /// left short; deleting the only cell of a row deletes the row.
    /// </summary>
    /// <exception cref="NotSupportedException"><see cref="ShiftCells.Up"/>, or a cell that is part of a vertical merge.</exception>
    public void Delete(ShiftCells shiftCells = ShiftCells.Left)
    {
        if (!EnumCompat.IsDefined(shiftCells))
        {
            throw new ArgumentOutOfRangeException(nameof(shiftCells));
        }
        if (shiftCells == ShiftCells.EntireRow)
        {
            Parent.Delete();
            return;
        }
        if (shiftCells == ShiftCells.EntireColumn)
        {
            Column?.Delete();
            return;
        }
        if (shiftCells == ShiftCells.Up)
        {
            throw new NotSupportedException("Shifting cells up is not supported.");
        }
        if (_tableCell.TableCellProperties?.VerticalMerge != null)
        {
            throw new NotSupportedException("Deleting a cell that is part of a vertical merge is not supported.");
        }

        TableRow row = (TableRow)_tableCell.Parent;
        if (row.Elements<TableCell>().Count() == 1)
        {
            Parent.Delete();
            return;
        }

        int span = Span(_tableCell);
        int width = WidthTwips;
        TableRowWrapper rowWrapper = Parent;
        _tableCell.Remove();

        //The grid does not change, so the row is shorter by the cell's columns.
        TableRowProperties properties = rowWrapper.GetOrAddProperties();
        GridAfter after = properties.GetFirstChild<GridAfter>();
        if (after == null)
        {
            properties.AppendChild(new GridAfter() { Val = span });
        }
        else
        {
            after.Val = (after.Val?.Value ?? 0) + span;
        }
        WidthAfterTableRow widthAfter = properties.GetFirstChild<WidthAfterTableRow>();
        if (widthAfter == null)
        {
            properties.AppendChild(new WidthAfterTableRow() { Width = width.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
        }
        else
        {
            widthAfter.Width = ((int.TryParse(widthAfter.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int existing) ? existing : 0) + width).ToString(CultureInfo.InvariantCulture);
        }
    }
}

public partial class TableRowWrapper
{
    internal TableRowProperties GetOrAddProperties()
    {
        return GetOrAddTableRowProperties();
    }
}

/// <summary>
/// The cells of a column, as <c>Column.Cells</c> in Word.
/// </summary>
public class TableColumnCellCollection : IEnumerable<TableCellWrapper>
{
    private readonly GridColumn _gridColumn;

    internal TableColumnCellCollection(GridColumn gridColumn)
    {
        _gridColumn = gridColumn;
    }

    private IList<TableCell> GetElementList()
    {
        Table table = _gridColumn.Ancestors<Table>().First();
        int index = _gridColumn.ElementsBefore().Count(e => e is GridColumn);
        return table.Elements<TableRow>().Select(r => TableCellWrapper.CellAtGridColumn(r, index)).Where(c => c != null).ToList();
    }

    /// <summary>The number of cells: one for each row that has a cell in the column.</summary>
    public int Count
    {
        get { return GetElementList().Count; }
    }

    /// <summary>The cell at the position (the row) in the column.</summary>
    public TableCellWrapper this[int index]
    {
        get
        {
            IList<TableCell> cells = GetElementList();
            if (index < 0 || index >= cells.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new TableCellWrapper(cells[index]);
        }
    }

    public IEnumerator<TableCellWrapper> GetEnumerator()
    {
        foreach (TableCell cell in GetElementList())
        {
            yield return new TableCellWrapper(cell);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

public partial class TableColumnWrapper
{
    /// <summary>The cells of the column, one in each row (<c>Column.Cells</c> in Word).</summary>
    public TableColumnCellCollection Cells
    {
        get { return new TableColumnCellCollection(_gridColumn); }
    }

    /// <summary>True for the first column of the table.</summary>
    public bool IsFirst
    {
        get { return _gridColumn.PreviousSibling<GridColumn>() == null; }
    }

    /// <summary>True for the last column of the table.</summary>
    public bool IsLast
    {
        get { return _gridColumn.NextSibling<GridColumn>() == null; }
    }

    /// <summary>The shading of the cells of the column, as <c>Column.Shading</c> in Word.</summary>
    public ShadingWrapper Shading
    {
        get { return new ShadingWrapper(create => Cells.Select(c => c.GetShading(create)).ToList()); }
    }

    /// <summary>What <see cref="PreferredWidth"/> means: points, a percentage or automatic; of the first cell when they differ.</summary>
    public PreferredWidthType PreferredWidthType
    {
        get { return Cells.FirstOrDefault()?.PreferredWidthType ?? PreferredWidthType.Auto; }
        set
        {
            foreach (TableCellWrapper cell in Cells.Where(c => TableCellWrapper.Span(c.TableCell) == 1))
            {
                cell.PreferredWidthType = value;
            }
        }
    }

    /// <summary>The preferred width of the cells of the column (those that span more than one column are left alone).</summary>
    public double PreferredWidth
    {
        get { return Cells.FirstOrDefault()?.PreferredWidth ?? 0; }
        set
        {
            PreferredWidthType kind = PreferredWidthType == PreferredWidthType.Percent ? PreferredWidthType.Percent : PreferredWidthType.Points;
            foreach (TableCellWrapper cell in Cells.Where(c => TableCellWrapper.Span(c.TableCell) == 1))
            {
                cell.SetPreferredWidth(kind, value);
            }
        }
    }
}

public partial class TableColumnWrapperCollection
{
    private IList<TableColumnWrapper> Items
    {
        get { return this.ToList(); }
    }

    /// <summary>The first column, or null if there are none.</summary>
    public TableColumnWrapper First
    {
        get { return Items.FirstOrDefault(); }
    }

    /// <summary>The last column, or null if there are none.</summary>
    public TableColumnWrapper Last
    {
        get { return Items.LastOrDefault(); }
    }

    /// <summary>
    /// Adds a column before another one, as <c>Columns.Add</c> in Word: the new column has the width of that column and an empty cell in every
    /// row. A cell that spans across the position is made one column wider instead. With null the column is added at the end.
    /// </summary>
    /// <exception cref="ArgumentException">The column is not a column of this table.</exception>
    public TableColumnWrapper Add(TableColumnWrapper beforeColumn)
    {
        if (beforeColumn == null)
        {
            return Add();
        }

        Table table = _range.StartElement as Table;
        if (table == null)
        {
            throw new InvalidOperationException("Columns can only be added to a collection of the columns of a table.");
        }
        GridColumn before = beforeColumn.GridColumn;
        if (before.Ancestors<Table>().FirstOrDefault() != table)
        {
            throw new ArgumentException("The column is not a column of the table.", nameof(beforeColumn));
        }

        int index = before.ElementsBefore().Count(e => e is GridColumn);
        GridColumn gridColumn = (GridColumn)before.CloneNode(true);
        before.InsertBeforeSelf(gridColumn);
        int gridWidth = gridColumn.Width == null ? 0 : Convert.ToInt32(gridColumn.Width.Value, CultureInfo.InvariantCulture);

        foreach (TableRow row in table.Elements<TableRow>())
        {
            TableCell cell = TableCellWrapper.CellAtGridColumn(row, index);
            if (cell == null)
            {
                continue;
            }

            if (TableCellWrapper.GetGridStart(cell) < index)
            {
                //The cell spans across the new column, so it gets wider.
                new TableCellWrapper(cell).WidenBy(1, gridWidth);
                continue;
            }

            TableCell added = (TableCell)cell.Clone();
            added.RemoveAnchors();
            added.ClearText();
            added.SetGridSpan(1);
            added.TableCellProperties?.RemoveAllChildren<VerticalMerge>();
            if (gridWidth > 0)
            {
                added.SetWidth(gridWidth, TableWidthUnitValues.Dxa);
            }
            cell.InsertBeforeSelf(added);
        }

        return new TableColumnWrapper(gridColumn);
    }

    /// <summary>
    /// Makes the columns the same width, as <c>Columns.DistributeWidth</c> in Word: the width of the table is shared out evenly.
    /// </summary>
    public void DistributeWidth()
    {
        IList<TableColumnWrapper> columns = Items;
        if (columns.Count == 0)
        {
            return;
        }

        double total = columns.Sum(c => c.Width);
        double each = Math.Round(total * 20 / columns.Count, MidpointRounding.AwayFromZero) / 20;
        foreach (TableColumnWrapper column in columns)
        {
            column.Width = each;
        }
    }
}

public partial class TableCellWrapper
{
    //Makes a cell that spans columns wider by more columns of the given width.
    internal void WidenBy(int columns, int columnWidthTwips)
    {
        GetOrAddTableCellProperties().GridSpan = new GridSpan() { Val = Span(_tableCell) + columns };
        if (WidthTwips > 0)
        {
            WidthTwips += columns * columnWidthTwips;
        }
    }
}

public partial class TableColumnWrapper
{
    /// <summary>
    /// The borders of the column, as <c>Column.Borders</c> in Word. They are stored on the cells of the column: <c>Left</c> and <c>Right</c> on
    /// every cell, <c>Top</c> on the first, <c>Bottom</c> on the last, and <c>Horizontal</c> between the cells. A column has no vertical or
    /// diagonal border.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return new BordersWrapper(_gridColumn, false); }
    }
}
