using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The part of a table that a conditional format of a table style is for, as <c>WdConditionCode</c> in Word (with the same values).
/// </summary>
public enum ConditionCode
{
    FirstRow = 0,
    LastRow = 1,
    OddRowBanding = 2,
    EvenRowBanding = 3,
    FirstColumn = 4,
    LastColumn = 5,
    OddColumnBanding = 6,
    EvenColumnBanding = 7,
    NorthEastCell = 8,
    NorthWestCell = 9,
    SouthEastCell = 10,
    SouthWestCell = 11,
}

/// <summary>
/// The table properties of a table style, as <c>Style.Table</c> in Word.
/// </summary>
/// <remarks>
/// The markup is what Word 16 writes: the alignment, the space between cells and the break across pages are written to the table and row
/// properties of the style, the paddings and the shading to the cell properties (as Word does), and the borders to the table borders.
/// Reading gives only what the style itself sets, not what it inherits.
/// </remarks>
public class TableStyleWrapper
{
    private static readonly TableStyleOverrideValues[] ConditionOrder =
    {
        TableStyleOverrideValues.WholeTable, TableStyleOverrideValues.FirstRow, TableStyleOverrideValues.LastRow, TableStyleOverrideValues.FirstColumn,
        TableStyleOverrideValues.LastColumn, TableStyleOverrideValues.Band1Vertical, TableStyleOverrideValues.Band2Vertical,
        TableStyleOverrideValues.Band1Horizontal, TableStyleOverrideValues.Band2Horizontal, TableStyleOverrideValues.NorthEastCell,
        TableStyleOverrideValues.NorthWestCell, TableStyleOverrideValues.SouthEastCell, TableStyleOverrideValues.SouthWestCell,
    };

    private readonly Style _style;

    internal TableStyleWrapper(Style style)
    {
        _style = style;
    }

    private StyleTableProperties GetTableProperties(bool create)
    {
        if (_style.StyleTableProperties == null && create)
        {
            _style.StyleTableProperties = new StyleTableProperties();
        }
        return _style.StyleTableProperties;
    }

    private TableStyleConditionalFormattingTableRowProperties GetRowProperties(bool create)
    {
        if (_style.TableStyleConditionalFormattingTableRowProperties == null && create)
        {
            _style.TableStyleConditionalFormattingTableRowProperties = new TableStyleConditionalFormattingTableRowProperties();
        }
        return _style.TableStyleConditionalFormattingTableRowProperties;
    }

    private StyleTableCellProperties GetCellProperties(bool create)
    {
        if (_style.StyleTableCellProperties == null && create)
        {
            _style.StyleTableCellProperties = new StyleTableCellProperties();
        }
        return _style.StyleTableCellProperties;
    }

    /// <summary>The horizontal alignment of tables in the style. The default is Left.</summary>
    public RowAlignment Alignment
    {
        get
        {
            TableRowAlignmentValues? value = GetTableProperties(false)?.TableJustification?.Val?.Value;
            return value == TableRowAlignmentValues.Center ? RowAlignment.Center : value == TableRowAlignmentValues.Right ? RowAlignment.Right : RowAlignment.Left;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            TableRowAlignmentValues alignment = value == RowAlignment.Center ? TableRowAlignmentValues.Center : TableRowAlignmentValues.Right;
            if (value == RowAlignment.Left)
            {
                GetTableProperties(false)?.RemoveAllChildren<TableJustification>();
                GetRowProperties(false)?.RemoveAllChildren<TableJustification>();
            }
            else
            {
                GetTableProperties(true).TableJustification = new TableJustification() { Val = alignment };
                SetRowProperty(new TableJustification() { Val = alignment });
            }
        }
    }

    /// <summary>Whether a row can break across pages. The default is true.</summary>
    public bool AllowBreakAcrossPage
    {
        get { return GetRowProperties(false)?.GetFirstChild<CantSplit>() == null; }
        set
        {
            GetRowProperties(false)?.RemoveAllChildren<CantSplit>();
            if (!value)
            {
                SetRowProperty(new CantSplit());
            }
        }
    }

    /// <summary>The space between cells in points. The default is 0.</summary>
    public double Spacing
    {
        get { return PreferredWidths.ToPoints(GetTableProperties(false)?.TableCellSpacing?.Width?.Value); }
        set
        {
            Guard.ThrowIfNegative(value);
            GetTableProperties(false)?.RemoveAllChildren<TableCellSpacing>();
            GetRowProperties(false)?.RemoveAllChildren<TableCellSpacing>();
            if (value > 0)
            {
                string twips = Twips(value);
                GetTableProperties(true).TableCellSpacing = new TableCellSpacing() { Width = twips, Type = TableWidthUnitValues.Dxa };
                SetRowProperty(new TableCellSpacing() { Width = twips, Type = TableWidthUnitValues.Dxa });
            }
        }
    }

    /// <summary>How far the tables are indented from the left margin, in points.</summary>
    public double LeftIndent
    {
        get { return PreferredWidths.ToPoints(GetTableProperties(false)?.TableIndentation?.Width?.Value); }
        set { GetTableProperties(true).TableIndentation = new TableIndentation() { Width = (int)Math.Round(value * 20, MidpointRounding.AwayFromZero), Type = TableWidthUnitValues.Dxa }; }
    }

    /// <summary>The number of rows in a band of rows that are formatted alike. The default is 1.</summary>
    public int RowStripe
    {
        get { return GetTableProperties(false)?.TableStyleRowBandSize?.Val?.Value ?? 1; }
        set
        {
            Guard.ThrowIfLessThan(value, 1);
            GetTableProperties(true).TableStyleRowBandSize = new TableStyleRowBandSize() { Val = (short)value };
        }
    }

    /// <summary>The number of columns in a band of columns that are formatted alike. The default is 1.</summary>
    public int ColumnStripe
    {
        get { return GetTableProperties(false)?.TableStyleColumnBandSize?.Val?.Value ?? 1; }
        set
        {
            Guard.ThrowIfLessThan(value, 1);
            GetTableProperties(true).TableStyleColumnBandSize = new TableStyleColumnBandSize() { Val = (short)value };
        }
    }

    /// <summary>The borders of tables in the style: the four edges and the lines between rows and columns.</summary>
    public BordersWrapper Borders
    {
        get { return BordersWrapper.ForTableStyle(_style); }
    }

    /// <summary>The shading of the cells of tables in the style.</summary>
    public ShadingWrapper Shading
    {
        get
        {
            return new ShadingWrapper(create =>
            {
                StyleTableCellProperties properties = GetCellProperties(create);
                Shading shading = properties?.Shading;
                if (shading == null && create)
                {
                    shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    properties.Shading = shading;
                }
                return new List<Shading> { shading };
            });
        }
    }

    /// <summary>The space between the text and the top of the cells, in points.</summary>
    public double TopPadding
    {
        get { return GetPadding(m => m.TopMargin?.Width?.Value, m => m.TopMargin?.Type?.Value); }
        set { SetPadding(value, m => m.TopMargin = new TopMargin() { Width = Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space between the text and the bottom of the cells, in points.</summary>
    public double BottomPadding
    {
        get { return GetPadding(m => m.BottomMargin?.Width?.Value, m => m.BottomMargin?.Type?.Value); }
        set { SetPadding(value, m => m.BottomMargin = new BottomMargin() { Width = Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space between the text and the left of the cells, in points.</summary>
    public double LeftPadding
    {
        get { return GetPadding(m => (m.LeftMargin?.Width ?? m.StartMargin?.Width)?.Value, m => (m.LeftMargin?.Type ?? m.StartMargin?.Type)?.Value); }
        set { SetPadding(value, m => m.LeftMargin = new LeftMargin() { Width = Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space between the text and the right of the cells, in points.</summary>
    public double RightPadding
    {
        get { return GetPadding(m => (m.RightMargin?.Width ?? m.EndMargin?.Width)?.Value, m => (m.RightMargin?.Type ?? m.EndMargin?.Type)?.Value); }
        set { SetPadding(value, m => m.RightMargin = new RightMargin() { Width = Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>
    /// The format of a part of the tables, such as the first row (<c>TableStyle.Condition</c>). It is made when something is set on it.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The code is not one of <see cref="ConditionCode"/>.</exception>
    public ConditionalStyleWrapper Condition(ConditionCode code)
    {
        if (!EnumCompat.IsDefined(code))
        {
            throw new ArgumentOutOfRangeException(nameof(code));
        }
        return new ConditionalStyleWrapper(_style, ToOverride(code));
    }

    private double GetPadding(Func<TableCellMargin, string> width, Func<TableCellMargin, TableWidthUnitValues?> type)
    {
        TableCellMargin margin = GetCellProperties(false)?.TableCellMargin;
        string twips = margin == null ? null : width(margin);
        TableWidthUnitValues? unit = margin == null ? null : type(margin);
        return twips != null && unit != TableWidthUnitValues.Nil ? PreferredWidths.ToPoints(twips) : 0;
    }

    private void SetPadding(double value, Action<TableCellMargin> write)
    {
        Guard.ThrowIfNegative(value);
        StyleTableCellProperties properties = GetCellProperties(true);
        if (properties.TableCellMargin == null)
        {
            properties.TableCellMargin = new TableCellMargin();
        }
        write(properties.TableCellMargin);
    }

    private void SetRowProperty(OpenXmlElement property)
    {
        TableStyleConditionalFormattingTableRowProperties properties = GetRowProperties(true);
        foreach (OpenXmlElement existing in properties.ChildElements.Where(e => e.LocalName == property.LocalName).ToList())
        {
            existing.Remove();
        }

        //The order Word writes them in.
        string[] order = { "cantSplit", "tblCellSpacing", "jc" };
        int rank = Array.IndexOf(order, property.LocalName);
        OpenXmlElement next = properties.ChildElements.FirstOrDefault(e => Array.IndexOf(order, e.LocalName) > rank);
        if (next != null)
        {
            next.InsertBeforeSelf(property);
        }
        else
        {
            properties.AppendChild(property);
        }
    }

    internal static string Twips(double points)
    {
        return ((long)Math.Round(points * 20, MidpointRounding.AwayFromZero)).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static TableStyleOverrideValues ToOverride(ConditionCode code)
    {
        switch (code)
        {
            case ConditionCode.FirstRow: return TableStyleOverrideValues.FirstRow;
            case ConditionCode.LastRow: return TableStyleOverrideValues.LastRow;
            case ConditionCode.OddRowBanding: return TableStyleOverrideValues.Band1Horizontal;
            case ConditionCode.EvenRowBanding: return TableStyleOverrideValues.Band2Horizontal;
            case ConditionCode.FirstColumn: return TableStyleOverrideValues.FirstColumn;
            case ConditionCode.LastColumn: return TableStyleOverrideValues.LastColumn;
            case ConditionCode.OddColumnBanding: return TableStyleOverrideValues.Band1Vertical;
            case ConditionCode.EvenColumnBanding: return TableStyleOverrideValues.Band2Vertical;
            case ConditionCode.NorthEastCell: return TableStyleOverrideValues.NorthEastCell;
            case ConditionCode.NorthWestCell: return TableStyleOverrideValues.NorthWestCell;
            case ConditionCode.SouthEastCell: return TableStyleOverrideValues.SouthEastCell;
            default: return TableStyleOverrideValues.SouthWestCell;
        }
    }

    //The conditional format of the style for a part of the table, made if asked to, among the others in the order Word writes them.
    internal static TableStyleProperties GetConditional(Style style, TableStyleOverrideValues type, bool create)
    {
        TableStyleProperties found = style.Elements<TableStyleProperties>().FirstOrDefault(c => c.Type != null && c.Type.HasValue && c.Type.Value == type);
        if (found != null || !create)
        {
            return found;
        }

        found = new TableStyleProperties() { Type = type };
        int rank = Array.IndexOf(ConditionOrder, type);
        TableStyleProperties next = style.Elements<TableStyleProperties>().FirstOrDefault(c => c.Type != null && c.Type.HasValue && Array.IndexOf(ConditionOrder, c.Type.Value) > rank);
        if (next != null)
        {
            next.InsertBeforeSelf(found);
        }
        else
        {
            style.AppendChild(found);
        }
        return found;
    }
}

/// <summary>
/// The format of a part of the tables of a table style, such as the first row, as <c>ConditionalStyle</c> in Word.
/// </summary>
/// <remarks>
/// Setting the shading, a border or a padding writes what Word writes with it: an empty table properties element, and cell properties in
/// which the borders that are not set are none. A padding that is set leaves the others none. Reading gives only what the conditional
/// format itself sets.
/// </remarks>
public class ConditionalStyleWrapper
{
    private readonly Style _style;
    private readonly TableStyleOverrideValues _type;

    internal ConditionalStyleWrapper(Style style, TableStyleOverrideValues type)
    {
        _style = style;
        _type = type;
    }

    private TableStyleProperties Get(bool create)
    {
        return TableStyleWrapper.GetConditional(_style, _type, create);
    }

    /// <summary>The character formatting for the part of the table. Only what is set directly is read.</summary>
    public FontWrapper Font
    {
        get { return new FontWrapper(_style, Get); }
    }

    /// <summary>The paragraph formatting for the part of the table. Only what is set directly is read.</summary>
    public ParagraphFormatWrapper ParagraphFormat
    {
        get { return new ParagraphFormatWrapper(_style, Get); }
    }

    /// <summary>The shading of the cells of the part of the table.</summary>
    public ShadingWrapper Shading
    {
        get
        {
            return new ShadingWrapper(create =>
            {
                TableStyleConditionalFormattingTableCellProperties cells = GetCells(create);
                Shading shading = cells?.Shading;
                if (shading == null && create)
                {
                    //Word writes the borders as none beside it, so that the shading shows on its own.
                    if (cells.TableCellBorders == null)
                    {
                        cells.TableCellBorders = NoBorders();
                    }
                    shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    cells.Shading = shading;
                }
                return new List<Shading> { shading };
            });
        }
    }

    //Cell borders that are all none, which is what Word makes of the borders of a conditional format when it writes any of them.
    internal static TableCellBorders NoBorders()
    {
        return new TableCellBorders(
            new TopBorder() { Val = BorderValues.Nil },
            new LeftBorder() { Val = BorderValues.Nil },
            new BottomBorder() { Val = BorderValues.Nil },
            new RightBorder() { Val = BorderValues.Nil },
            new InsideHorizontalBorder() { Val = BorderValues.Nil },
            new InsideVerticalBorder() { Val = BorderValues.Nil });
    }

    /// <summary>The borders of the cells of the part of the table: the four edges and the lines between cells.</summary>
    public BordersWrapper Borders
    {
        get { return BordersWrapper.ForConditionalStyle(GetCells); }
    }

    /// <summary>The space between the text and the top of the cells, in points.</summary>
    public double TopPadding
    {
        get { return GetPadding(m => m.TopMargin?.Width?.Value, m => m.TopMargin?.Type?.Value); }
        set { SetPadding(value, m => m.TopMargin = new TopMargin() { Width = TableStyleWrapper.Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space between the text and the bottom of the cells, in points.</summary>
    public double BottomPadding
    {
        get { return GetPadding(m => m.BottomMargin?.Width?.Value, m => m.BottomMargin?.Type?.Value); }
        set { SetPadding(value, m => m.BottomMargin = new BottomMargin() { Width = TableStyleWrapper.Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space between the text and the left of the cells, in points.</summary>
    public double LeftPadding
    {
        get { return GetPadding(m => (m.LeftMargin?.Width ?? m.StartMargin?.Width)?.Value, m => (m.LeftMargin?.Type ?? m.StartMargin?.Type)?.Value); }
        set { SetPadding(value, m => m.LeftMargin = new LeftMargin() { Width = TableStyleWrapper.Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    /// <summary>The space between the text and the right of the cells, in points.</summary>
    public double RightPadding
    {
        get { return GetPadding(m => (m.RightMargin?.Width ?? m.EndMargin?.Width)?.Value, m => (m.RightMargin?.Type ?? m.EndMargin?.Type)?.Value); }
        set { SetPadding(value, m => m.RightMargin = new RightMargin() { Width = TableStyleWrapper.Twips(value), Type = TableWidthUnitValues.Dxa }); }
    }

    //The cell properties of the conditional format, made if asked to. Making them also makes the empty table properties Word writes with them.
    internal TableStyleConditionalFormattingTableCellProperties GetCells(bool create)
    {
        TableStyleProperties conditional = Get(create);
        if (conditional == null)
        {
            return null;
        }
        if (conditional.TableStyleConditionalFormattingTableCellProperties == null && create)
        {
            if (conditional.TableStyleConditionalFormattingTableProperties == null)
            {
                conditional.TableStyleConditionalFormattingTableProperties = new TableStyleConditionalFormattingTableProperties();
            }
            conditional.TableStyleConditionalFormattingTableCellProperties = new TableStyleConditionalFormattingTableCellProperties();
        }
        return conditional.TableStyleConditionalFormattingTableCellProperties;
    }

    private double GetPadding(Func<TableCellMargin, string> width, Func<TableCellMargin, TableWidthUnitValues?> type)
    {
        TableCellMargin margin = GetCells(false)?.TableCellMargin;
        string twips = margin == null ? null : width(margin);
        TableWidthUnitValues? unit = margin == null ? null : type(margin);
        return twips != null && unit != TableWidthUnitValues.Nil ? PreferredWidths.ToPoints(twips) : 0;
    }

    //Word makes the margins that are not set none.
    private void SetPadding(double value, Action<TableCellMargin> write)
    {
        Guard.ThrowIfNegative(value);
        TableStyleConditionalFormattingTableCellProperties cells = GetCells(true);
        if (cells.TableCellMargin == null)
        {
            cells.TableCellMargin = new TableCellMargin()
            {
                TopMargin = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Nil },
                LeftMargin = new LeftMargin() { Width = "0", Type = TableWidthUnitValues.Nil },
                BottomMargin = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Nil },
                RightMargin = new RightMargin() { Width = "0", Type = TableWidthUnitValues.Nil },
            };
        }
        write(cells.TableCellMargin);
    }
}
