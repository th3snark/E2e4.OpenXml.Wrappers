using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A paper size, as <c>WdPaperSize</c> in Word (with the same values).
/// </summary>
public enum PaperSize
{
    Paper11x17 = 1,
    Letter = 2,
    Legal = 4,
    Executive = 5,
    A3 = 6,
    A4 = 7,
    A5 = 9,
    B4 = 10,
    B5 = 11,
    Folio = 18,
    Ledger = 19,
    Note = 20,
    Statement = 22,
    Tabloid = 23,
    Custom = 41
}

/// <summary>
/// How the text is lined up between the top and bottom margins of a page, as <c>WdVerticalAlignment</c> in Word (with the same values).
/// </summary>
public enum PageVerticalAlignment
{
    Top = 0,
    Center = 1,
    Justify = 2,
    Bottom = 3
}

/// <summary>
/// Where the gutter is, as <c>WdGutterStyle</c> in Word (with the same values).
/// </summary>
public enum GutterPosition
{
    Left = 0,
    Top = 1,
    Right = 2
}

/// <summary>
/// When line numbers start again, as <c>WdNumberingRule</c> in Word (with the same values).
/// </summary>
public enum LineNumberingRestart
{
    Continuous = 0,
    Section = 1,
    Page = 2
}

/// <summary>
/// How page numbers are written, as <c>WdPageNumberStyle</c> in Word (with the same values).
/// </summary>
public enum PageNumberStyle
{
    Arabic = 0,
    UppercaseRoman = 1,
    LowercaseRoman = 2,
    UppercaseLetter = 3,
    LowercaseLetter = 4
}

/// <summary>
/// Where a page number goes in a header or footer, as <c>WdPageNumberAlignment</c> in Word (with the same values).
/// </summary>
public enum PageNumberAlignment
{
    Left = 0,
    Center = 1,
    Right = 2,
    Inside = 3,
    Outside = 4
}

public partial class PageSetupWrapper
{
    //The sizes in twips, portrait.
    private static readonly (PaperSize Size, uint Width, uint Height)[] PaperSizes =
    {
        (PaperSize.Paper11x17, 15840, 24480), (PaperSize.Letter, 12240, 15840), (PaperSize.Legal, 12240, 20160), (PaperSize.Executive, 10440, 15120),
        (PaperSize.A3, 16838, 23811), (PaperSize.A4, 11906, 16838), (PaperSize.A5, 8391, 11906), (PaperSize.B4, 14570, 20636), (PaperSize.B5, 10319, 14570),
        (PaperSize.Folio, 12240, 18720), (PaperSize.Ledger, 24480, 15840), (PaperSize.Note, 12240, 15840), (PaperSize.Statement, 7920, 12240),
        (PaperSize.Tabloid, 15840, 24480)
    };

    /// <summary>
    /// The paper size, as <c>PageSetup.PaperSize</c> in Word: <see cref="PaperSize.Custom"/> when the page is not one of the known sizes.
    /// Setting a size sets the width and height of the page (swapped when the page is landscape).
    /// </summary>
    public PaperSize PaperSize
    {
        get
        {
            uint width = (uint)Math.Round(Math.Min(PageWidth, PageHeight) * 20);
            uint height = (uint)Math.Round(Math.Max(PageWidth, PageHeight) * 20);
            foreach ((PaperSize size, uint w, uint h) in PaperSizes)
            {
                //Ledger is the tabloid size turned on its side.
                uint shortSide = Math.Min(w, h);
                uint longSide = Math.Max(w, h);
                if (size != PaperSize.Ledger && size != PaperSize.Note && size != PaperSize.Paper11x17 && Math.Abs((int)shortSide - (int)width) <= 2 && Math.Abs((int)longSide - (int)height) <= 2)
                {
                    return size;
                }
            }
            return PaperSize.Custom;
        }
        set
        {
            if (value == PaperSize.Custom || !EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "Set the width and height for a custom paper size.");
            }

            (PaperSize Size, uint Width, uint Height) entry = PaperSizes.First(p => p.Size == value);
            bool landscape = Orientation == PageOrientation.Landscape;
            uint shortSide = Math.Min(entry.Width, entry.Height);
            uint longSide = Math.Max(entry.Width, entry.Height);
            PageSize pageSize = GetOrAdd<PageSize>();
            pageSize.Width = landscape ? longSide : shortSide;
            pageSize.Height = landscape ? shortSide : longSide;
        }
    }

    /// <summary>How the text is lined up between the top and bottom margins of the pages of the section (<c>PageSetup.VerticalAlignment</c>).</summary>
    public PageVerticalAlignment VerticalAlignment
    {
        get
        {
            VerticalJustificationValues? alignment = Properties?.GetFirstChild<VerticalTextAlignmentOnPage>()?.Val?.Value;
            if (alignment == VerticalJustificationValues.Center)
            {
                return PageVerticalAlignment.Center;
            }
            if (alignment == VerticalJustificationValues.Both)
            {
                return PageVerticalAlignment.Justify;
            }
            return alignment == VerticalJustificationValues.Bottom ? PageVerticalAlignment.Bottom : PageVerticalAlignment.Top;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            if (value == PageVerticalAlignment.Top)
            {
                Properties?.RemoveAllChildren<VerticalTextAlignmentOnPage>();
                return;
            }
            GetOrAdd<VerticalTextAlignmentOnPage>().Val = value == PageVerticalAlignment.Center ? VerticalJustificationValues.Center
                : value == PageVerticalAlignment.Justify ? VerticalJustificationValues.Both : VerticalJustificationValues.Bottom;
        }
    }

    /// <summary>True to leave the endnotes of the section out (they go at the end of the document).</summary>
    public bool SuppressEndnotes
    {
        get { return Properties?.GetFirstChild<NoEndnote>() != null; }
        set
        {
            if (value)
            {
                GetOrAdd<NoEndnote>();
            }
            else
            {
                Properties?.RemoveAllChildren<NoEndnote>();
            }
        }
    }

    private Settings GetSettings(bool create)
    {
        MainDocumentPart mainDocumentPart = _section.GetMainDocumentPart();
        if (mainDocumentPart == null)
        {
            if (create)
            {
                throw new InvalidOperationException("The section is not part of a document.");
            }
            return null;
        }

        DocumentSettingsPart settingsPart = mainDocumentPart.DocumentSettingsPart;
        if (settingsPart == null && create)
        {
            settingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>();
        }
        if (settingsPart != null && settingsPart.Settings == null && create)
        {
            settingsPart.Settings = new Settings();
        }
        return settingsPart?.Settings;
    }

    private static bool IsOn(OnOffType element)
    {
        return element != null && (element.Val == null || !element.Val.HasValue || element.Val.Value);
    }

    /// <summary>
    /// True to swap the inside and outside margins on facing pages, as <c>PageSetup.MirrorMargins</c> in Word. This is a setting of the whole
    /// document.
    /// </summary>
    public bool MirrorMargins
    {
        get { return IsOn(GetSettings(false)?.GetFirstChild<MirrorMargins>()); }
        set
        {
            if (!value)
            {
                GetSettings(false)?.RemoveAllChildren<MirrorMargins>();
                return;
            }
            Settings settings = GetSettings(true);
            if (settings.GetFirstChild<MirrorMargins>() == null)
            {
                settings.AddChild(new MirrorMargins());
            }
        }
    }

    /// <summary>
    /// Where the gutter is, as <c>PageSetup.GutterPos</c> in Word. The gutter at the top is a setting of the whole document; a gutter on
    /// the right is a property of the section.
    /// </summary>
    public GutterPosition GutterPos
    {
        get
        {
            if (IsOn(GetSettings(false)?.GetFirstChild<GutterAtTop>()))
            {
                return GutterPosition.Top;
            }
            return Properties?.GetFirstChild<GutterOnRight>() != null ? GutterPosition.Right : GutterPosition.Left;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            GetSettings(false)?.RemoveAllChildren<GutterAtTop>();
            Properties?.RemoveAllChildren<GutterOnRight>();
            if (value == GutterPosition.Top)
            {
                GetSettings(true).AddChild(new GutterAtTop());
            }
            else if (value == GutterPosition.Right)
            {
                GetOrAdd<GutterOnRight>();
            }
        }
    }

    /// <summary>The line numbering of the section, as <c>PageSetup.LineNumbering</c> in Word.</summary>
    public LineNumberingWrapper LineNumbering
    {
        get { return new LineNumberingWrapper(this); }
    }

    /// <summary>The columns of text of the section, as <c>PageSetup.TextColumns</c> in Word.</summary>
    public TextColumnsWrapper TextColumns
    {
        get { return new TextColumnsWrapper(this); }
    }

    /// <summary>
    /// The page borders of the section, as <c>PageSetup.Borders</c> in Word. A section has the four edges only (<c>Top</c>, <c>Left</c>,
    /// <c>Bottom</c> and <c>Right</c>); the distance of the border from the text is 24 points unless set.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return new BordersWrapper(_section.GetSectionProperties(true), false); }
    }

    internal T GetProperty<T>() where T : OpenXmlElement
    {
        return Properties?.GetFirstChild<T>();
    }

    internal T GetOrAddProperty<T>() where T : OpenXmlElement, new()
    {
        return GetOrAdd<T>();
    }

    internal void RemoveProperty<T>() where T : OpenXmlElement
    {
        Properties?.RemoveAllChildren<T>();
    }

    internal double GetTextWidth()
    {
        PageMargin margin = Properties?.GetFirstChild<PageMargin>();
        double left = (margin?.Left?.Value ?? 1440u) / 20.0;
        double right = (margin?.Right?.Value ?? 1440u) / 20.0;
        double gutter = (margin?.Gutter?.Value ?? 0u) / 20.0;
        return PageWidth - left - right - gutter;
    }
}

/// <summary>
/// The line numbers of a section, as <c>LineNumbering</c> in Word.
/// </summary>
public class LineNumberingWrapper
{
    private readonly PageSetupWrapper _setup;

    internal LineNumberingWrapper(PageSetupWrapper setup)
    {
        _setup = setup;
    }

    /// <summary>True if the lines are numbered. Setting it to true numbers every line, restarting on each page.</summary>
    public bool Active
    {
        get { return _setup.GetProperty<LineNumberType>() != null; }
        set
        {
            if (!value)
            {
                _setup.RemoveProperty<LineNumberType>();
            }
            else if (!Active)
            {
                _setup.GetOrAddProperty<LineNumberType>().Restart = LineNumberRestartValues.NewPage;
            }
        }
    }

    /// <summary>The number of lines between the numbers shown (every line when 1).</summary>
    public int CountBy
    {
        get { return _setup.GetProperty<LineNumberType>()?.CountBy?.Value ?? 1; }
        set
        {
            if (value < 1 || value > 100)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The count must be between 1 and 100.");
            }
            _setup.GetOrAddProperty<LineNumberType>().CountBy = (short)value;
        }
    }

    /// <summary>The number of the first line.</summary>
    public int StartingNumber
    {
        get { return (_setup.GetProperty<LineNumberType>()?.Start?.Value ?? 0) + 1; }
        set
        {
            if (value < 1 || value > 32767)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The number must be between 1 and 32767.");
            }
            _setup.GetOrAddProperty<LineNumberType>().Start = (short)(value - 1);
        }
    }

    /// <summary>The distance of the numbers from the text in points; 0 for the automatic distance.</summary>
    public double DistanceFromText
    {
        get
        {
            string distance = _setup.GetProperty<LineNumberType>()?.Distance?.Value;
            return distance != null && double.TryParse(distance, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips) ? twips / 20 : 0;
        }
        set
        {
            if (double.IsNaN(value) || value < 0 || value > 1584)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The distance must be between 0 and 1584 points.");
            }
            _setup.GetOrAddProperty<LineNumberType>().Distance = ((uint)Math.Round(value * 20, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
        }
    }

    /// <summary>When the numbers start again.</summary>
    public LineNumberingRestart RestartMode
    {
        get
        {
            LineNumberRestartValues? restart = _setup.GetProperty<LineNumberType>()?.Restart?.Value;
            if (restart == LineNumberRestartValues.NewPage)
            {
                return LineNumberingRestart.Page;
            }
            return restart == LineNumberRestartValues.NewSection ? LineNumberingRestart.Section : LineNumberingRestart.Continuous;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _setup.GetOrAddProperty<LineNumberType>().Restart = value == LineNumberingRestart.Page ? LineNumberRestartValues.NewPage
                : value == LineNumberingRestart.Section ? LineNumberRestartValues.NewSection : LineNumberRestartValues.Continuous;
        }
    }
}

/// <summary>
/// The columns of text of a section, as <c>TextColumns</c> in Word. Columns that are all the same width are described by their number and
/// the space between them.
/// </summary>
public class TextColumnsWrapper
{
    private readonly PageSetupWrapper _setup;

    internal TextColumnsWrapper(PageSetupWrapper setup)
    {
        _setup = setup;
    }

    private Columns Existing
    {
        get { return _setup.GetProperty<Columns>(); }
    }

    /// <summary>The number of columns.</summary>
    public int Count
    {
        get
        {
            Columns columns = Existing;
            if (columns == null)
            {
                return 1;
            }
            return columns.Elements<Column>().Any() ? columns.Elements<Column>().Count() : columns.ColumnCount?.Value ?? 1;
        }
    }

    /// <summary>True if the columns are the same width.</summary>
    public bool EvenlySpaced
    {
        get
        {
            Columns columns = Existing;
            return columns == null || !columns.Elements<Column>().Any();
        }
        set
        {
            Columns columns = _setup.GetOrAddProperty<Columns>();
            if (value)
            {
                int count = Count;
                columns.RemoveAllChildren<Column>();
                columns.EqualWidth = true;
                columns.ColumnCount = (short)count;
            }
            else if (!columns.Elements<Column>().Any())
            {
                //Columns that can differ are listed one by one.
                int count = Count;
                double each = (_setup.GetTextWidth() - Spacing * (count - 1)) / count;
                columns.EqualWidth = false;
                columns.ColumnCount = null;
                for (int i = 0; i < count; i++)
                {
                    Column column = new Column() { Width = ((uint)Math.Round(each * 20)).ToString(CultureInfo.InvariantCulture) };
                    if (i < count - 1)
                    {
                        column.Space = ((uint)Math.Round(Spacing * 20)).ToString(CultureInfo.InvariantCulture);
                    }
                    columns.AppendChild(column);
                }
            }
        }
    }

    /// <summary>Makes the text have the number of columns, all the same width.</summary>
    public void SetCount(int numColumns)
    {
        if (numColumns < 1 || numColumns > 45)
        {
            throw new ArgumentOutOfRangeException(nameof(numColumns), "The number of columns must be between 1 and 45.");
        }

        if (numColumns == 1 && Existing != null)
        {
            _setup.RemoveProperty<Columns>();
            return;
        }
        if (numColumns == 1)
        {
            return;
        }

        Columns columns = _setup.GetOrAddProperty<Columns>();
        bool hadSpace = columns.Space != null;
        columns.RemoveAllChildren<Column>();
        columns.EqualWidth = true;
        columns.ColumnCount = (short)numColumns;
        if (!hadSpace)
        {
            columns.Space = "720";
        }
    }

    /// <summary>True to draw a line between the columns.</summary>
    public bool LineBetween
    {
        get
        {
            OnOffValue separator = Existing?.Separator;
            return separator != null && separator.HasValue && separator.Value;
        }
        set
        {
            if (value)
            {
                _setup.GetOrAddProperty<Columns>().Separator = true;
            }
            else if (Existing != null)
            {
                Existing.Separator = null;
            }
        }
    }

    /// <summary>The space between evenly spaced columns in points (36 by default).</summary>
    public double Spacing
    {
        get
        {
            Columns columns = Existing;
            string space = columns?.Elements<Column>().FirstOrDefault()?.Space?.Value ?? columns?.Space?.Value;
            return space != null && double.TryParse(space, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips) ? twips / 20 : 36;
        }
        set
        {
            if (double.IsNaN(value) || value < 0 || value > 1584)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The spacing must be between 0 and 1584 points.");
            }

            Columns columns = _setup.GetOrAddProperty<Columns>();
            string twips = ((uint)Math.Round(value * 20, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture);
            columns.Space = twips;
            foreach (Column column in columns.Elements<Column>().Take(Math.Max(0, columns.Elements<Column>().Count() - 1)))
            {
                column.Space = twips;
            }
        }
    }

    /// <summary>The width of each column in points when they are evenly spaced.</summary>
    public double Width
    {
        get
        {
            Columns columns = Existing;
            Column first = columns?.Elements<Column>().FirstOrDefault();
            if (first?.Width != null && double.TryParse(first.Width.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double twips))
            {
                return twips / 20;
            }
            int count = Count;
            return (_setup.GetTextWidth() - Spacing * (count - 1)) / count;
        }
    }
}

/// <summary>
/// The page numbers of a header or footer, as <c>PageNumbers</c> in Word. The numbering options belong to the section.
/// </summary>
public class PageNumbersWrapper
{
    private readonly HeaderFooterWrapper _headerFooter;

    internal PageNumbersWrapper(HeaderFooterWrapper headerFooter)
    {
        _headerFooter = headerFooter;
    }

    private PageSetupWrapper Setup
    {
        get { return _headerFooter.Parent.PageSetup; }
    }

    private PageNumberType Existing
    {
        get { return Setup.GetProperty<PageNumberType>(); }
    }

    /// <summary>The number of page number fields at the start of a paragraph of the header or footer of their own.</summary>
    public int Count
    {
        get { return _headerFooter.Range.Fields.Count(f => f.Type == FieldType.Page); }
    }

    /// <summary>
    /// Adds a page number to the header or footer: a PAGE field in a paragraph of its own at the end, lined up as asked. Without
    /// <paramref name="firstPage"/> the first page has no number: the section gets a different first page header and footer.
    /// </summary>
    /// <exception cref="NotSupportedException">The alignment is Inside or Outside.</exception>
    public FieldWrapper Add(PageNumberAlignment alignment = PageNumberAlignment.Right, bool firstPage = true)
    {
        if (!EnumCompat.IsDefined(alignment))
        {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }
        if (alignment == PageNumberAlignment.Inside || alignment == PageNumberAlignment.Outside)
        {
            throw new NotSupportedException("Inside and outside page numbers are not supported.");
        }

        RangeWrapper range = _headerFooter.Range;
        ParagraphWrapper paragraph = range.Paragraphs.Count == 1 && range.Paragraphs[0].Text == "\r" ? range.Paragraphs[0] : range.InsertParagraphAfter();
        paragraph.Alignment = alignment == PageNumberAlignment.Left ? ParagraphAlignment.Left : alignment == PageNumberAlignment.Center ? ParagraphAlignment.Center : ParagraphAlignment.Right;
        FieldWrapper field = range.Fields.Add(paragraph.Range, FieldType.Page);

        if (!firstPage)
        {
            Setup.DifferentFirstPageHeaderFooter = true;
        }
        return field;
    }

    /// <summary>True if the numbering starts again at <see cref="StartingNumber"/> in the section.</summary>
    public bool RestartNumberingAtSection
    {
        get { return Existing?.Start != null && Existing.Start.HasValue; }
        set
        {
            if (value)
            {
                PageNumberType numbering = Setup.GetOrAddProperty<PageNumberType>();
                if (numbering.Start == null || !numbering.Start.HasValue)
                {
                    numbering.Start = 1;
                }
            }
            else if (Existing != null)
            {
                Existing.Start = null;
                if (!Existing.HasAttributes)
                {
                    Setup.RemoveProperty<PageNumberType>();
                }
            }
        }
    }

    /// <summary>The number of the first page of the section when the numbering restarts (1 by default).</summary>
    public int StartingNumber
    {
        get { return Existing?.Start?.Value ?? 1; }
        set
        {
            if (value < 0 || value > 32767)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The number must be between 0 and 32767.");
            }
            Setup.GetOrAddProperty<PageNumberType>().Start = value;
        }
    }

    /// <summary>How the numbers are written (Arabic, roman or letters).</summary>
    public PageNumberStyle NumberStyle
    {
        get
        {
            NumberFormatValues? format = Existing?.Format?.Value;
            if (format == NumberFormatValues.UpperRoman)
            {
                return PageNumberStyle.UppercaseRoman;
            }
            if (format == NumberFormatValues.LowerRoman)
            {
                return PageNumberStyle.LowercaseRoman;
            }
            if (format == NumberFormatValues.UpperLetter)
            {
                return PageNumberStyle.UppercaseLetter;
            }
            return format == NumberFormatValues.LowerLetter ? PageNumberStyle.LowercaseLetter : PageNumberStyle.Arabic;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            if (value == PageNumberStyle.Arabic)
            {
                if (Existing != null)
                {
                    Existing.Format = null;
                }
                return;
            }
            Setup.GetOrAddProperty<PageNumberType>().Format = value == PageNumberStyle.UppercaseRoman ? NumberFormatValues.UpperRoman
                : value == PageNumberStyle.LowercaseRoman ? NumberFormatValues.LowerRoman
                : value == PageNumberStyle.UppercaseLetter ? NumberFormatValues.UpperLetter : NumberFormatValues.LowerLetter;
        }
    }
}
