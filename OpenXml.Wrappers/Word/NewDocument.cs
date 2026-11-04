using System;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using W14DocumentId = DocumentFormat.OpenXml.Office2010.Word.DocumentId;
using W15DocumentId = DocumentFormat.OpenXml.Office2013.Word.PersistentDocumentId;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What Word writes for a new blank document. The parts are the ones Word 365 saved for a blank document with the current Office theme, kept
/// as embedded resources in <c>Templates</c> (written by <c>tools/WordReference</c>, without revision and document ids); the values that Word
/// takes from Windows' regional settings (page size, header distance, language, decimal symbol and list separator) are set from the current
/// culture.
/// </summary>
internal static class NewDocument
{
    private const string TemplatePrefix = "OpenXml.Wrappers.Word.Templates.";

    /// <summary>Opens one of the embedded templates: document.xml, styles.xml, settings.xml, BuiltInStyles.xml and so on.</summary>
    public static Stream OpenTemplate(string fileName)
    {
        Stream stream = typeof(NewDocument).Assembly.GetManifestResourceStream(TemplatePrefix + fileName);
        if (stream == null)
        {
            throw new InvalidOperationException("The template " + fileName + " is not in the assembly.");
        }
        return stream;
    }

    /// <summary>Makes the main document a new one: one empty paragraph and the section properties of the page.</summary>
    public static void SetDocument(MainDocumentPart mainDocumentPart)
    {
        Document document = new Document(ReadRootXml("document.xml"));
        mainDocumentPart.Document = document;
        RegionalDefaults defaults = RegionalDefaults.Current;
        SectionProperties section = document.Body.Elements<SectionProperties>().Last();
        PageSize size = section.GetFirstChild<PageSize>();
        size.Width = defaults.PageWidth;
        size.Height = defaults.PageHeight;
        PageMargin margin = section.GetFirstChild<PageMargin>();
        margin.Top = RegionalDefaults.Margin;
        margin.Bottom = RegionalDefaults.Margin;
        margin.Left = (uint)RegionalDefaults.Margin;
        margin.Right = (uint)RegionalDefaults.Margin;
        margin.Header = defaults.HeaderDistance;
        margin.Footer = defaults.HeaderDistance;
        section.GetFirstChild<Columns>().Space = defaults.HeaderDistance.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Adds the parts a new document has that the main document part does not have yet: settings, styles, the theme, the font table and web
    /// settings.
    /// </summary>
    public static void AddMissingParts(MainDocumentPart mainDocumentPart)
    {
        RegionalDefaults defaults = RegionalDefaults.Current;
        if (mainDocumentPart.DocumentSettingsPart == null)
        {
            DocumentSettingsPart part = mainDocumentPart.AddNewPart<DocumentSettingsPart>();
            Feed(part, "settings.xml");
            SetSettings(part.Settings, defaults);
        }
        if (mainDocumentPart.StyleDefinitionsPart == null)
        {
            StyleDefinitionsPart part = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            Feed(part, "styles.xml");
            Languages language = part.Styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<Languages>();
            if (language != null)
            {
                language.Val = defaults.Language;
            }
        }
        if (mainDocumentPart.ThemePart == null)
        {
            Feed(mainDocumentPart.AddNewPart<ThemePart>(), "theme1.xml");
        }
        if (mainDocumentPart.FontTablePart == null)
        {
            Feed(mainDocumentPart.AddNewPart<FontTablePart>(), "fontTable.xml");
        }
        if (mainDocumentPart.WebSettingsPart == null)
        {
            Feed(mainDocumentPart.AddNewPart<WebSettingsPart>(), "webSettings.xml");
        }
    }

    /// <summary>
    /// The document properties Word writes for a new document, if the package has none: the core properties (revision 1, created and
    /// modified now; no author, which Word takes from its user name) and the application properties (the template; no statistics, which the
    /// library does not keep up to date).
    /// </summary>
    public static void AddMissingPackageParts(WordprocessingDocument document)
    {
#pragma warning disable OOXML0001
        DocumentFormat.OpenXml.Packaging.IPackageProperties properties = DocumentPropertiesWrapper.GetCoreProperties(document);
#pragma warning restore OOXML0001
        if (properties.Created == null && properties.Revision == null)
        {
            //Word stores the times to the minute.
            DateTime now = DateTime.UtcNow;
            DateTime minute = new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0, DateTimeKind.Utc);
            properties.Revision = "1";
            properties.Created = minute;
            properties.Modified = minute;
        }
        if (document.ExtendedFilePropertiesPart == null)
        {
            Feed(document.AddExtendedFilePropertiesPart(), "app.xml");
        }
    }

    //The root element of a template as XML, without the XML declaration, which an element cannot be built from.
    private static string ReadRootXml(string template)
    {
        using (Stream stream = OpenTemplate(template))
        using (StreamReader reader = new StreamReader(stream))
        {
            string xml = reader.ReadToEnd();
            return xml.StartsWith("<?xml", StringComparison.Ordinal) ? xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2) : xml;
        }
    }

    private static void Feed(OpenXmlPart part, string template)
    {
        using (Stream stream = OpenTemplate(template))
        {
            part.FeedData(stream);
        }
    }

    private static void SetSettings(Settings settings, RegionalDefaults defaults)
    {
        ThemeFontLanguages themeFontLanguages = settings.GetFirstChild<ThemeFontLanguages>();
        if (themeFontLanguages != null)
        {
            themeFontLanguages.Val = defaults.Language;
        }
        DecimalSymbol decimalSymbol = settings.GetFirstChild<DecimalSymbol>();
        if (decimalSymbol != null)
        {
            decimalSymbol.Val = defaults.DecimalSymbol;
        }
        ListSeparator listSeparator = settings.GetFirstChild<ListSeparator>();
        if (listSeparator != null)
        {
            listSeparator.Val = defaults.ListSeparator;
        }

        //Word gives every document two ids of its own: w14:docId (eight hex digits, below 80000000) and w15:docId (a GUID).
        byte[] bytes = Guid.NewGuid().ToByteArray();
        uint id = BitConverter.ToUInt32(bytes, 0) & 0x7FFFFFFF;
        W14DocumentId w14 = new W14DocumentId() { Val = id.ToString("X8", CultureInfo.InvariantCulture) };
        W15DocumentId w15 = new W15DocumentId() { Val = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}" };
        OpenXmlElement chartTracking = settings.GetFirstChild<DocumentFormat.OpenXml.Office2013.Word.ChartTrackingRefBased>();
        if (chartTracking != null)
        {
            chartTracking.InsertBeforeSelf(w14);
            chartTracking.InsertAfterSelf(w15);
        }
        else
        {
            settings.AppendChild(w14);
            settings.AppendChild(w15);
        }
    }
}

/// <summary>
/// The defaults Word takes from Windows' regional settings for a new document, worked out from a culture: the paper (Letter in the regions
/// whose Windows default is Letter, otherwise A4), the header and footer distance and column spacing (half an inch, or 1.25 cm where the
/// measurement system is metric), the editing language, the decimal symbol and the list separator.
/// </summary>
internal sealed class RegionalDefaults
{
    /// <summary>The page margins of a new document in Word 365: one inch (2.54 cm) all round, whatever the region.</summary>
    public const int Margin = 1440;

    //The regions whose Windows default paper size is Letter.
    private static readonly string[] LetterRegions = { "US", "CA", "MX", "PR", "PH", "CL", "CO", "CR", "DO", "GT", "NI", "PA", "SV", "VE" };

    private RegionalDefaults(CultureInfo culture)
    {
        CultureInfo specific = ToSpecific(culture);
        RegionInfo region = new RegionInfo(specific.Name);
        Language = specific.Name;
        bool letter = LetterRegions.Contains(region.TwoLetterISORegionName);
        PageWidth = letter ? 12240U : 11906U;
        PageHeight = letter ? 15840U : 16838U;
        HeaderDistance = region.IsMetric ? 708U : 720U;
        DecimalSymbol = specific.NumberFormat.NumberDecimalSeparator;
        ListSeparator = specific.TextInfo.ListSeparator;
    }

    public static RegionalDefaults Current
    {
        get { return new RegionalDefaults(CultureInfo.CurrentCulture); }
    }

    public string Language { get; }

    public uint PageWidth { get; }

    public uint PageHeight { get; }

    /// <summary>The header and footer distance and the column spacing, in twips.</summary>
    public uint HeaderDistance { get; }

    public string DecimalSymbol { get; }

    public string ListSeparator { get; }

    //The invariant culture and neutral cultures (such as "en") have no region: take the culture Windows would (en-US for the invariant one).
    private static CultureInfo ToSpecific(CultureInfo culture)
    {
        if (culture.Name.Length == 0)
        {
            return CultureInfo.GetCultureInfo("en-US");
        }
        if (!culture.IsNeutralCulture)
        {
            return culture;
        }
        try
        {
            CultureInfo specific = CultureInfo.CreateSpecificCulture(culture.Name);
            return specific.Name.Length == 0 ? CultureInfo.GetCultureInfo("en-US") : specific;
        }
        catch (ArgumentException)
        {
            return CultureInfo.GetCultureInfo("en-US");
        }
    }
}
