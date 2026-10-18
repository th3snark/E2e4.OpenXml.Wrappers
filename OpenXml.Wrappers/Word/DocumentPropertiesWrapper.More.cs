using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Wrappers.Word;

public partial class DocumentPropertiesWrapper
{
    private string GetExtended(Func<Properties, OpenXmlLeafTextElement> find)
    {
        OpenXmlLeafTextElement element = _document.ExtendedFilePropertiesPart?.Properties == null ? null : find(_document.ExtendedFilePropertiesPart.Properties);
        return element?.Text;
    }

    private void SetExtended<T>(Func<Properties, T> find, Action<Properties, T> set, Func<string, T> create, string value) where T : OpenXmlLeafTextElement
    {
        T existing = _document.ExtendedFilePropertiesPart?.Properties == null ? null : find(_document.ExtendedFilePropertiesPart.Properties);
        if (value == null)
        {
            existing?.Remove();
            return;
        }

        Properties properties = GetOrAddExtendedProperties();
        existing?.Remove();
        set(properties, create(value));
    }

    private static int? ToInt(string text)
    {
        return text != null && Int32.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number) ? number : (int?)null;
    }

    private static string FromInt(int? number)
    {
        return number?.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>The revision number (<c>Revision Number</c> in Word).</summary>
    public int? Revision
    {
        get { return ToInt(_document.PackageProperties.Revision); }
        set { _document.PackageProperties.Revision = FromInt(value); }
    }

    /// <summary>The language of the document, such as <c>en-US</c>.</summary>
    public string Language
    {
        get { return _document.PackageProperties.Language; }
        set { _document.PackageProperties.Language = value; }
    }

    /// <summary>The status of the content, such as <c>Draft</c>.</summary>
    public string ContentStatus
    {
        get { return _document.PackageProperties.ContentStatus; }
        set { _document.PackageProperties.ContentStatus = value; }
    }

    /// <summary>When the document was last printed (UTC).</summary>
    public DateTime? LastPrintDate
    {
        get { return _document.PackageProperties.LastPrinted; }
        set { _document.PackageProperties.LastPrinted = value; }
    }

    /// <summary>The template the document is based on, from the extended properties.</summary>
    public string Template
    {
        get { return GetExtended(p => p.Template); }
        set { SetExtended(p => p.Template, (p, e) => p.Template = e, v => new Template(v), value); }
    }

    /// <summary>The time spent editing the document in minutes, from the extended properties (<c>Total Editing Time</c> in Word).</summary>
    public int? TotalEditingTime
    {
        get { return ToInt(GetExtended(p => p.TotalTime)); }
        set { SetExtended(p => p.TotalTime, (p, e) => p.TotalTime = e, v => new TotalTime(v), FromInt(value)); }
    }

    /// <summary>The base address of the relative hyperlinks of the document (<c>Hyperlink base</c> in Word).</summary>
    public string HyperlinkBase
    {
        get { return GetExtended(p => p.HyperlinkBase); }
        set { SetExtended(p => p.HyperlinkBase, (p, e) => p.HyperlinkBase = e, v => new HyperlinkBase(v), value); }
    }

    /// <summary>The name of the application that saved the document, from the extended properties.</summary>
    public string ApplicationName
    {
        get { return GetExtended(p => p.Application); }
        set { SetExtended(p => p.Application, (p, e) => p.Application = e, v => new DocumentFormat.OpenXml.ExtendedProperties.Application(v), value); }
    }

    /// <summary>
    /// The number of pages that was stored when the document was saved (<c>Number of pages</c> in Word). It is not worked out by the library:
    /// see <see cref="WordprocessingDocumentWrapper.ComputeStatistics"/> for the counts that are.
    /// </summary>
    public int? Pages
    {
        get { return ToInt(GetExtended(p => p.Pages)); }
        set { SetExtended(p => p.Pages, (p, e) => p.Pages = e, v => new Pages(v), FromInt(value)); }
    }

    /// <summary>The number of words that was stored when the document was saved.</summary>
    public int? Words
    {
        get { return ToInt(GetExtended(p => p.Words)); }
        set { SetExtended(p => p.Words, (p, e) => p.Words = e, v => new Words(v), FromInt(value)); }
    }

    /// <summary>The number of characters that was stored when the document was saved.</summary>
    public int? Characters
    {
        get { return ToInt(GetExtended(p => p.Characters)); }
        set { SetExtended(p => p.Characters, (p, e) => p.Characters = e, v => new Characters(v), FromInt(value)); }
    }

    /// <summary>The number of characters with spaces that was stored when the document was saved.</summary>
    public int? CharactersWithSpaces
    {
        get { return ToInt(GetExtended(p => p.CharactersWithSpaces)); }
        set { SetExtended(p => p.CharactersWithSpaces, (p, e) => p.CharactersWithSpaces = e, v => new CharactersWithSpaces(v), FromInt(value)); }
    }

    /// <summary>The number of lines that was stored when the document was saved.</summary>
    public int? Lines
    {
        get { return ToInt(GetExtended(p => p.Lines)); }
        set { SetExtended(p => p.Lines, (p, e) => p.Lines = e, v => new Lines(v), FromInt(value)); }
    }

    /// <summary>The number of paragraphs that was stored when the document was saved.</summary>
    public int? Paragraphs
    {
        get { return ToInt(GetExtended(p => p.Paragraphs)); }
        set { SetExtended(p => p.Paragraphs, (p, e) => p.Paragraphs = e, v => new Paragraphs(v), FromInt(value)); }
    }

    /// <summary>
    /// A built-in property by the name that Word gives it (such as "Title", "Last Author", "Creation Date" or "Number of words"; case is ignored),
    /// as <c>BuiltInDocumentProperties("Title")</c> in Word. The value is a string, a number or a date (UTC) according to the property.
    /// </summary>
    /// <exception cref="ArgumentException">There is no built-in property with the name.</exception>
    public object this[string name]
    {
        get
        {
            switch (Normalise(name))
            {
                case "TITLE": return Title;
                case "SUBJECT": return Subject;
                case "AUTHOR": return Author;
                case "KEYWORDS": return Keywords;
                case "COMMENTS": return Comments;
                case "CATEGORY": return Category;
                case "LASTAUTHOR": return LastAuthor;
                case "CREATIONDATE": return CreationDate;
                case "LASTSAVETIME": return LastSaveTime;
                case "COMPANY": return Company;
                case "MANAGER": return Manager;
                case "TEMPLATE": return Template;
                case "REVISIONNUMBER": return Revision;
                case "LANGUAGE": return Language;
                case "CONTENTSTATUS": return ContentStatus;
                case "LASTPRINTDATE": return LastPrintDate;
                case "TOTALEDITINGTIME": return TotalEditingTime;
                case "HYPERLINKBASE": return HyperlinkBase;
                case "APPLICATIONNAME": return ApplicationName;
                case "NUMBEROFPAGES": return Pages;
                case "NUMBEROFWORDS": return Words;
                case "NUMBEROFCHARACTERS": return Characters;
                case "NUMBEROFCHARACTERS(WITHSPACES)": return CharactersWithSpaces;
                case "NUMBEROFLINES": return Lines;
                case "NUMBEROFPARAGRAPHS": return Paragraphs;
                default: throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "There is no built-in document property named '{0}'.", name), nameof(name));
            }
        }
        set
        {
            switch (Normalise(name))
            {
                case "TITLE": Title = (string)value; break;
                case "SUBJECT": Subject = (string)value; break;
                case "AUTHOR": Author = (string)value; break;
                case "KEYWORDS": Keywords = (string)value; break;
                case "COMMENTS": Comments = (string)value; break;
                case "CATEGORY": Category = (string)value; break;
                case "LASTAUTHOR": LastAuthor = (string)value; break;
                case "CREATIONDATE": CreationDate = (DateTime?)value; break;
                case "LASTSAVETIME": LastSaveTime = (DateTime?)value; break;
                case "COMPANY": Company = (string)value; break;
                case "MANAGER": Manager = (string)value; break;
                case "TEMPLATE": Template = (string)value; break;
                case "REVISIONNUMBER": Revision = (int?)value; break;
                case "LANGUAGE": Language = (string)value; break;
                case "CONTENTSTATUS": ContentStatus = (string)value; break;
                case "LASTPRINTDATE": LastPrintDate = (DateTime?)value; break;
                case "TOTALEDITINGTIME": TotalEditingTime = (int?)value; break;
                case "HYPERLINKBASE": HyperlinkBase = (string)value; break;
                case "APPLICATIONNAME": ApplicationName = (string)value; break;
                case "NUMBEROFPAGES": Pages = (int?)value; break;
                case "NUMBEROFWORDS": Words = (int?)value; break;
                case "NUMBEROFCHARACTERS": Characters = (int?)value; break;
                case "NUMBEROFCHARACTERS(WITHSPACES)": CharactersWithSpaces = (int?)value; break;
                case "NUMBEROFLINES": Lines = (int?)value; break;
                case "NUMBEROFPARAGRAPHS": Paragraphs = (int?)value; break;
                default: throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "There is no built-in document property named '{0}'.", name), nameof(name));
            }
        }
    }

    private static string Normalise(string name)
    {
        Guard.ThrowIfNull(name);
        return name.Replace(" ", String.Empty, StringComparison.Ordinal).ToUpperInvariant();
    }
}
