using System;

using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The built-in properties of a document, as <c>Document.BuiltInDocumentProperties</c> in Word.
/// </summary>
/// <remarks>
/// A property that has not been set reads as null. Setting a property requires a document that was opened as editable.
/// </remarks>
public partial class DocumentPropertiesWrapper
{
    private readonly WordprocessingDocument _document;

    public DocumentPropertiesWrapper(WordprocessingDocument document)
    {
        Guard.ThrowIfNull(document);
        _document = document;
    }

#pragma warning disable OOXML0001 // Only the name of the experimental interface is used; the properties themselves are not marked.
    private IPackageProperties CoreProperties
    {
        get { return GetCoreProperties(_document); }
    }

    /// <summary>
    /// The core properties of the package (title, author, dates and so on). An editable document that has none gets an empty
    /// <c>/docProps/core.xml</c> first, where Word keeps them: the packaging layer would otherwise put them in a part of its own naming
    /// (<c>/package/services/metadata/core-properties/*.psmdcp</c>). It picks the part on the first use, so every use goes through here.
    /// </summary>
    internal static IPackageProperties GetCoreProperties(WordprocessingDocument document)
    {
        if (document.CoreFilePropertiesPart == null && document.FileOpenAccess != System.IO.FileAccess.Read)
        {
            CoreFilePropertiesPart part = document.AddCoreFilePropertiesPart();
            using (System.IO.Stream stream = part.GetStream(System.IO.FileMode.Create))
            using (System.IO.StreamWriter writer = new System.IO.StreamWriter(stream, new System.Text.UTF8Encoding(false)))
            {
                writer.Write(EmptyCoreProperties);
            }
        }
        return document.PackageProperties;
    }
#pragma warning restore OOXML0001

    private const string EmptyCoreProperties = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        + "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\""
        + " xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\""
        + " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>";

    public string Title
    {
        get { return CoreProperties.Title; }
        set { CoreProperties.Title = value; }
    }

    public string Subject
    {
        get { return CoreProperties.Subject; }
        set { CoreProperties.Subject = value; }
    }

    /// <summary>
    /// The author of the document (the creator in the core properties).
    /// </summary>
    public string Author
    {
        get { return CoreProperties.Creator; }
        set { CoreProperties.Creator = value; }
    }

    public string Keywords
    {
        get { return CoreProperties.Keywords; }
        set { CoreProperties.Keywords = value; }
    }

    /// <summary>
    /// The comments about the document (the description in the core properties).
    /// </summary>
    public string Comments
    {
        get { return CoreProperties.Description; }
        set { CoreProperties.Description = value; }
    }

    public string Category
    {
        get { return CoreProperties.Category; }
        set { CoreProperties.Category = value; }
    }

    public string LastAuthor
    {
        get { return CoreProperties.LastModifiedBy; }
        set { CoreProperties.LastModifiedBy = value; }
    }

    /// <summary>
    /// When the document was created (UTC).
    /// </summary>
    public DateTime? CreationDate
    {
        get { return CoreProperties.Created; }
        set { CoreProperties.Created = value; }
    }

    /// <summary>
    /// When the document was last saved (UTC).
    /// </summary>
    public DateTime? LastSaveTime
    {
        get { return CoreProperties.Modified; }
        set { CoreProperties.Modified = value; }
    }

    /// <summary>
    /// The company, from the extended properties.
    /// </summary>
    public string Company
    {
        get { return _document.ExtendedFilePropertiesPart?.Properties?.Company?.Text; }
        set
        {
            if (value == null && _document.ExtendedFilePropertiesPart?.Properties?.Company == null)
            {
                return;
            }
            Properties properties = GetOrAddExtendedProperties();
            if (value == null)
            {
                properties.Company.Remove();
            }
            else
            {
                properties.Company = new Company(value);
            }
        }
    }

    /// <summary>
    /// The manager, from the extended properties.
    /// </summary>
    public string Manager
    {
        get { return _document.ExtendedFilePropertiesPart?.Properties?.Manager?.Text; }
        set
        {
            if (value == null && _document.ExtendedFilePropertiesPart?.Properties?.Manager == null)
            {
                return;
            }
            Properties properties = GetOrAddExtendedProperties();
            if (value == null)
            {
                properties.Manager.Remove();
            }
            else
            {
                properties.Manager = new Manager(value);
            }
        }
    }

    private Properties GetOrAddExtendedProperties()
    {
        ExtendedFilePropertiesPart part = _document.ExtendedFilePropertiesPart ?? _document.AddExtendedFilePropertiesPart();
        if (part.Properties == null)
        {
            part.Properties = new Properties();
        }
        return part.Properties;
    }
}
