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
public class DocumentPropertiesWrapper
{
    private readonly WordprocessingDocument _document;

    public DocumentPropertiesWrapper(WordprocessingDocument document)
    {
        Guard.ThrowIfNull(document);
        _document = document;
    }

    public string Title
    {
        get { return _document.PackageProperties.Title; }
        set { _document.PackageProperties.Title = value; }
    }

    public string Subject
    {
        get { return _document.PackageProperties.Subject; }
        set { _document.PackageProperties.Subject = value; }
    }

    /// <summary>
    /// The author of the document (the creator in the core properties).
    /// </summary>
    public string Author
    {
        get { return _document.PackageProperties.Creator; }
        set { _document.PackageProperties.Creator = value; }
    }

    public string Keywords
    {
        get { return _document.PackageProperties.Keywords; }
        set { _document.PackageProperties.Keywords = value; }
    }

    /// <summary>
    /// The comments about the document (the description in the core properties).
    /// </summary>
    public string Comments
    {
        get { return _document.PackageProperties.Description; }
        set { _document.PackageProperties.Description = value; }
    }

    public string Category
    {
        get { return _document.PackageProperties.Category; }
        set { _document.PackageProperties.Category = value; }
    }

    public string LastAuthor
    {
        get { return _document.PackageProperties.LastModifiedBy; }
        set { _document.PackageProperties.LastModifiedBy = value; }
    }

    /// <summary>
    /// When the document was created (UTC).
    /// </summary>
    public DateTime? CreationDate
    {
        get { return _document.PackageProperties.Created; }
        set { _document.PackageProperties.Created = value; }
    }

    /// <summary>
    /// When the document was last saved (UTC).
    /// </summary>
    public DateTime? LastSaveTime
    {
        get { return _document.PackageProperties.Modified; }
        set { _document.PackageProperties.Modified = value; }
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
