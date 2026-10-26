using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The custom properties of a document, as <c>Document.CustomDocumentProperties</c> in Word. Looking at the properties never changes the
/// document; the part that holds them is created by the first <see cref="Add"/>.
/// </summary>
public class CustomDocumentPropertyCollection : IEnumerable<CustomDocumentProperty>
{
    //The format id Word gives to every custom property.
    private const string FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    private readonly WordprocessingDocument _document;

    public CustomDocumentPropertyCollection(WordprocessingDocument document)
    {
        Guard.ThrowIfNull(document);
        _document = document;
    }

    private IList<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty> GetElementList()
    {
        return _document.CustomFilePropertiesPart?.Properties?.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>().ToList()
            ?? new List<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>();
    }

    public CustomDocumentProperty this[int index]
    {
        get
        {
            IList<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty> properties = GetElementList();
            if (index < 0 || index >= properties.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new CustomDocumentProperty(properties[index]);
        }
    }

    /// <summary>
    /// The property with the name (ignoring case), or null if there is none.
    /// </summary>
    public CustomDocumentProperty this[string name]
    {
        get
        {
            DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property = GetElementList()
                .FirstOrDefault(p => string.Equals(p.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
            return property == null ? null : new CustomDocumentProperty(property);
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    public bool Exists(string name)
    {
        return this[name] != null;
    }

    /// <summary>
    /// Adds a property, as <c>CustomDocumentProperties.Add</c> in Word (without the link to content), and returns it.
    /// </summary>
    /// <param name="name">The name, which must not be used by another property (ignoring case).</param>
    /// <param name="type">The type of the property.</param>
    /// <param name="value">The value, converted to the type. A null value is the empty string, zero or false, or the current time.</param>
    /// <exception cref="ArgumentException">The name is empty or taken, or the value cannot be converted to the type.</exception>
    public CustomDocumentProperty Add(string name, DocumentPropertyType type, object value)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A property needs a name.", nameof(name));
        }
        if (Exists(name))
        {
            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "A property named '{0}' already exists.", name), nameof(name));
        }
        object converted = Convert(type, value);

        CustomFilePropertiesPart part = _document.CustomFilePropertiesPart ?? _document.AddCustomFilePropertiesPart();
        if (part.Properties == null)
        {
            part.Properties = new Properties();
        }

        //Property ids start at 2 and are not reused while the property that has them exists.
        int id = part.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
            .Select(p => p.PropertyId?.Value ?? 1).DefaultIfEmpty(1).Max() + 1;
        DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property = new DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty()
        {
            FormatId = FormatId,
            PropertyId = id,
            Name = name
        };
        property.AppendChild(CustomDocumentProperty.CreateValue(converted));
        part.Properties.AppendChild(property);
        return new CustomDocumentProperty(property);
    }

    private static object Convert(DocumentPropertyType type, object value)
    {
        try
        {
            switch (type)
            {
                case DocumentPropertyType.Boolean:
                    return value != null && System.Convert.ToBoolean(value, CultureInfo.InvariantCulture);
                case DocumentPropertyType.Date:
                    return value == null ? DateTime.UtcNow : System.Convert.ToDateTime(value, CultureInfo.InvariantCulture);
                case DocumentPropertyType.FloatingPoint:
                    return value == null ? 0.0 : System.Convert.ToDouble(value, CultureInfo.InvariantCulture);
                case DocumentPropertyType.Number:
                    return value == null ? 0 : System.Convert.ToInt32(value, CultureInfo.InvariantCulture);
                case DocumentPropertyType.Text:
                    return value == null ? string.Empty : System.Convert.ToString(value, CultureInfo.InvariantCulture);
                default:
                    throw new ArgumentOutOfRangeException(nameof(type));
            }
        }
        catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException)
        {
            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "The value cannot be converted to the type {0}.", type), nameof(value), ex);
        }
    }

    #region IEnumerable Members

    public IEnumerator<CustomDocumentProperty> GetEnumerator()
    {
        foreach (DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property in GetElementList())
        {
            yield return new CustomDocumentProperty(property);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
