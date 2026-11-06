using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A custom XML part of the document, as <c>CustomXMLPart</c> in Word: a piece of XML stored in the package that content controls can be
/// bound to.
/// </summary>
public class CustomXMLPart : IEquatable<CustomXMLPart>
{
    private readonly MainDocumentPart _main;
    private readonly CustomXmlPart _part;

    internal CustomXMLPart(MainDocumentPart main, CustomXmlPart part)
    {
        _main = main;
        _part = part;
    }

    internal CustomXmlPart Part
    {
        get { return _part; }
    }

    private XDocument Load()
    {
        using (Stream stream = _part.GetStream(FileMode.Open, FileAccess.Read))
        {
            return XDocument.Load(stream);
        }
    }

    internal XDocument LoadDocument()
    {
        return Load();
    }

    //Saves a changed document and shows the new values in the content controls bound to the part.
    internal void SaveDocument(XDocument document)
    {
        Save(document, _part);
        string id = Id;
        foreach (SdtElement sdt in _main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<SdtElement>()).ToList())
        {
            DataBinding binding = sdt.SdtProperties?.GetFirstChild<DataBinding>();
            if (binding?.StoreItemId?.Value != id || binding.XPath?.Value == null)
            {
                continue;
            }
            object found = ((IEnumerable)document.XPathEvaluate(binding.XPath.Value, CustomXMLNode.Namespaces(binding.PrefixMappings?.Value))).Cast<object>().FirstOrDefault();
            string value = found is XElement element ? element.Value : found is XAttribute attribute ? attribute.Value : null;
            if (value != null)
            {
                new ContentControlWrapper(sdt).ReplaceContents(value);
            }
        }
    }

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<CustomXmlPart, CustomXMLPrefixMappings> Prefixes =
        new System.Runtime.CompilerServices.ConditionalWeakTable<CustomXmlPart, CustomXMLPrefixMappings>();

    /// <summary>The prefixes the XPath expressions given to <see cref="SelectNodes"/> and <see cref="SelectSingleNode"/> can use, as <c>CustomXMLPart.NamespaceManager</c> in Word.</summary>
    public CustomXMLPrefixMappings NamespaceManager
    {
        get { return Prefixes.GetOrCreateValue(_part); }
    }

    /// <summary>The root element of the part, as <c>CustomXMLPart.DocumentElement</c> in Word.</summary>
    public CustomXMLNode DocumentElement
    {
        get { return CustomXMLNode.From(this, Load().Root); }
    }

    /// <summary>The first node an XPath expression matches, as <c>CustomXMLPart.SelectSingleNode</c> in Word; null if there is none.</summary>
    /// <exception cref="ArgumentNullException">The expression is null.</exception>
    public CustomXMLNode SelectSingleNode(string xPath)
    {
        return SelectNodes(xPath).FirstOrDefault();
    }

    /// <summary>The nodes an XPath expression matches, as <c>CustomXMLPart.SelectNodes</c> in Word. The prefixes are those of <see cref="NamespaceManager"/>.</summary>
    /// <exception cref="ArgumentNullException">The expression is null.</exception>
    public CustomXMLNodes SelectNodes(string xPath)
    {
        Guard.ThrowIfNull(xPath);
        XDocument document = Load();
        object result = document.XPathEvaluate(xPath, NamespaceManager.Manager());
        IEnumerable<object> found = result is IEnumerable sequence && !(result is string) ? sequence.Cast<object>() : Enumerable.Empty<object>();
        return new CustomXMLNodes(found.OfType<XObject>().Select(o => CustomXMLNode.From(this, o)).ToList());
    }

    /// <summary>The id of the part, as <c>CustomXMLPart.Id</c> in Word: a GUID in braces.</summary>
    public string Id
    {
        get { return _part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value ?? string.Empty; }
    }

    /// <summary>The namespace of the root element of the part, as <c>CustomXMLPart.NamespaceURI</c> in Word.</summary>
    public string NamespaceURI
    {
        get { return Load().Root?.Name.NamespaceName ?? string.Empty; }
    }

    /// <summary>The XML of the part, as <c>CustomXMLPart.XML</c> in Word.</summary>
    public string XML
    {
        get { return Load().ToString(System.Xml.Linq.SaveOptions.DisableFormatting); }
    }

    /// <summary>True for the parts Word keeps itself (such as the bibliography sources), as <c>CustomXMLPart.BuiltIn</c> in Word.</summary>
    public bool BuiltIn
    {
        get { return Load().Root?.Name.NamespaceName == "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"; }
    }

    /// <summary>Replaces the XML of the part, as <c>CustomXMLPart.LoadXML</c> in Word.</summary>
    /// <exception cref="XmlException">The text is not well-formed XML.</exception>
    public bool LoadXML(string xml)
    {
        Guard.ThrowIfNull(xml);
        Save(XDocument.Parse(xml), _part);
        return true;
    }

    /// <summary>Replaces the XML of the part with that of a file, as <c>CustomXMLPart.Load</c> in Word.</summary>
    public bool Load(string path)
    {
        Guard.ThrowIfNull(path);
        Save(XDocument.Load(path), _part);
        return true;
    }

    internal static void Save(XDocument document, CustomXmlPart part)
    {
        using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write))
        {
            document.Save(stream);
        }
        CustomXmlPropertiesPart properties = part.CustomXmlPropertiesPart;
        if (properties?.DataStoreItem != null)
        {
            string ns = document.Root?.Name.NamespaceName;
            SchemaReferences references = new SchemaReferences();
            if (!string.IsNullOrEmpty(ns))
            {
                references.AppendChild(new SchemaReference() { Uri = ns });
            }
            properties.DataStoreItem.SchemaReferences = references;
        }
    }

    /// <summary>
    /// Takes the part out of the document, as <c>CustomXMLPart.Delete</c> in Word. Content controls bound to it keep their contents but
    /// are no longer bound.
    /// </summary>
    public void Delete()
    {
        string id = Id;
        foreach (SdtElement sdt in _main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<SdtElement>()).ToList())
        {
            if (sdt.SdtProperties?.GetFirstChild<DataBinding>()?.StoreItemId?.Value == id)
            {
                new ContentControlWrapper(sdt).XMLMapping.Delete();
            }
        }
        _main.DeletePart(_part);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as CustomXMLPart);
    }

    /// <inheritdoc/>
    public bool Equals(CustomXMLPart other)
    {
        return !(other is null) && ReferenceEquals(_part, other._part);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _part.GetHashCode();
    }

    public static bool operator ==(CustomXMLPart left, CustomXMLPart right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(CustomXMLPart left, CustomXMLPart right)
    {
        return !(left == right);
    }
}

/// <summary>
/// The custom XML parts of a document, as <c>CustomXMLParts</c> in Word.
/// </summary>
public class CustomXMLParts : IEnumerable<CustomXMLPart>
{
    private readonly MainDocumentPart _main;

    internal CustomXMLParts(MainDocumentPart main)
    {
        _main = main;
    }

    private List<CustomXMLPart> Items()
    {
        return _main.CustomXmlParts.Select(p => new CustomXMLPart(_main, p)).ToList();
    }

    /// <summary>The number of parts.</summary>
    public int Count
    {
        get { return _main.CustomXmlParts.Count(); }
    }

    /// <summary>The part at the index (0-based), as <c>CustomXMLParts.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no part at the index.</exception>
    public CustomXMLPart this[int index]
    {
        get
        {
            List<CustomXMLPart> items = Items();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>Adds a part with the XML, as <c>CustomXMLParts.Add</c> in Word.</summary>
    /// <exception cref="XmlException">The text is not well-formed XML.</exception>
    public CustomXMLPart Add(string xml)
    {
        Guard.ThrowIfNull(xml);
        XDocument document = XDocument.Parse(xml);
        CustomXmlPart part = _main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        CustomXmlPropertiesPart properties = part.AddNewPart<CustomXmlPropertiesPart>();
        properties.DataStoreItem = new DataStoreItem()
        {
            ItemId = "{" + Guid.NewGuid().ToString("D").ToUpperInvariant() + "}"
        };
        CustomXMLPart.Save(document, part);
        return new CustomXMLPart(_main, part);
    }

    /// <summary>The part with the id, as <c>CustomXMLParts.SelectByID</c> in Word; null if there is none.</summary>
    public CustomXMLPart SelectByID(string id)
    {
        Guard.ThrowIfNull(id);
        return Items().FirstOrDefault(p => string.Equals(p.Id, id, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>The parts whose root element is in the namespace, as <c>CustomXMLParts.SelectByNamespace</c> in Word.</summary>
    public IReadOnlyList<CustomXMLPart> SelectByNamespace(string namespaceUri)
    {
        Guard.ThrowIfNull(namespaceUri);
        return Items().Where(p => p.NamespaceURI == namespaceUri).ToList();
    }

    /// <inheritdoc/>
    public IEnumerator<CustomXMLPart> GetEnumerator()
    {
        return Items().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// The binding of a content control to an element of a custom XML part, as <c>XMLMapping</c> in Word.
/// </summary>
public class XMLMapping
{
    private readonly ContentControlWrapper _control;

    internal XMLMapping(ContentControlWrapper control)
    {
        _control = control;
    }

    private DataBinding Binding
    {
        get { return _control.Element.SdtProperties?.GetFirstChild<DataBinding>(); }
    }

    /// <summary>True if the control is bound, as <c>XMLMapping.IsMapped</c> in Word.</summary>
    public bool IsMapped
    {
        get { return Binding != null; }
    }

    /// <summary>The XPath expression of the element the control is bound to, as <c>XMLMapping.XPath</c> in Word.</summary>
    public string XPath
    {
        get { return Binding?.XPath?.Value ?? string.Empty; }
    }

    /// <summary>The namespace prefixes the XPath expression uses, as <c>XMLMapping.PrefixMappings</c> in Word.</summary>
    public string PrefixMappings
    {
        get { return Binding?.PrefixMappings?.Value ?? string.Empty; }
    }

    /// <summary>The custom XML part the control is bound to, as <c>XMLMapping.CustomXMLPart</c> in Word; null when it is not bound.</summary>
    public CustomXMLPart CustomXMLPart
    {
        get
        {
            string id = Binding?.StoreItemId?.Value;
            MainDocumentPart main = StyleHelper.GetMainDocumentPart(_control.Element);
            return id == null || main == null ? null : new CustomXMLParts(main).SelectByID(id);
        }
    }

    private static XmlNamespaceManager Namespaces(string prefixMappings)
    {
        XmlNamespaceManager manager = new XmlNamespaceManager(new NameTable());
        foreach (Match match in Regex.Matches(prefixMappings ?? string.Empty, @"xmlns:(?<p>[^=\s]+)=(['""])(?<u>.*?)\1"))
        {
            manager.AddNamespace(match.Groups["p"].Value, match.Groups["u"].Value);
        }
        return manager;
    }

    /// <summary>
    /// Binds the control to an element of a custom XML part, as <c>XMLMapping.SetMapping</c> in Word, and shows the text of the element in
    /// it. Returns false, and leaves the control as it is, when the XPath expression matches no element.
    /// </summary>
    /// <param name="xPath">The XPath expression of the element.</param>
    /// <param name="prefixMappings">The prefixes the expression uses, as <c>xmlns:ns='urn:x'</c> separated by spaces.</param>
    /// <param name="source">The part; the first one the expression matches in when null.</param>
    public bool SetMapping(string xPath, string prefixMappings = null, CustomXMLPart source = null)
    {
        Guard.ThrowIfNull(xPath);
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(_control.Element);
        if (main == null)
        {
            return false;
        }
        XmlNamespaceManager namespaces = Namespaces(prefixMappings);
        CustomXMLPart found = null;
        string text = null;
        foreach (CustomXMLPart candidate in source != null ? new[] { source } : new CustomXMLParts(main).ToArray())
        {
            XElement element = XDocument.Parse(candidate.XML).XPathSelectElement(xPath, namespaces);
            if (element != null)
            {
                found = candidate;
                text = element.Value;
                break;
            }
        }
        if (found == null)
        {
            return false;
        }

        _control.SetBinding(new DataBinding()
        {
            PrefixMappings = string.IsNullOrEmpty(prefixMappings) ? null : prefixMappings,
            XPath = xPath,
            StoreItemId = found.Id
        });
        _control.ReplaceContents(text);
        return true;
    }

    /// <summary>Binds the control to a node, as <c>XMLMapping.SetMappingByNode</c> in Word, and shows the text of the node in it.</summary>
    /// <exception cref="ArgumentNullException">The node is null.</exception>
    public bool SetMappingByNode(CustomXMLNode node)
    {
        Guard.ThrowIfNull(node);
        return SetMapping(node.XPath, node.PrefixMappings, node.OwnerPart);
    }

    /// <summary>The node the control is bound to, as <c>XMLMapping.CustomXMLNode</c> in Word; null when it is not bound or the node is gone.</summary>
    public CustomXMLNode CustomXMLNode
    {
        get
        {
            CustomXMLPart part = CustomXMLPart;
            if (part == null || Binding?.XPath?.Value == null)
            {
                return null;
            }
            XDocument document = XDocument.Parse(part.XML);
            XObject target = ((IEnumerable)document.XPathEvaluate(Binding.XPath.Value, Namespaces(PrefixMappings))).OfType<XObject>().FirstOrDefault();
            return target == null ? null : Word.CustomXMLNode.From(part, target);
        }
    }

    /// <summary>Removes the binding, as <c>XMLMapping.Delete</c> in Word. The control keeps its contents.</summary>
    public void Delete()
    {
        _control.SetBinding(null);
    }
}
