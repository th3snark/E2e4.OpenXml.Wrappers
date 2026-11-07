using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a node of a custom XML part is, as <c>MsoCustomXMLNodeType</c> in Office (with the same values).
/// </summary>
public enum CustomXMLNodeType
{
    /// <summary>An element.</summary>
    Element = 1,

    /// <summary>An attribute.</summary>
    Attribute = 2,

    /// <summary>Text.</summary>
    Text = 3,

    /// <summary>A CDATA section.</summary>
    CData = 4,

    /// <summary>A processing instruction.</summary>
    ProcessingInstruction = 5,

    /// <summary>A comment.</summary>
    Comment = 6,

    /// <summary>The document.</summary>
    Document = 7
}

/// <summary>
/// The prefixes a custom XML part's XPath expressions can use, as <c>CustomXMLPrefixMappings</c> in Word.
/// </summary>
public class CustomXMLPrefixMappings
{
    private readonly Dictionary<string, string> _prefixes = new Dictionary<string, string>(StringComparer.Ordinal);

    internal XmlNamespaceManager Manager()
    {
        XmlNamespaceManager manager = new XmlNamespaceManager(new NameTable());
        foreach (KeyValuePair<string, string> pair in _prefixes)
        {
            manager.AddNamespace(pair.Key, pair.Value);
        }
        return manager;
    }

    /// <summary>The number of prefixes.</summary>
    public int Count
    {
        get { return _prefixes.Count; }
    }

    /// <summary>Adds a prefix for a namespace, as <c>CustomXMLPrefixMappings.AddNamespace</c> in Word.</summary>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    public void AddNamespace(string prefix, string namespaceUri)
    {
        Guard.ThrowIfNull(prefix);
        Guard.ThrowIfNull(namespaceUri);
        _prefixes[prefix] = namespaceUri;
    }

    /// <summary>The namespace of a prefix, as <c>CustomXMLPrefixMappings.LookupNamespace</c> in Word; an empty string if there is none.</summary>
    public string LookupNamespace(string prefix)
    {
        Guard.ThrowIfNull(prefix);
        return _prefixes.TryGetValue(prefix, out string uri) ? uri : string.Empty;
    }

    /// <summary>A prefix of a namespace, as <c>CustomXMLPrefixMappings.LookupPrefix</c> in Word; an empty string if there is none.</summary>
    public string LookupPrefix(string namespaceUri)
    {
        Guard.ThrowIfNull(namespaceUri);
        return _prefixes.Where(p => p.Value == namespaceUri).Select(p => p.Key).FirstOrDefault() ?? string.Empty;
    }
}

/// <summary>
/// A node of a custom XML part, as <c>CustomXMLNode</c> in Word: an element, an attribute or a piece of text. The node is found again by
/// its position (an XPath expression such as <c>/ns0:root[1]/ns0:name[1]</c>) whenever it is used, so it stands for the place in the part.
/// </summary>
public class CustomXMLNode : IEquatable<CustomXMLNode>
{
    private readonly CustomXMLPart _part;
    private readonly string _xPath;
    private readonly string _prefixMappings;

    private CustomXMLNode(CustomXMLPart part, string xPath, string prefixMappings)
    {
        _part = part;
        _xPath = xPath;
        _prefixMappings = prefixMappings;
    }

    //The node for an object of a loaded document, named by its position and the prefixes in the order they appear down from the root.
    internal static CustomXMLNode From(CustomXMLPart part, XObject target)
    {
        List<XObject> chain = new List<XObject>();
        for (XObject each = target; each != null; each = each is XAttribute owner ? owner.Parent : (each as XNode)?.Parent)
        {
            chain.Insert(0, each);
        }

        List<string> namespaces = new List<string>();
        Func<XNamespace, string> prefixOf = ns =>
        {
            if (string.IsNullOrEmpty(ns.NamespaceName))
            {
                return string.Empty;
            }
            int at = namespaces.IndexOf(ns.NamespaceName);
            if (at < 0)
            {
                namespaces.Add(ns.NamespaceName);
                at = namespaces.Count - 1;
            }
            return "ns" + at.ToString(CultureInfo.InvariantCulture) + ":";
        };

        StringBuilder path = new StringBuilder();
        foreach (XObject step in chain)
        {
            if (step is XElement element)
            {
                string prefix = prefixOf(element.Name.Namespace);
                int index = element.Parent == null ? 1 : element.Parent.Elements(element.Name).TakeWhile(e => !ReferenceEquals(e, element)).Count() + 1;
                path.Append('/').Append(prefix).Append(element.Name.LocalName).Append('[').Append(index.ToString(CultureInfo.InvariantCulture)).Append(']');
            }
            else if (step is XAttribute attribute)
            {
                path.Append("/@").Append(prefixOf(attribute.Name.Namespace)).Append(attribute.Name.LocalName);
            }
            else if (step is XText text)
            {
                int index = text.Parent.Nodes().OfType<XText>().TakeWhile(t => !ReferenceEquals(t, text)).Count() + 1;
                path.Append("/text()[").Append(index.ToString(CultureInfo.InvariantCulture)).Append(']');
            }
            else if (step is XComment comment)
            {
                int index = comment.Parent.Nodes().OfType<XComment>().TakeWhile(c => !ReferenceEquals(c, comment)).Count() + 1;
                path.Append("/comment()[").Append(index.ToString(CultureInfo.InvariantCulture)).Append(']');
            }
        }

        StringBuilder mappings = new StringBuilder();
        for (int i = 0; i < namespaces.Count; i++)
        {
            mappings.Append("xmlns:ns").Append(i.ToString(CultureInfo.InvariantCulture)).Append("='").Append(namespaces[i]).Append("' ");
        }
        return new CustomXMLNode(part, path.ToString(), mappings.ToString());
    }

    internal static XmlNamespaceManager Namespaces(string prefixMappings)
    {
        XmlNamespaceManager manager = new XmlNamespaceManager(new NameTable());
        foreach (Match match in Regex.Matches(prefixMappings ?? string.Empty, @"xmlns:(?<p>[^=\s]+)=(['""])(?<u>.*?)\1"))
        {
            manager.AddNamespace(match.Groups["p"].Value, match.Groups["u"].Value);
        }
        return manager;
    }

    //The object in a loaded document, or null if the place is empty now.
    internal XObject Resolve(XDocument document)
    {
        return ((IEnumerable)document.XPathEvaluate(_xPath, Namespaces(_prefixMappings))).Cast<object>().OfType<XObject>().FirstOrDefault();
    }

    private XObject Target(XDocument document)
    {
        return Resolve(document) ?? throw new InvalidOperationException("The node is no longer in the part.");
    }

    /// <summary>The part the node is in, as <c>CustomXMLNode.OwnerPart</c> in Word.</summary>
    public CustomXMLPart OwnerPart
    {
        get { return _part; }
    }

    /// <summary>The position of the node as an XPath expression, as <c>CustomXMLNode.XPath</c> in Word.</summary>
    public string XPath
    {
        get { return _xPath; }
    }

    internal string PrefixMappings
    {
        get { return _prefixMappings; }
    }

    /// <summary>What the node is, as <c>CustomXMLNode.NodeType</c> in Word.</summary>
    public CustomXMLNodeType NodeType
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            return target is XElement ? CustomXMLNodeType.Element
                : target is XAttribute ? CustomXMLNodeType.Attribute
                : target is XCData ? CustomXMLNodeType.CData
                : target is XText ? CustomXMLNodeType.Text
                : target is XComment ? CustomXMLNodeType.Comment
                : target is XProcessingInstruction ? CustomXMLNodeType.ProcessingInstruction : CustomXMLNodeType.Document;
        }
    }

    /// <summary>The name of the node without a prefix, as <c>CustomXMLNode.BaseName</c> in Word; empty for text.</summary>
    public string BaseName
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            return target is XElement element ? element.Name.LocalName : target is XAttribute attribute ? attribute.Name.LocalName : string.Empty;
        }
    }

    /// <summary>The namespace of the node, as <c>CustomXMLNode.NamespaceURI</c> in Word; empty if it has none.</summary>
    public string NamespaceURI
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            return target is XElement element ? element.Name.NamespaceName : target is XAttribute attribute ? attribute.Name.NamespaceName : string.Empty;
        }
    }

    /// <summary>The text of the node, as <c>CustomXMLNode.Text</c> in Word. Setting it changes the part and the content controls bound to it.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public string Text
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            return target is XElement element ? element.Value : target is XAttribute attribute ? attribute.Value : target is XText text ? text.Value
                : target is XComment comment ? comment.Value : string.Empty;
        }
        set
        {
            Guard.ThrowIfNull(value);
            Change(target =>
            {
                if (target is XElement element)
                {
                    element.Value = value;
                }
                else if (target is XAttribute attribute)
                {
                    attribute.Value = value;
                }
                else if (target is XText text)
                {
                    text.Value = value;
                }
            });
        }
    }

    /// <summary>The XML of the node, as <c>CustomXMLNode.XML</c> in Word.</summary>
    public string XML
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            return target is XNode node ? node.ToString(System.Xml.Linq.SaveOptions.DisableFormatting) : target is XAttribute attribute ? attribute.ToString() : string.Empty;
        }
    }

    private void ChangeDocument(Action<XDocument> change)
    {
        XDocument document = _part.LoadDocument();
        change(document);
        _part.SaveDocument(document);
    }

    private void Change(Action<XObject> change)
    {
        XDocument document = _part.LoadDocument();
        change(Target(document));
        _part.SaveDocument(document);
    }

    private IEnumerable<XObject> Children(XDocument document)
    {
        XObject target = Target(document);
        return target is XElement element ? element.Nodes().Where(n => n is XElement || n is XText || n is XComment || n is XProcessingInstruction).Cast<XObject>() : Enumerable.Empty<XObject>();
    }

    private CustomXMLNodes Wrap(IEnumerable<XObject> objects)
    {
        return new CustomXMLNodes(objects.Select(o => From(_part, o)).ToList());
    }

    /// <summary>The nodes inside the node, as <c>CustomXMLNode.ChildNodes</c> in Word: its elements and its text.</summary>
    public CustomXMLNodes ChildNodes
    {
        get { return Wrap(Children(_part.LoadDocument())); }
    }

    /// <summary>
    /// The nodes an XPath expression matches, with the node as its context, as <c>CustomXMLNode.SelectNodes</c> in Word: <c>b</c> and
    /// <c>./b</c> are its children, <c>..</c> its parent and <c>/</c> the root of the part. The prefixes are those of
    /// <see cref="CustomXMLPart.NamespaceManager"/>; a name without a prefix has no namespace, so it does not match an element in one.
    /// </summary>
    /// <exception cref="ArgumentNullException">The expression is null.</exception>
    public CustomXMLNodes SelectNodes(string xPath)
    {
        Guard.ThrowIfNull(xPath);
        XObject target = Target(_part.LoadDocument());
        object result = ((XNode)(target is XAttribute attribute ? attribute.Parent : target)).XPathEvaluate(xPath, _part.NamespaceManager.Manager());
        IEnumerable<object> found = result is IEnumerable sequence && !(result is string) ? sequence.Cast<object>() : Enumerable.Empty<object>();
        return new CustomXMLNodes(found.OfType<XObject>().Select(o => From(_part, o)).ToList());
    }

    /// <summary>The first node an XPath expression matches, as <c>CustomXMLNode.SelectSingleNode</c> in Word; null if there is none. See <see cref="SelectNodes"/>.</summary>
    /// <exception cref="ArgumentNullException">The expression is null.</exception>
    public CustomXMLNode SelectSingleNode(string xPath)
    {
        return SelectNodes(xPath).FirstOrDefault();
    }

    /// <summary>The attributes of the node, as <c>CustomXMLNode.Attributes</c> in Word.</summary>
    public CustomXMLNodes Attributes
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            return Wrap(target is XElement element ? element.Attributes().Where(a => !a.IsNamespaceDeclaration).Cast<XObject>() : Enumerable.Empty<XObject>());
        }
    }

    /// <summary>True if the node has anything inside it, as <c>CustomXMLNode.HasChildNodes</c> in Word.</summary>
    public bool HasChildNodes()
    {
        return Children(_part.LoadDocument()).Any();
    }

    /// <summary>The first node inside the node, as <c>CustomXMLNode.FirstChild</c> in Word; null if there is none.</summary>
    public CustomXMLNode FirstChild
    {
        get { return ChildNodes.FirstOrDefault(); }
    }

    /// <summary>The last node inside the node, as <c>CustomXMLNode.LastChild</c> in Word; null if there is none.</summary>
    public CustomXMLNode LastChild
    {
        get { return ChildNodes.LastOrDefault(); }
    }

    /// <summary>The element the node is in, as <c>CustomXMLNode.ParentNode</c> in Word; null for the root element.</summary>
    public CustomXMLNode ParentNode
    {
        get
        {
            XObject target = Target(_part.LoadDocument());
            XElement parent = target is XAttribute attribute ? attribute.Parent : (target as XNode)?.Parent;
            return parent == null ? null : From(_part, parent);
        }
    }

    /// <summary>The node after the node in the same element, as <c>CustomXMLNode.NextSibling</c> in Word; null if there is none.</summary>
    public CustomXMLNode NextSibling
    {
        get { return Sibling(1); }
    }

    /// <summary>The node before the node in the same element, as <c>CustomXMLNode.PreviousSibling</c> in Word; null if there is none.</summary>
    public CustomXMLNode PreviousSibling
    {
        get { return Sibling(-1); }
    }

    private CustomXMLNode Sibling(int direction)
    {
        XObject target = Target(_part.LoadDocument());
        if (target is XAttribute || !(target is XNode node) || node.Parent == null)
        {
            return null;
        }
        List<XNode> all = node.Parent.Nodes().Where(n => n is XElement || n is XText || n is XComment || n is XProcessingInstruction).ToList();
        int at = all.FindIndex(n => ReferenceEquals(n, node)) + direction;
        return at < 0 || at >= all.Count ? null : From(_part, all[at]);
    }

    //Makes a node of the kind asked for: an element, an attribute, text, CDATA or a comment.
    private static XObject Make(string name, string namespaceUri, CustomXMLNodeType nodeType, string nodeValue)
    {
        XNamespace ns = namespaceUri ?? string.Empty;
        switch (nodeType)
        {
            case CustomXMLNodeType.Element:
                Guard.ThrowIfNull(name);
                return new XElement(ns + name, nodeValue ?? string.Empty);
            case CustomXMLNodeType.Attribute:
                Guard.ThrowIfNull(name);
                return new XAttribute(ns + name, nodeValue ?? string.Empty);
            case CustomXMLNodeType.Text:
                return new XText(nodeValue ?? string.Empty);
            case CustomXMLNodeType.CData:
                return new XCData(nodeValue ?? string.Empty);
            case CustomXMLNodeType.Comment:
                return new XComment(nodeValue ?? string.Empty);
            default:
                throw new ArgumentOutOfRangeException(nameof(nodeType));
        }
    }

    //Puts what was made inside an element: an attribute among its attributes, anything else among its nodes, before a node or at the end.
    private void Put(XDocument document, XObject made, XObject before)
    {
        if (!(Target(document) is XElement parent))
        {
            throw new InvalidOperationException("Only an element can have nodes inside it.");
        }
        if (made is XAttribute attribute)
        {
            parent.SetAttributeValue(attribute.Name, attribute.Value);
        }
        else if (before is XNode next && ReferenceEquals(next.Parent, parent))
        {
            next.AddBeforeSelf(made);
        }
        else
        {
            parent.Add(made);
        }
    }

    private static XObject InDocument(XDocument document, CustomXMLNode other)
    {
        return other == null ? null : other.Resolve(document) ?? throw new ArgumentException("The node is not in the part.", nameof(other));
    }

    /// <summary>Adds a node at the end of the node, as <c>CustomXMLNode.AppendChildNode</c> in Word: an element, an attribute, text, CDATA or a comment.</summary>
    /// <param name="name">The name; not used for text, CDATA and comments.</param>
    /// <param name="namespaceUri">The namespace of the name.</param>
    /// <param name="nodeType">What to add.</param>
    /// <param name="nodeValue">The text or value of the new node.</param>
    public void AppendChildNode(string name, string namespaceUri = "", CustomXMLNodeType nodeType = CustomXMLNodeType.Element, string nodeValue = "")
    {
        ChangeDocument(document => Put(document, Make(name, namespaceUri, nodeType, nodeValue), null));
    }

    /// <summary>Adds XML at the end of the node, as <c>CustomXMLNode.AppendChildSubtree</c> in Word.</summary>
    /// <exception cref="XmlException">The text is not well-formed XML.</exception>
    public void AppendChildSubtree(string xml)
    {
        Guard.ThrowIfNull(xml);
        XElement subtree = XElement.Parse(xml);
        ChangeDocument(document => Put(document, subtree, null));
    }

    /// <summary>Adds a node before another node of this one, as <c>CustomXMLNode.InsertNodeBefore</c> in Word; at the end when it is null.</summary>
    public void InsertNodeBefore(string name, string namespaceUri = "", CustomXMLNodeType nodeType = CustomXMLNodeType.Element, string nodeValue = "", CustomXMLNode nextSibling = null)
    {
        ChangeDocument(document => Put(document, Make(name, namespaceUri, nodeType, nodeValue), InDocument(document, nextSibling)));
    }

    /// <summary>Adds XML before another node of this one, as <c>CustomXMLNode.InsertSubtreeBefore</c> in Word; at the end when it is null.</summary>
    /// <exception cref="XmlException">The text is not well-formed XML.</exception>
    public void InsertSubtreeBefore(string xml, CustomXMLNode nextSibling = null)
    {
        Guard.ThrowIfNull(xml);
        XElement subtree = XElement.Parse(xml);
        ChangeDocument(document => Put(document, subtree, InDocument(document, nextSibling)));
    }

    /// <summary>Replaces a node of this one with a new node, as <c>CustomXMLNode.ReplaceChildNode</c> in Word.</summary>
    /// <exception cref="ArgumentNullException">The old node is null.</exception>
    public void ReplaceChildNode(CustomXMLNode oldNode, string name, string namespaceUri = "", CustomXMLNodeType nodeType = CustomXMLNodeType.Element, string nodeValue = "")
    {
        Guard.ThrowIfNull(oldNode);
        ChangeDocument(document =>
        {
            XObject old = InDocument(document, oldNode);
            Put(document, Make(name, namespaceUri, nodeType, nodeValue), old is XNode node ? node : null);
            RemoveObject(old);
        });
    }

    /// <summary>Replaces a node of this one with XML, as <c>CustomXMLNode.ReplaceChildSubtree</c> in Word.</summary>
    /// <exception cref="ArgumentNullException">The old node is null.</exception>
    /// <exception cref="XmlException">The text is not well-formed XML.</exception>
    public void ReplaceChildSubtree(string xml, CustomXMLNode oldNode)
    {
        Guard.ThrowIfNull(xml);
        Guard.ThrowIfNull(oldNode);
        XElement subtree = XElement.Parse(xml);
        ChangeDocument(document =>
        {
            XObject old = InDocument(document, oldNode);
            Put(document, subtree, old is XNode node ? node : null);
            RemoveObject(old);
        });
    }

    private static void RemoveObject(XObject target)
    {
        if (target is XAttribute attribute)
        {
            attribute.Remove();
        }
        else if (target is XNode node)
        {
            node.Remove();
        }
    }
    /// <summary>Takes the node out of the part, as <c>CustomXMLNode.Delete</c> in Word. The content controls bound to it keep their contents.</summary>
    public void Delete()
    {
        Change(target =>
        {
            if (target is XAttribute attribute)
            {
                attribute.Remove();
            }
            else if (target is XNode node)
            {
                node.Remove();
            }
        });
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as CustomXMLNode);
    }

    /// <inheritdoc/>
    public bool Equals(CustomXMLNode other)
    {
        return !(other is null) && _part == other._part && _xPath == other._xPath && _prefixMappings == other._prefixMappings;
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _xPath.GetHashCode();
    }

    public static bool operator ==(CustomXMLNode left, CustomXMLNode right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(CustomXMLNode left, CustomXMLNode right)
    {
        return !(left == right);
    }
}

/// <summary>
/// A list of nodes of a custom XML part, as <c>CustomXMLNodes</c> in Word.
/// </summary>
public class CustomXMLNodes : IEnumerable<CustomXMLNode>
{
    private readonly List<CustomXMLNode> _nodes;

    internal CustomXMLNodes(List<CustomXMLNode> nodes)
    {
        _nodes = nodes;
    }

    /// <summary>The number of nodes.</summary>
    public int Count
    {
        get { return _nodes.Count; }
    }

    /// <summary>The node at the index (0-based), as <c>CustomXMLNodes.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no node at the index.</exception>
    public CustomXMLNode this[int index]
    {
        get
        {
            if (index < 0 || index >= _nodes.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return _nodes[index];
        }
    }

    /// <inheritdoc/>
    public IEnumerator<CustomXMLNode> GetEnumerator()
    {
        return _nodes.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
