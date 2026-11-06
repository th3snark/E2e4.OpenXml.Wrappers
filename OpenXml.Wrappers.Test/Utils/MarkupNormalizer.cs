using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace OpenXml.Wrappers.Test.Utils;

/// <summary>
/// Puts the XML parts of a .docx package into a canonical form, so that a document Word wrote can be compared with one the library wrote.
/// </summary>
/// <remarks>
/// <para>
/// Values that change on every save are taken out: revision ids (<c>w:rsid*</c>), <c>w14:paraId</c> and <c>w14:textId</c>, proofing marks and state,
/// <c>w:lastRenderedPageBreak</c> and Word's <c>_GoBack</c> bookmark. Ids (relationship, bookmark, comment, note, comment paragraph and
/// durable ids, <c>docPr</c> ids) become ordinal placeholders such as <c>{bookmark1}</c>, numbered in order of first use, so that what refers
/// to what is kept. Dates, document ids, password hashes and salts become fixed placeholders. In <c>comments.xml</c> the <c>w14:paraId</c>
/// of each paragraph becomes a placeholder rather than being removed, because <c>commentsExtended.xml</c> refers to it.
/// </para>
/// <para>
/// Attributes and namespace declarations are sorted, namespace declarations that nothing uses are dropped, <c>mc:Ignorable</c> lists only
/// the prefixes still used (sorted), relationships are sorted, relationship targets are made absolute and
/// each part is listed with its content type, whether a <c>Default</c> or an <c>Override</c> gives it. The result is
/// indented XML with LF line ends and no XML declaration.
/// </para>
/// <para>
/// One implementation serves both sides: <c>tools/WordReference/New-WordReferences.ps1</c> runs it through <c>tools/WordReference/Tool</c>
/// (which links this file) to write the Word snapshots, and <c>MarkupConformanceTest</c> runs it on the library's output.
/// </para>
/// </remarks>
public static class MarkupNormalizer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace W16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private static readonly XNamespace W16cex = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
    private static readonly XNamespace PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";

    private const string ContentTypesEntry = "[Content_Types].xml";
    private const string PackageRelationshipsEntry = "_rels/.rels";
    private const string RelationshipsFolder = "_rels/";

    //The text of a DATE (or similar) field result, which is today's date in the short format of the machine.
    private static readonly Regex DateText = new Regex(@"^(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{4}-\d{2}-\d{2})$", RegexOptions.CultureInvariant);

    /// <summary>
    /// The name a package entry is written under, or null when the entry is not compared: <c>[Content_Types].xml</c> becomes
    /// <c>content-types.xml</c>, <c>_rels/.rels</c> becomes <c>package.rels</c>, and an XML part under <c>word/</c> keeps its path below
    /// <c>word/</c> with <c>/</c> turned into <c>.</c> (<c>document.xml</c>, <c>theme.theme1.xml</c>, <c>document.xml.rels</c>). Document
    /// properties and binary parts are left out; custom XML parts are named <c>customXml.item1.xml</c>.
    /// </summary>
    public static string GetPartFileName(string entryName)
    {
        if (entryName == ContentTypesEntry)
        {
            return "content-types.xml";
        }
        if (entryName == PackageRelationshipsEntry)
        {
            return "package.rels";
        }
        if (!(entryName.EndsWith(".xml", StringComparison.Ordinal) || entryName.EndsWith(".rels", StringComparison.Ordinal)))
        {
            return null;
        }
        if (entryName.StartsWith("customXml/", StringComparison.Ordinal))
        {
            return CanonicalPartName(entryName).Replace(RelationshipsFolder, string.Empty).Replace('/', '.');
        }
        if (!entryName.StartsWith("word/", StringComparison.Ordinal))
        {
            return null;
        }
        return entryName.Substring("word/".Length).Replace(RelationshipsFolder, string.Empty).Replace('/', '.');
    }

    /// <summary>Normalises every compared part of the package in the file. See <see cref="NormalizePackage(Stream)"/>.</summary>
    public static SortedDictionary<string, string> NormalizePackage(string path)
    {
        using (FileStream stream = File.OpenRead(path))
        {
            return NormalizePackage(stream);
        }
    }

    /// <summary>
    /// Normalises every compared part of the package, keyed by <see cref="GetPartFileName"/>. The stream is left open.
    /// </summary>
    public static SortedDictionary<string, string> NormalizePackage(Stream package)
    {
        Dictionary<string, XDocument> parts = new Dictionary<string, XDocument>(StringComparer.Ordinal);
        List<string> entryNames = new List<string>();
        using (ZipArchive archive = new ZipArchive(package, ZipArchiveMode.Read, true))
        {
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                entryNames.Add(entry.FullName);
                if (GetPartFileName(entry.FullName) != null)
                {
                    using (Stream stream = entry.Open())
                    {
                        parts.Add(entry.FullName, XDocument.Load(stream, LoadOptions.PreserveWhitespace));
                    }
                }
            }
        }

        //Placeholders are numbered in order of first use, so the parts are visited in a fixed order: the main document first, then the other
        //parts by name (which puts comments.xml before the parts that refer to its paragraphs), then the relationship parts, whose ids the
        //parts they belong to have mostly numbered already.
        Placeholders placeholders = new Placeholders();
        List<string> order = parts.Keys
            .OrderBy(IsRelationshipsEntry)
            .ThenBy(e => e != "word/document.xml")
            .ThenBy(e => GetPartFileName(e), StringComparer.Ordinal)
            .ToList();
        SortedDictionary<string, string> result = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (string entryName in order)
        {
            XElement root = parts[entryName].Root;
            if (entryName == ContentTypesEntry)
            {
                NormalizeContentTypes(root, entryNames);
            }
            else if (IsRelationshipsEntry(entryName))
            {
                NormalizeRelationships(root, GetRelationshipsSource(entryName), placeholders);
            }
            else
            {
                NormalizePart(root, entryName, placeholders);
            }
            NormalizeNamespaces(root);
            SortAttributes(root);

            string fileName = GetPartFileName(entryName);
            if (result.ContainsKey(fileName))
            {
                throw new InvalidDataException(string.Format(CultureInfo.InvariantCulture, "Two parts of the package are both written as '{0}'.", fileName));
            }
            result.Add(fileName, Write(new[] { root }));
        }
        return result;
    }

    /// <summary>
    /// Reads normalised XML back and writes only the elements <paramref name="select"/> picks from it, in the same canonical form, so that a
    /// part of a part can be compared.
    /// </summary>
    public static string Select(string normalizedXml, Func<XElement, IEnumerable<XElement>> select)
    {
        XElement root = XElement.Parse(normalizedXml, LoadOptions.PreserveWhitespace);
        RemoveFormattingWhitespace(root);
        return Write(select(root));
    }

    private static bool IsRelationshipsEntry(string entryName)
    {
        return entryName.EndsWith(".rels", StringComparison.Ordinal);
    }

    //The entry a relationships entry belongs to: word/_rels/document.xml.rels belongs to word/document.xml and _rels/.rels to the package ("").
    private static string GetRelationshipsSource(string entryName)
    {
        int index = entryName.LastIndexOf(RelationshipsFolder, StringComparison.Ordinal);
        string withoutFolder = entryName.Remove(index, RelationshipsFolder.Length);
        return withoutFolder.Substring(0, withoutFolder.Length - ".rels".Length);
    }

    private static void NormalizePart(XElement root, string entryName, Placeholders placeholders)
    {
        Dictionary<string, string> relationshipIds = placeholders.ForRelationshipsOf(entryName);
        bool commentParagraphIds = entryName == "word/comments.xml";

        HashSet<string> goBackIds = new HashSet<string>(root.Descendants(W + "bookmarkStart")
            .Where(b => (string)b.Attribute(W + "name") == "_GoBack")
            .Select(b => (string)b.Attribute(W + "id")), StringComparer.Ordinal);
        root.Descendants().Where(e => IsRemoved(e, goBackIds)).ToList().Remove();

        foreach (XElement element in root.DescendantsAndSelf())
        {
            foreach (XAttribute attribute in element.Attributes().ToList())
            {
                NormalizeAttribute(element, attribute, commentParagraphIds, relationshipIds, placeholders);
            }
            if (element.Name == W + "t" && DateText.IsMatch(element.Value))
            {
                element.Value = "{date}";
            }
            if (element.Name == W + "instrText")
            {
                //The names of the bookmarks of a table of contents are numbers Word makes up.
                element.Value = TocName.Replace(element.Value, m => "_Toc" + placeholders.Get("toc", m.Groups[1].Value));
            }
        }
        RemoveFormattingWhitespace(root);
        MergeRuns(root);
    }

    //Word keeps runs apart that differ only in their revision ids, which are removed above, so runs of the same formatting that follow one another
    //are put together here, by the same rules as OpenXmlExtensions.CoalesceRuns. A test on the library's XML checks that it merges them itself.
    private static void MergeRuns(XElement root)
    {
        XName run = W + "r";
        XName text = W + "t";
        XName tab = W + "tab";
        XName breakName = W + "br";
        XName properties = W + "rPr";
        XNamespace xml = XNamespace.Xml;
        foreach (XElement parent in root.Descendants(run).Select(r => r.Parent).Distinct().ToList())
        {
            XElement previous = null;
            foreach (XElement current in parent.Elements().ToList())
            {
                if (current.Name != run)
                {
                    previous = null;
                    continue;
                }
                if (previous != null)
                {
                    List<XElement> first = previous.Elements().Where(e => e.Name != properties).ToList();
                    List<XElement> second = current.Elements().Where(e => e.Name != properties).ToList();
                    bool simple = first.Count > 0 && second.Count > 0 && second[0].Name == text
                        && first.Concat(second).All(e => e.Name == text || e.Name == tab || e.Name == breakName);
                    XElement lastOfFirst = simple ? first[first.Count - 1] : null;
                    bool afterBlockBreak = simple && lastOfFirst.Name == breakName && lastOfFirst.Attribute(W + "type") != null;
                    if (simple && !afterBlockBreak && XNode.DeepEquals(previous.Element(properties), current.Element(properties)))
                    {
                        if (lastOfFirst.Name == text)
                        {
                            string value = lastOfFirst.Value + second[0].Value;
                            lastOfFirst.Value = value;
                            lastOfFirst.SetAttributeValue(xml + "space", value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1])) ? "preserve" : null);
                            second.RemoveAt(0);
                        }
                        foreach (XElement moved in second)
                        {
                            moved.Remove();
                            previous.Add(moved);
                        }
                        current.Remove();
                        continue;
                    }
                }
                previous = current;
            }
        }
    }

    private static bool IsRemoved(XElement element, HashSet<string> goBackIds)
    {
        XName name = element.Name;
        //w:proofState records how far Word's background spelling and grammar check got before the save.
        if (name == W + "proofErr" || name == W + "proofState" || name == W + "lastRenderedPageBreak")
        {
            return true;
        }
        //w:rsids, w:rsidRoot and w:rsid in the settings.
        if (name.Namespace == W && name.LocalName.StartsWith("rsid", StringComparison.Ordinal))
        {
            return true;
        }
        return (name == W + "bookmarkStart" || name == W + "bookmarkEnd") && goBackIds.Contains((string)element.Attribute(W + "id"));
    }

    private static readonly System.Text.RegularExpressions.Regex TocName = new System.Text.RegularExpressions.Regex(@"_Toc(\d+)");

    private static void NormalizeAttribute(XElement element, XAttribute attribute, bool commentParagraphIds, Dictionary<string, string> relationshipIds,
        Placeholders placeholders)
    {
        if (attribute.IsNamespaceDeclaration)
        {
            return;
        }
        XName name = attribute.Name;
        if (name == W + "dirty")
        {
            //Word does not mark a field to be updated when it saves; the library does for the fields that need layout.
            attribute.Remove();
        }
        else if ((name == W + "name" || name == W + "anchor") && TocName.IsMatch(attribute.Value))
        {
            attribute.Value = TocName.Replace(attribute.Value, m => "_Toc" + placeholders.Get("toc", m.Groups[1].Value));
        }
        else if ((name.Namespace == W && name.LocalName.StartsWith("rsid", StringComparison.Ordinal)) || name == W14 + "textId" || name == WP14 + "anchorId" || name == WP14 + "editId")
        {
            attribute.Remove();
        }
        else if (name == W14 + "paraId")
        {
            if (commentParagraphIds)
            {
                attribute.Value = placeholders.Get("para", attribute.Value);
            }
            else
            {
                attribute.Remove();
            }
        }
        else if (name.Namespace == R)
        {
            attribute.Value = Placeholders.Get(relationshipIds, "rId", attribute.Value);
        }
        else if (name == W + "id")
        {
            attribute.Value = placeholders.Get(GetIdKind(element.Name), attribute.Value);
        }
        else if (element.Name == W + "id" && name == W + "val")
        {
            //The id of a content control (w:sdtPr/w:id).
            attribute.Value = placeholders.Get("sdt", attribute.Value);
        }
        else if (element.Name == WP + "docPr" && name == "id")
        {
            attribute.Value = placeholders.Get("docPr", attribute.Value);
        }
        else if (element.Name == Pic + "cNvPr" && name == "id")
        {
            //Word gives a picture a random id of its own.
            attribute.Value = placeholders.Get("pic", attribute.Value);
        }
        else if (name == W15 + "paraId" || name == W15 + "paraIdParent" || name == W16cid + "paraId")
        {
            attribute.Value = placeholders.Get("para", attribute.Value);
        }
        else if (element.Name == W + "guid" && name == W + "val")
        {
            attribute.Value = placeholders.Get("guid", attribute.Value);
        }
        else if (name.LocalName == "itemID" || name.LocalName == "storeItemID")
        {
            //The id of a custom XML part, in its properties and in the content controls bound to it.
            attribute.Value = placeholders.Get("store", attribute.Value.ToUpperInvariant());
        }
        else if (name == W + "val" && (element.Name == W + "docPart" || (element.Name == W + "name" && element.Parent?.Name == W + "docPartPr"))
            && System.Text.RegularExpressions.Regex.IsMatch(attribute.Value, "^[0-9A-F]{32}$"))
        {
            //The name Word gives a placeholder the user wrote, in the glossary document and in the control that uses it.
            attribute.Value = placeholders.Get("docpart", attribute.Value);
        }
        else if (element.Name == W + "tmpl" && name == W + "val")
        {
            attribute.Value = placeholders.Get("tmpl", attribute.Value);
        }
        else if (element.Name == W + "nsid" && name == W + "val")
        {
            attribute.Value = placeholders.Get("nsid", attribute.Value);
        }
        else if (name.LocalName == "durableId" && (name.Namespace == W16cid || name.Namespace == W16cex))
        {
            attribute.Value = placeholders.Get("durable", attribute.Value);
        }
        else if (name == W + "date" || name.LocalName == "dateUtc")
        {
            attribute.Value = "{date}";
        }
        else if (element.Name.LocalName == "docId" && name.LocalName == "val")
        {
            attribute.Value = "{docId}";
        }
        else if (name == W + "hash" || name == W + "hashValue")
        {
            attribute.Value = "{hash}";
        }
        else if (name == W + "salt" || name == W + "saltValue")
        {
            attribute.Value = "{salt}";
        }
    }

    private static string GetIdKind(XName elementName)
    {
        switch (elementName.LocalName)
        {
            case "bookmarkStart":
            case "bookmarkEnd":
                return "bookmark";
            case "comment":
            case "commentRangeStart":
            case "commentRangeEnd":
            case "commentReference":
                return "comment";
            case "footnote":
            case "footnoteReference":
            case "endnote":
            case "endnoteReference":
                return "note";
            default:
                return "id";
        }
    }

    //The Open XML SDK names the first custom XML part item.xml where Word names it item1.xml. The name means nothing, so it is compared as Word's.
    private static string CanonicalPartName(string name)
    {
        //Likewise the SDK puts pictures in /media, numbered from nothing, where Word puts them in /word/media, numbered from 1.
        return System.Text.RegularExpressions.Regex.Replace(
            name.Replace("customXml/item.xml", "customXml/item1.xml").Replace("customXml/_rels/item.xml.rels", "customXml/_rels/item1.xml.rels"),
            @"^(/?)media/image(\d*)\.",
            m => m.Groups[1].Value + "word/media/image" + (m.Groups[2].Value.Length == 0 ? 1 : int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture) + 1).ToString(CultureInfo.InvariantCulture) + ".");
    }

    private static void NormalizeRelationships(XElement root, string source, Placeholders placeholders)
    {
        Dictionary<string, string> relationshipIds = placeholders.ForRelationshipsOf(source);
        string folder = source.Substring(0, source.LastIndexOf('/') + 1);
        List<XElement> relationships = root.Elements(PackageRelationships + "Relationship").ToList();
        foreach (XElement relationship in relationships)
        {
            XAttribute target = relationship.Attribute("Target");
            if (target != null && (string)relationship.Attribute("TargetMode") == "External" && target.Value.StartsWith("file:", StringComparison.OrdinalIgnoreCase))
            {
                //A linked file is somewhere on the machine that wrote the document.
                target.Value = "file:///{path}";
            }
            if (target != null && (string)relationship.Attribute("TargetMode") != "External")
            {
                target.Value = CanonicalPartName(ResolveTarget(folder, target.Value));
            }
        }
        relationships = relationships
            .OrderBy(r => (string)r.Attribute("Type"), StringComparer.Ordinal)
            .ThenBy(r => (string)r.Attribute("Target"), StringComparer.Ordinal)
            .ToList();
        foreach (XElement relationship in relationships)
        {
            XAttribute id = relationship.Attribute("Id");
            id.Value = Placeholders.Get(relationshipIds, "rId", id.Value);
        }
        root.ReplaceNodes(relationships);
    }

    //Makes a relationship target an absolute part name: "styles.xml" from word/document.xml becomes "/word/styles.xml".
    private static string ResolveTarget(string folder, string target)
    {
        if (target.Length > 0 && target[0] == '/')
        {
            return target;
        }
        List<string> segments = folder.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries).ToList();
        foreach (string segment in target.Split('/'))
        {
            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.RemoveAt(segments.Count - 1);
                }
            }
            else if (segment.Length > 0 && segment != ".")
            {
                segments.Add(segment);
            }
        }
        return "/" + string.Join("/", segments);
    }

    //The content type of each part, whether it comes from a Default (by extension) or an Override: which of the two a package uses depends on
    //the order its writer added the parts in, and means the same. Relationship parts keep the Default that gives them their type.
    private static void NormalizeContentTypes(XElement root, List<string> entryNames)
    {
        Dictionary<string, string> defaults = root.Elements(ContentTypes + "Default")
            .ToDictionary(e => ((string)e.Attribute("Extension")).ToUpperInvariant(), e => (string)e.Attribute("ContentType"), StringComparer.Ordinal);
        Dictionary<string, string> overrides = root.Elements(ContentTypes + "Override")
            .ToDictionary(e => (string)e.Attribute("PartName"), e => (string)e.Attribute("ContentType"), StringComparer.OrdinalIgnoreCase);

        List<XElement> elements = new List<XElement>();
        if (defaults.TryGetValue("RELS", out string relationships))
        {
            elements.Add(new XElement(ContentTypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", relationships)));
        }
        foreach (string entryName in entryNames.Where(n => n != ContentTypesEntry && n[n.Length - 1] != '/' && !IsRelationshipsEntry(n))
            .OrderBy(n => CanonicalPartName(n), StringComparer.OrdinalIgnoreCase))
        {
            string partName = "/" + entryName;
            string extension = Path.GetExtension(entryName).TrimStart('.').ToUpperInvariant();
            if (overrides.TryGetValue(partName, out string contentType) || defaults.TryGetValue(extension, out contentType))
            {
                elements.Add(new XElement(ContentTypes + "Override", new XAttribute("PartName", CanonicalPartName(partName)), new XAttribute("ContentType", contentType)));
            }
        }
        root.ReplaceNodes(elements);
    }

    //Removes the whitespace between elements (indentation), keeping the text of elements that hold only text, such as w:t.
    private static void RemoveFormattingWhitespace(XElement root)
    {
        root.DescendantNodes().OfType<XText>().Where(t => t.Parent != null && t.Parent.HasElements && string.IsNullOrWhiteSpace(t.Value)).ToList().Remove();
    }

    private static void NormalizeNamespaces(XElement root)
    {
        HashSet<XNamespace> used = GetUsedNamespaces(root);
        foreach (XElement element in root.DescendantsAndSelf())
        {
            XAttribute ignorable = element.Attribute(MC + "Ignorable");
            if (ignorable != null)
            {
                string[] prefixes = ignorable.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(p => element.GetNamespaceOfPrefix(p) is XNamespace ns && used.Contains(ns))
                    .Distinct()
                    .OrderBy(p => p, StringComparer.Ordinal)
                    .ToArray();
                if (prefixes.Length == 0)
                {
                    ignorable.Remove();
                }
                else
                {
                    ignorable.Value = string.Join(" ", prefixes);
                }
            }
        }

        used = GetUsedNamespaces(root);
        foreach (XElement element in root.DescendantsAndSelf())
        {
            element.Attributes().Where(a => a.IsNamespaceDeclaration && !used.Contains(XNamespace.Get(a.Value))).ToList().Remove();
        }

        //Where a namespace is declared does not matter, so prefixed declarations on inner elements move to the root unless that would clash
        //with a declaration already there.
        foreach (XAttribute declaration in root.Descendants().Attributes().Where(a => a.IsNamespaceDeclaration && a.Name.Namespace == XNamespace.Xmlns).ToList())
        {
            XAttribute atRoot = root.Attribute(declaration.Name);
            string prefixAtRoot = root.GetPrefixOfNamespace(XNamespace.Get(declaration.Value));
            if (atRoot == null && prefixAtRoot == null)
            {
                root.Add(new XAttribute(declaration));
                declaration.Remove();
            }
            else if (atRoot != null && atRoot.Value == declaration.Value)
            {
                declaration.Remove();
            }
        }
    }

    //The namespaces of every element and attribute name, and those named by prefix in markup compatibility attributes other than Ignorable.
    private static HashSet<XNamespace> GetUsedNamespaces(XElement root)
    {
        HashSet<XNamespace> used = new HashSet<XNamespace>();
        foreach (XElement element in root.DescendantsAndSelf())
        {
            used.Add(element.Name.Namespace);
            foreach (XAttribute attribute in element.Attributes())
            {
                if (attribute.IsNamespaceDeclaration)
                {
                    continue;
                }
                used.Add(attribute.Name.Namespace);
                bool namesPrefixes = (element.Name.Namespace == MC && attribute.Name == "Requires")
                    || attribute.Name == MC + "MustUnderstand" || attribute.Name == MC + "ProcessContent"
                    || attribute.Name == MC + "PreserveElements" || attribute.Name == MC + "PreserveAttributes";
                if (namesPrefixes)
                {
                    foreach (string token in attribute.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        int colon = token.IndexOf(':');
                        XNamespace ns = element.GetNamespaceOfPrefix(colon < 0 ? token : token.Substring(0, colon));
                        if (ns != null)
                        {
                            used.Add(ns);
                        }
                    }
                }
            }
        }
        return used;
    }

    //Namespace declarations first (the default one, then by prefix), then the other attributes by qualified name, unprefixed ones first.
    private static void SortAttributes(XElement root)
    {
        foreach (XElement element in root.DescendantsAndSelf())
        {
            List<XAttribute> attributes = element.Attributes().ToList();
            if (attributes.Count > 1)
            {
                List<XAttribute> sorted = attributes.OrderBy(a => GetSortKey(element, a), StringComparer.Ordinal).ToList();
                element.ReplaceAttributes(sorted);
            }
        }
    }

    private static string GetSortKey(XElement element, XAttribute attribute)
    {
        if (attribute.IsNamespaceDeclaration)
        {
            return attribute.Name.Namespace == XNamespace.None ? "0" : "0:" + attribute.Name.LocalName;
        }
        string prefix = attribute.Name.Namespace == XNamespace.None ? string.Empty : element.GetPrefixOfNamespace(attribute.Name.Namespace) ?? string.Empty;
        return "1" + prefix + ":" + attribute.Name.LocalName;
    }

    private static string Write(IEnumerable<XElement> elements)
    {
        XmlWriterSettings settings = new XmlWriterSettings()
        {
            Indent = true,
            IndentChars = "  ",
            NewLineChars = "\n",
            NewLineHandling = NewLineHandling.Replace,
            OmitXmlDeclaration = true,
            ConformanceLevel = ConformanceLevel.Fragment,
        };
        StringBuilder builder = new StringBuilder();
        using (StringWriter stringWriter = new StringWriter(builder, CultureInfo.InvariantCulture))
        {
            foreach (XElement element in elements)
            {
                using (XmlWriter writer = XmlWriter.Create(stringWriter, settings))
                {
                    element.WriteTo(writer);
                }
                stringWriter.Write('\n');
            }
        }
        return builder.ToString();
    }

    private sealed class Placeholders
    {
        private readonly Dictionary<string, Dictionary<string, string>> _byKind = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
        private readonly Dictionary<string, Dictionary<string, string>> _relationshipIds = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);

        //Relationship ids belong to the part that owns the relationships, so each part has its own numbering.
        public Dictionary<string, string> ForRelationshipsOf(string entryName)
        {
            if (!_relationshipIds.TryGetValue(entryName, out Dictionary<string, string> map))
            {
                map = new Dictionary<string, string>(StringComparer.Ordinal);
                _relationshipIds.Add(entryName, map);
            }
            return map;
        }

        public string Get(string kind, string value)
        {
            if (!_byKind.TryGetValue(kind, out Dictionary<string, string> map))
            {
                map = new Dictionary<string, string>(StringComparer.Ordinal);
                _byKind.Add(kind, map);
            }
            return Get(map, kind, value);
        }

        public static string Get(Dictionary<string, string> map, string kind, string value)
        {
            if (!map.TryGetValue(value, out string placeholder))
            {
                placeholder = "{" + kind + (map.Count + 1).ToString(CultureInfo.InvariantCulture) + "}";
                map.Add(value, placeholder);
            }
            return placeholder;
        }
    }
}
