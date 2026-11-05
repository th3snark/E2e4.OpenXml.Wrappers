using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OpenXml.Wrappers.Test.Utils;

/// <summary>
/// Checks a document for what Word needs beyond schema validity: references that resolve and ids that are unique. Each check returns a list
/// of problems, empty when the document passes.
/// </summary>
/// <remarks>
/// The schema check is the single validator configuration for the tests (<see cref="FileFormatVersions.Microsoft365"/>); use
/// <see cref="SchemaErrors"/> rather than creating an <see cref="OpenXmlValidator"/>.
/// </remarks>
public static class DocumentIntegrity
{
    /// <summary>Schema validation with <see cref="FileFormatVersions.Microsoft365"/>.</summary>
    public const string Schema = "Schema";

    /// <summary>Every <c>w:pStyle</c>, <c>w:rStyle</c> and <c>w:tblStyle</c> names a style in the styles part.</summary>
    public const string Styles = "Styles";

    /// <summary>Every <c>w:numId</c> (other than 0) names a <c>w:num</c>, and every <c>w:num</c> an abstract numbering.</summary>
    public const string Numbering = "Numbering";

    /// <summary>Every attribute in the relationships namespace (<c>r:id</c>, <c>r:embed</c>, ...) names a relationship of its part.</summary>
    public const string Relationships = "Relationships";

    /// <summary>Bookmark ids are unique across all the parts.</summary>
    public const string BookmarkIds = "BookmarkIds";

    /// <summary><c>w14:paraId</c> values are unique across all the parts and below 0x80000000.</summary>
    public const string ParagraphIds = "ParagraphIds";

    /// <summary><c>wp:docPr</c> ids are unique across all the parts.</summary>
    public const string DrawingIds = "DrawingIds";

    /// <summary><c>commentsExtended</c> (and <c>commentsIds</c>) refer to the last paragraph of a comment, and replies to a comment that is there.</summary>
    public const string Comments = "Comments";

    public static readonly IReadOnlyList<string> All = new[] { Schema, Styles, Numbering, Relationships, BookmarkIds, ParagraphIds, DrawingIds, Comments };

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace W16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>Runs one check (see <see cref="All"/>).</summary>
    public static IList<string> Check(WordprocessingDocument document, string check)
    {
        switch (check)
        {
            case Schema:
                return SchemaErrors(document);
            case Styles:
                return CheckStyles(document);
            case Numbering:
                return CheckNumbering(document);
            case Relationships:
                return CheckRelationships(document);
            case BookmarkIds:
                return Duplicates(document, "bookmark id", e => e.Name == W + "bookmarkStart" ? (string)e.Attribute(W + "id") : null);
            case ParagraphIds:
                return CheckParagraphIds(document);
            case DrawingIds:
                return Duplicates(document, "docPr id", e => e.Name == WP + "docPr" ? (string)e.Attribute("id") : null);
            case Comments:
                return CheckComments(document);
            default:
                throw new ArgumentOutOfRangeException(nameof(check), check, "There is no such check.");
        }
    }

    /// <summary>Every problem every check finds, each prefixed with the check's name.</summary>
    public static IList<string> CheckAll(WordprocessingDocument document)
    {
        return All.SelectMany(c => Check(document, c).Select(p => c + ": " + p)).ToList();
    }

    /// <summary>
    /// The tables of the body wider than the space they have: the text width of their section, or the width of the cell they are in. That is
    /// valid, and Word can do it too, so it is not one of <see cref="All"/>, but a document meant to be read should not.
    /// </summary>
    public static IList<string> TablesWiderThanTheirSpace(WordprocessingDocument document)
    {
        XElement body = XElement.Parse(document.MainDocumentPart.Document.Body.OuterXml);
        List<string> problems = new List<string>();
        foreach (XElement table in body.Descendants(W + "tbl"))
        {
            int columns = table.Element(W + "tblGrid").Elements(W + "gridCol").Sum(c => (int?)c.Attribute(W + "w") ?? 0);
            XElement width = table.Element(W + "tblPr")?.Element(W + "tblW");
            int preferred = (string)width?.Attribute(W + "type") == "dxa" ? (int?)width.Attribute(W + "w") ?? 0 : 0;
            int space = SpaceForTable(table, body);
            if (Math.Max(columns, preferred) > space)
            {
                problems.Add(string.Format(CultureInfo.InvariantCulture, "a table with columns of {0} twips and a width of {1} has only {2}", columns, preferred, space));
            }
        }
        return problems;
    }

    private static int SpaceForTable(XElement table, XElement body)
    {
        XElement cellWidth = table.Ancestors(W + "tc").FirstOrDefault()?.Element(W + "tcPr")?.Element(W + "tcW");
        if ((string)cellWidth?.Attribute(W + "type") == "dxa")
        {
            return (int)cellWidth.Attribute(W + "w");
        }

        //A section ends with the paragraph that holds its properties, and those at the end of the body are of the last section.
        XElement block = table.AncestorsAndSelf().First(e => e.Parent == body);
        XElement section = block.ElementsAfterSelf(W + "p").Select(p => p.Element(W + "pPr")?.Element(W + "sectPr")).FirstOrDefault(s => s != null) ?? body.Element(W + "sectPr");
        XElement margins = section.Element(W + "pgMar");
        return (int)section.Element(W + "pgSz").Attribute(W + "w") - (int)margins.Attribute(W + "left") - (int)margins.Attribute(W + "right") - ((int?)margins.Attribute(W + "gutter") ?? 0);
    }

    /// <summary>The schema errors of the document, as text, with <see cref="FileFormatVersions.Microsoft365"/>.</summary>
    public static IList<string> SchemaErrors(WordprocessingDocument document)
    {
        try
        {
            return new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document)
                .Select(e => string.Format(CultureInfo.InvariantCulture, "{0} at {1} in {2}", e.Description, e.Path?.XPath, e.Part?.Uri))
                .ToList();
        }
        catch (NullReferenceException exception)
        {
            //The validator itself fails on some broken documents, such as a w:headerReference whose r:id names no relationship.
            return new[] { "The validator failed (" + exception.Message + "); see the Relationships check." };
        }
    }

    //The XML of every part, with its URI, read from the part's stream so that the checks see what will be saved.
    private static IEnumerable<KeyValuePair<OpenXmlPart, XElement>> XmlParts(WordprocessingDocument document)
    {
        foreach (OpenXmlPart part in document.GetAllParts())
        {
            if (part.RootElement != null)
            {
                yield return new KeyValuePair<OpenXmlPart, XElement>(part, XElement.Parse(part.RootElement.OuterXml));
            }
        }
    }

    private static IList<string> CheckStyles(WordprocessingDocument document)
    {
        StyleDefinitionsPart stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        HashSet<string> styleIds = new HashSet<string>(stylesPart?.Styles == null
            ? Enumerable.Empty<string>()
            : XElement.Parse(stylesPart.Styles.OuterXml).Elements(W + "style").Select(s => (string)s.Attribute(W + "styleId")), StringComparer.Ordinal);
        List<string> problems = new List<string>();
        foreach (KeyValuePair<OpenXmlPart, XElement> part in XmlParts(document))
        {
            IEnumerable<string> missing = part.Value.Descendants()
                .Where(e => e.Name == W + "pStyle" || e.Name == W + "rStyle" || e.Name == W + "tblStyle")
                .Where(e => !styleIds.Contains((string)e.Attribute(W + "val") ?? string.Empty))
                .Select(e => e.Name.LocalName + " " + (string)e.Attribute(W + "val"))
                .Distinct();
            foreach (string reference in missing)
            {
                problems.Add(string.Format(CultureInfo.InvariantCulture, "{0} in {1} names no style", reference, part.Key.Uri));
            }
        }
        return problems;
    }

    private static IList<string> CheckNumbering(WordprocessingDocument document)
    {
        NumberingDefinitionsPart numberingPart = document.MainDocumentPart?.NumberingDefinitionsPart;
        XElement numbering = numberingPart?.Numbering == null ? new XElement(W + "numbering") : XElement.Parse(numberingPart.Numbering.OuterXml);
        HashSet<string> numIds = new HashSet<string>(numbering.Elements(W + "num").Select(n => (string)n.Attribute(W + "numId")), StringComparer.Ordinal);
        HashSet<string> abstractNumIds = new HashSet<string>(numbering.Elements(W + "abstractNum").Select(n => (string)n.Attribute(W + "abstractNumId")),
            StringComparer.Ordinal);
        List<string> problems = new List<string>();
        foreach (XElement num in numbering.Elements(W + "num"))
        {
            string abstractNumId = (string)num.Element(W + "abstractNumId")?.Attribute(W + "val");
            if (!abstractNumIds.Contains(abstractNumId ?? string.Empty))
            {
                problems.Add(string.Format(CultureInfo.InvariantCulture, "num {0} names no abstractNum ({1})", (string)num.Attribute(W + "numId"), abstractNumId));
            }
        }
        foreach (KeyValuePair<OpenXmlPart, XElement> part in XmlParts(document))
        {
            foreach (string numId in part.Value.Descendants(W + "numId").Select(e => (string)e.Attribute(W + "val")).Distinct())
            {
                if (numId != "0" && !numIds.Contains(numId))
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "numId {0} in {1} names no num", numId, part.Key.Uri));
                }
            }
        }
        return problems;
    }

    private static IList<string> CheckRelationships(WordprocessingDocument document)
    {
        List<string> problems = new List<string>();
        foreach (KeyValuePair<OpenXmlPart, XElement> part in XmlParts(document))
        {
            OpenXmlPart owner = part.Key;
            HashSet<string> ids = new HashSet<string>(owner.Parts.Select(p => p.RelationshipId)
                .Concat(owner.ExternalRelationships.Select(r => r.Id))
                .Concat(owner.HyperlinkRelationships.Select(r => r.Id))
                .Concat(owner.DataPartReferenceRelationships.Select(r => r.Id)), StringComparer.Ordinal);
            foreach (XAttribute attribute in part.Value.DescendantsAndSelf().Attributes().Where(a => a.Name.Namespace == R))
            {
                if (!ids.Contains(attribute.Value))
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "r:{0}=\"{1}\" on {2} in {3} names no relationship", attribute.Name.LocalName,
                        attribute.Value, attribute.Parent.Name.LocalName, owner.Uri));
                }
            }
        }
        return problems;
    }

    private static IList<string> CheckParagraphIds(WordprocessingDocument document)
    {
        List<string> problems = Duplicates(document, "w14:paraId", e => (string)e.Attribute(W14 + "paraId")).ToList();
        foreach (KeyValuePair<OpenXmlPart, XElement> part in XmlParts(document))
        {
            foreach (string paraId in part.Value.Descendants().Select(e => (string)e.Attribute(W14 + "paraId")).Where(v => v != null))
            {
                if (!int.TryParse(paraId, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int value) || value < 0)
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "w14:paraId {0} in {1} is not a hexadecimal number below 80000000", paraId, part.Key.Uri));
                }
            }
        }
        return problems;
    }

    private static IList<string> Duplicates(WordprocessingDocument document, string what, Func<XElement, string> id)
    {
        Dictionary<string, List<string>> places = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        foreach (KeyValuePair<OpenXmlPart, XElement> part in XmlParts(document))
        {
            foreach (string value in part.Value.DescendantsAndSelf().Select(id).Where(v => v != null))
            {
                if (!places.TryGetValue(value, out List<string> uris))
                {
                    uris = new List<string>();
                    places.Add(value, uris);
                }
                uris.Add(part.Key.Uri.ToString());
            }
        }
        return places.Where(p => p.Value.Count > 1)
            .Select(p => string.Format(CultureInfo.InvariantCulture, "{0} {1} is used {2} times ({3})", what, p.Key, p.Value.Count, string.Join(", ", p.Value.Distinct())))
            .ToList();
    }

    private static IList<string> CheckComments(WordprocessingDocument document)
    {
        List<string> problems = new List<string>();
        MainDocumentPart main = document.MainDocumentPart;
        XElement comments = main?.WordprocessingCommentsPart?.Comments == null ? null : XElement.Parse(main.WordprocessingCommentsPart.Comments.OuterXml);
        XElement commentsEx = main?.WordprocessingCommentsExPart?.CommentsEx == null ? null : XElement.Parse(main.WordprocessingCommentsExPart.CommentsEx.OuterXml);
        XElement commentsIds = main?.WordprocessingCommentsIdsPart?.CommentsIds == null ? null : XElement.Parse(main.WordprocessingCommentsIdsPart.CommentsIds.OuterXml);

        HashSet<string> commentIds = new HashSet<string>(comments == null ? Enumerable.Empty<string>() : comments.Elements(W + "comment").Select(c => (string)c.Attribute(W + "id")),
            StringComparer.Ordinal);
        if (main?.Document != null)
        {
            foreach (string id in XElement.Parse(main.Document.OuterXml).Descendants(W + "commentReference").Select(e => (string)e.Attribute(W + "id")).Distinct())
            {
                if (!commentIds.Contains(id))
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "commentReference {0} names no comment", id));
                }
            }
        }

        //Word links the extended information to a comment through the paraId of the comment's last paragraph.
        HashSet<string> lastParagraphIds = new HashSet<string>(comments == null
            ? Enumerable.Empty<string>()
            : comments.Elements(W + "comment").Select(c => (string)c.Elements(W + "p").LastOrDefault()?.Attribute(W14 + "paraId")).Where(v => v != null),
            StringComparer.Ordinal);
        if (commentsEx != null)
        {
            List<XElement> entries = commentsEx.Elements(W15 + "commentEx").ToList();
            HashSet<string> entryIds = new HashSet<string>(entries.Select(e => (string)e.Attribute(W15 + "paraId")), StringComparer.Ordinal);
            foreach (XElement entry in entries)
            {
                string paraId = (string)entry.Attribute(W15 + "paraId");
                string parent = (string)entry.Attribute(W15 + "paraIdParent");
                if (!lastParagraphIds.Contains(paraId ?? string.Empty))
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "commentEx {0} is not the paraId of the last paragraph of a comment", paraId));
                }
                if (parent != null && (!entryIds.Contains(parent) || parent == paraId))
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "commentEx {0} replies to {1}, which is not another comment", paraId, parent));
                }
            }
        }
        if (commentsIds != null)
        {
            foreach (string paraId in commentsIds.Elements(W16cid + "commentId").Select(e => (string)e.Attribute(W16cid + "paraId")))
            {
                if (!lastParagraphIds.Contains(paraId ?? string.Empty))
                {
                    problems.Add(string.Format(CultureInfo.InvariantCulture, "commentId {0} is not the paraId of the last paragraph of a comment", paraId));
                }
            }
        }
        return problems;
    }
}
