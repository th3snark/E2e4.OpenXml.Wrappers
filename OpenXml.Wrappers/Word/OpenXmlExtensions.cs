using System.Globalization;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Helpers;

namespace OpenXml.Wrappers.Word;

internal static class OpenXmlExtensions
{
    #region NestingLevel Routines

    public static int NestingLevel<T>(this OpenXmlElement openXmlElement) where T : OpenXmlElement
    {
        IList<T> list = openXmlElement.Ancestors<T>().ToList();
        return openXmlElement is T ? list.Count + 1 : list.Count;
    }

    public static IList<T> DescendantsAtRelativeNestingLevel<T>(this OpenXmlElement openXmlElement, int relativeNestingLevel) where T : OpenXmlElement
    {
        return openXmlElement.Descendants<T>().Where(e => e.NestingLevel<T>() == openXmlElement.NestingLevel<T>() + relativeNestingLevel).ToList();
    }

    public static IList<T> AncestorsAtRelativeNestingLevel<T>(this OpenXmlElement openXmlElement, int relativeNestingLevel) where T : OpenXmlElement
    {
        return openXmlElement.Ancestors<T>().Where(e => e.NestingLevel<T>() == openXmlElement.NestingLevel<T>() + relativeNestingLevel).ToList();
    }

    public static T ParentAtRelativeNestingLevel<T>(this OpenXmlElement openXmlElement, int relativeNestingLevel) where T : OpenXmlElement
    {
        IList<T> list = openXmlElement.AncestorsAtRelativeNestingLevel<T>(relativeNestingLevel);
        return list.Count > 0 ? list[0] : null;
    }

    #endregion

    #region Bookmark Routines

    public static BookmarkEnd GetBookmarkEnd(this BookmarkStart bookmarkStart)
    {
        return bookmarkStart.GetRoot().Descendants<BookmarkEnd>().Where(e => e.Id.Value == bookmarkStart.Id.Value).First();
    }

    public static IEnumerable<OpenXmlElement> GetElements(this BookmarkStart bookmarkStart)
    {
        return bookmarkStart.GetElements(bookmarkStart.GetBookmarkEnd());
    }

    public static IEnumerable<T> GetElements<T>(this BookmarkStart bookmarkStart) where T : OpenXmlElement
    {
        return bookmarkStart.GetElements<T>(bookmarkStart.GetBookmarkEnd());
    }

    #endregion

    #region Comment Routines

    public static CommentRangeEnd GetCommentRangeEnd(this CommentRangeStart commentRangeStart)
    {
        return commentRangeStart.GetRoot().Descendants<CommentRangeEnd>().Where(e => e.Id.Value == commentRangeStart.Id.Value).First();
    }

    public static CommentReference GetCommentReference(this CommentRangeStart commentRangeStart)
    {
        Document document = commentRangeStart.GetDocument();
        return document.Descendants<CommentReference>().Where(e => e.Id.Value == commentRangeStart.Id.Value).First();
    }

    public static Comment GetComment(this CommentRangeStart commentRangeStart)
    {
        Document document = commentRangeStart.GetDocument();
        return document.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>().Where(e => e.Id.Value == commentRangeStart.Id.Value).First();
    }

    #endregion

    #region Document Routines

    /// <summary>
    /// Returns the root element of the tree containing the element. Unlike <see cref="GetDocument"/> this also works for
    /// elements in header, footer and comment parts.
    /// </summary>
    public static OpenXmlElement GetRoot(this OpenXmlElement openXmlElement)
    {
        OpenXmlElement root = openXmlElement;
        while (root.Parent != null)
        {
            root = root.Parent;
        }
        return root;
    }

    public static Document GetDocument(this OpenXmlElement openXmlElement)
    {
        if (openXmlElement is Document)
        {
            return (Document)openXmlElement;
        }
        else
        {
            return openXmlElement.Ancestors<Document>().First();
        }
    }

    /// <summary>
    /// Writes the indentation of every <c>w:ind</c> under the root with the attribute names Word writes in a transitional document
    /// (<c>w:left</c>, <c>w:right</c>, <c>w:leftChars</c>, <c>w:rightChars</c>). The SDK reads those into its <c>Start</c>, <c>End</c>,
    /// <c>StartCharacters</c> and <c>EndCharacters</c> properties, so an element it has loaded is written back as <c>w:start</c> and so on.
    /// </summary>
    public static void UseTransitionalIndentationNames(this OpenXmlElement root)
    {
        foreach (Indentation indentation in root.Descendants<Indentation>())
        {
            if (indentation.Start != null && indentation.Left == null)
            {
                indentation.Left = indentation.Start.Value;
                indentation.Start = null;
            }
            if (indentation.End != null && indentation.Right == null)
            {
                indentation.Right = indentation.End.Value;
                indentation.End = null;
            }
            if (indentation.StartCharacters != null && indentation.LeftChars == null)
            {
                indentation.LeftChars = indentation.StartCharacters.Value;
                indentation.StartCharacters = null;
            }
            if (indentation.EndCharacters != null && indentation.RightChars == null)
            {
                indentation.RightChars = indentation.EndCharacters.Value;
                indentation.EndCharacters = null;
            }
        }
    }

    #endregion

    #region GetElements Routines

    public static IEnumerable<OpenXmlElement> GetElements(this OpenXmlElement startOpenXmlElement)
    {
        return startOpenXmlElement.GetElements<OpenXmlElement>();
    }

    public static IEnumerable<T> GetElements<T>(this OpenXmlElement startOpenXmlElement) where T : OpenXmlElement
    {
        if (startOpenXmlElement is T)
        {
            yield return (T)startOpenXmlElement;
        }

        foreach (T openXmlElement in startOpenXmlElement.Descendants<T>())
        {
            yield return openXmlElement;
        }
    }

    public static IEnumerable<OpenXmlElement> GetElements(this OpenXmlElement startOpenXmlElement, OpenXmlElement endOpenXmlElement)
    {
        return startOpenXmlElement.GetElementsBetween<OpenXmlElement>(endOpenXmlElement, true);
    }

    public static IEnumerable<T> GetElements<T>(this OpenXmlElement startOpenXmlElement, OpenXmlElement endOpenXmlElement) where T : OpenXmlElement
    {
        if (startOpenXmlElement.Equals(endOpenXmlElement))
        {
            //A range over a single element contains that element (if it is a T) as well as its descendants.
            foreach (T openXmlElement in startOpenXmlElement.GetElements<T>())
            {
                yield return openXmlElement;
            }
        }
        else
        {
            foreach (T openXmlElement in startOpenXmlElement.GetElementsBetween<T>(endOpenXmlElement, true))
            {
                yield return openXmlElement;
            }
        }
    }

    #endregion

    #region GetElementsBetween and GetLowestCommonAncestor Routines

    public static OpenXmlElement GetLowestCommonAncestor(this OpenXmlElement firstOpenXmlElement, OpenXmlElement secondOpenXmlElement)
    {
        if (firstOpenXmlElement.Equals(secondOpenXmlElement))
        {
            return firstOpenXmlElement;
        }
        else
        {
            return firstOpenXmlElement.Ancestors().Intersect(secondOpenXmlElement.Ancestors()).First();
        }
    }

    public static IEnumerable<T> GetElementsBetween<T>(this OpenXmlElement startOpenXmlElement, OpenXmlElement endOpenXmlElement, bool includeFirstAncestor) where T : OpenXmlElement
    {
        if (includeFirstAncestor)
        {
            T ancestor = startOpenXmlElement.Ancestors<T>().FirstOrDefault();
            if (ancestor != default(T))
            {
                yield return ancestor;
            }
        }

        foreach (OpenXmlElement openXmlElement in startOpenXmlElement.GetElementsBetween(endOpenXmlElement))
        {
            T defaultT = openXmlElement as T;
            if (defaultT != null)
            {
                yield return defaultT;
            }
        }
    }

    private static IEnumerable<OpenXmlElement> GetElementsBetween(this OpenXmlElement startOpenXmlElement, OpenXmlElement endOpenXmlElement)
    {
        bool startFound = false;
        foreach (OpenXmlElement openXmlElement in startOpenXmlElement.GetLowestCommonAncestor(endOpenXmlElement).Descendants())
        {
            if (openXmlElement == startOpenXmlElement)
            {
                startFound = true;
            }

            if (startFound)
            {
                yield return openXmlElement;

                if (openXmlElement == endOpenXmlElement)
                {
                    yield break;
                }
            }
        }
    }

    #endregion

    #region SectionProperties Routines

    public static Columns GetColumns(this SectionProperties sectionProperties)
    {
        return sectionProperties.Elements<Columns>().FirstOrDefault();
    }

    public static void SetNumAndSpaceColumns(this Columns columns, int num, int space)
    {
        columns.ColumnCount = Convert.ToInt16(num);
        columns.Space = Convert.ToString(space, CultureInfo.InvariantCulture);
    }

    public static void ConfigureColumns(this SectionProperties sectionProperties, int num, int space)
    {
        Columns columns = sectionProperties.GetColumns();
        if (columns == null)
        {
            columns = new Columns();
            sectionProperties.AppendChild(columns);
        }
        columns.SetNumAndSpaceColumns(num, space);
    }

    #endregion

    #region Anchor Routines

    /// <summary>
    /// Removes the bookmark and comment anchors from within the element. This is used on cloned content so that the clone does not
    /// duplicate the ids of the anchors in the original (the ids must be unique in a document).
    /// </summary>
    private static bool IsIdentity(OpenXmlAttribute attribute)
    {
        return (attribute.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordml" && (attribute.LocalName == "paraId" || attribute.LocalName == "textId"))
            || (attribute.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main" && attribute.LocalName.StartsWith("rsid", StringComparison.Ordinal));
    }

    /// <summary>
    /// The roots of the parts of the document that hold text: the main document, headers and footers, notes and comments.
    /// </summary>
    internal static IEnumerable<OpenXmlElement> GetStoryRoots(this MainDocumentPart mainDocumentPart)
    {
        yield return mainDocumentPart.Document;
        foreach (HeaderPart part in mainDocumentPart.HeaderParts)
        {
            yield return part.Header;
        }
        foreach (FooterPart part in mainDocumentPart.FooterParts)
        {
            yield return part.Footer;
        }
        if (mainDocumentPart.FootnotesPart?.Footnotes != null)
        {
            yield return mainDocumentPart.FootnotesPart.Footnotes;
        }
        if (mainDocumentPart.EndnotesPart?.Endnotes != null)
        {
            yield return mainDocumentPart.EndnotesPart.Endnotes;
        }
        if (mainDocumentPart.WordprocessingCommentsPart?.Comments != null)
        {
            yield return mainDocumentPart.WordprocessingCommentsPart.Comments;
        }
    }

    /// <summary>
    /// Gives the drawings under the element new ids, which must be unique across the document, for a copy that has just gone into it.
    /// </summary>
    internal static void GiveDrawingsNewIds(this OpenXmlElement element)
    {
        List<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties> drawings = element.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().ToList();
        MainDocumentPart mainDocumentPart = StyleHelper.GetMainDocumentPart(element);
        if (drawings.Count == 0 || mainDocumentPart == null)
        {
            return;
        }
        uint id = mainDocumentPart.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            .Select(d => d.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max();
        foreach (DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties drawing in drawings)
        {
            drawing.Id = ++id;
        }
    }

    internal static void RemoveAnchors(this OpenXmlElement openXmlElement)
    {
        //What makes the original unique in the document does not belong on a copy.
        foreach (OpenXmlElement element in openXmlElement.Descendants().Concat(new[] { openXmlElement }).ToList())
        {
            foreach (OpenXmlAttribute attribute in element.GetAttributes().Where(a => IsIdentity(a)).ToList())
            {
                element.RemoveAttribute(attribute.LocalName, attribute.NamespaceUri);
            }
        }

        IList<OpenXmlElement> anchors = openXmlElement
            .Descendants()
            .Where(e => e is BookmarkStart || e is BookmarkEnd || e is CommentRangeStart || e is CommentRangeEnd || e is CommentReference)
            .ToList();
        foreach (OpenXmlElement anchor in anchors)
        {
            //A comment reference lives in a run of its own, so remove the run as well if nothing else is left in it.
            OpenXmlElement parent = anchor.Parent;
            anchor.Remove();
            if (anchor is CommentReference && parent is Run && !parent.Elements().Any(e => !(e is RunProperties)))
            {
                parent.Remove();
            }
        }
    }

    #endregion

    #region Run Routines

    /// <summary>
    /// A text element as Word writes it: <c>xml:space="preserve"</c> only when the text starts or ends with white space, which would otherwise
    /// be lost.
    /// </summary>
    internal static Text NewText(string value)
    {
        Text text = new Text(value);
        SetSpacePreservation(text);
        return text;
    }

    /// <summary>Gives a text element <c>xml:space="preserve"</c> if its text needs it, and takes it away if it does not.</summary>
    internal static void SetSpacePreservation(Text text)
    {
        string value = text.Text;
        bool needed = value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]));
        text.Space = needed ? SpaceProcessingModeValues.Preserve : (EnumValue<SpaceProcessingModeValues>)null;
    }

    /// <summary>The value of an on and off attribute that is on, as Word writes it (<c>"1"</c>).</summary>
    internal static OnOffValue OnValue()
    {
        return new OnOffValue() { InnerText = "1" };
    }

    /// <summary>The value of an on and off property that is off, as Word writes it (<c>w:val="0"</c>).</summary>
    internal static OnOffValue OffValue()
    {
        return new OnOffValue() { InnerText = "0" };
    }

    private sealed class SiblingCursor
    {
        public OpenXmlElement Checkpoint;
    }

    private static readonly ConditionalWeakTable<OpenXmlElement, SiblingCursor> SiblingCursors = new ConditionalWeakTable<OpenXmlElement, SiblingCursor>();

    //Parents with no more children before the element than this are not worth remembering anything about.
    private const int FewSiblings = 16;

    //How far behind the element last looked for the checkpoint stays, so that the next few looked for, on either side of it, are close by.
    private const int CheckpointBehind = 32;

    /// <summary>
    /// The element before this one under the same parent, as <c>PreviousSibling()</c>. Children only link forwards, so that walks from the first
    /// child, and looking back from each row of a long table in turn takes time that grows with the square of its length. This walks from a
    /// checkpoint kept for the parent instead, a little before the element last looked for. The checkpoint is only a place to start: a change
    /// to the children at worst means walking from the first child, as <c>PreviousSibling()</c> does.
    /// </summary>
    internal static OpenXmlElement PreviousSiblingNear(this OpenXmlElement element)
    {
        OpenXmlElement parent = element.Parent;
        OpenXmlElement first = parent?.FirstChild;
        if (first == null || ReferenceEquals(first, element))
        {
            return null;
        }

        OpenXmlElement near = first;
        for (int i = 0; i < FewSiblings; i++)
        {
            OpenXmlElement next = near.NextSibling();
            if (ReferenceEquals(next, element))
            {
                return near;
            }
            near = next;
        }

        SiblingCursor cursor = SiblingCursors.GetOrCreateValue(parent);
        OpenXmlElement checkpoint = cursor.Checkpoint;
        OpenXmlElement previous = null;
        if (checkpoint != null && ReferenceEquals(checkpoint.Parent, parent))
        {
            previous = WalkToPrevious(checkpoint, element, cursor);
        }
        return previous ?? WalkToPrevious(near, element, cursor);
    }

    //Walks forward from an element for the one before the given element, or null if the walk comes to the end without finding it.
    private static OpenXmlElement WalkToPrevious(OpenXmlElement from, OpenXmlElement element, SiblingCursor cursor)
    {
        OpenXmlElement behind = from;
        int walked = 0;
        for (OpenXmlElement current = from; current != null; current = current.NextSibling())
        {
            if (ReferenceEquals(current.NextSibling(), element))
            {
                cursor.Checkpoint = behind;
                return current;
            }
            if (++walked > CheckpointBehind)
            {
                behind = behind.NextSibling();
            }
        }
        return null;
    }

    /// <summary>
    /// Merges the runs of the paragraph (or hyperlink) that follow one another and have the same formatting, as Word does, so that editing
    /// does not leave text in more and more pieces. Text and field characters stay as they are: a run that holds anything but text, tabs and
    /// breaks is left alone, and text never joins a tab or break that comes after it, or a page or column break before it, which Word gives
    /// a run of their own. Ranges on the text are moved with it.
    /// </summary>
    internal static void CoalesceRuns(this OpenXmlElement parent)
    {
        Run previous = null;
        foreach (OpenXmlElement child in parent.ChildElements.ToList())
        {
            if (child is Hyperlink)
            {
                child.CoalesceRuns();
            }

            Run run = child as Run;
            if (run != null && previous != null && CanMergeRuns(previous, run))
            {
                MergeRuns(previous, run);
            }
            else
            {
                previous = run;
            }
        }
    }

    private static bool CanMergeRuns(Run first, Run second)
    {
        List<OpenXmlElement> firstContent = first.ChildElements.Where(e => !(e is RunProperties)).ToList();
        List<OpenXmlElement> secondContent = second.ChildElements.Where(e => !(e is RunProperties)).ToList();
        if (firstContent.Count == 0 || secondContent.Count == 0 || !(secondContent[0] is Text))
        {
            return false;
        }
        if (firstContent.Concat(secondContent).Any(e => !(e is Text || e is TabChar || e is Break)))
        {
            return false;
        }
        if (firstContent[firstContent.Count - 1] is Break last && last.Type != null && last.Type.Value != BreakValues.TextWrapping)
        {
            return false;
        }
        return SameProperties(first.RunProperties, second.RunProperties);
    }

    private static void MergeRuns(Run first, Run second)
    {
        List<OpenXmlElement> moved = second.ChildElements.Where(e => !(e is RunProperties)).ToList();
        if (first.LastChild is Text firstText)
        {
            Text secondText = (Text)moved[0];
            int length = firstText.Text.Length;
            firstText.Text += secondText.Text;
            SetSpacePreservation(firstText);
            SpanRegistry.TextJoined(firstText, secondText, length);
            moved.RemoveAt(0);
        }
        foreach (OpenXmlElement element in moved)
        {
            element.Remove();
            first.AppendChild(element);
        }
        second.Remove();
    }

    /// <summary>Whether two sets of properties say the same: the same elements in the same order, with the same attributes.</summary>
    internal static bool SameProperties(OpenXmlElement first, OpenXmlElement second)
    {
        int firstCount = first == null ? 0 : first.ChildElements.Count;
        int secondCount = second == null ? 0 : second.ChildElements.Count;
        if (firstCount != secondCount)
        {
            return false;
        }
        for (int i = 0; i < firstCount; i++)
        {
            OpenXmlElement a = first.ChildElements[i];
            OpenXmlElement b = second.ChildElements[i];
            if (a.GetType() != b.GetType() || !SameAttributes(a, b) || !SameProperties(a, b))
            {
                return false;
            }
        }
        return true;
    }

    private static bool SameAttributes(OpenXmlElement first, OpenXmlElement second)
    {
        IList<OpenXmlAttribute> a = first.GetAttributes();
        IList<OpenXmlAttribute> b = second.GetAttributes();
        return a.Count == b.Count && a.All(x => b.Any(y => y.NamespaceUri == x.NamespaceUri && y.LocalName == x.LocalName && y.Value == x.Value));
    }

    /// <summary>The properties of a paragraph mark that has the formatting of a run.</summary>
    internal static ParagraphMarkRunProperties ToParagraphMarkProperties(this RunProperties runProperties)
    {
        return new ParagraphMarkRunProperties(runProperties.ChildElements.Where(e => !(e is RunPropertiesChange)).Select(e => e.CloneNode(true)));
    }

    /// <summary>Removes paragraph properties that have nothing in them.</summary>
    internal static void RemoveEmptyProperties(this Paragraph paragraph)
    {
        ParagraphProperties properties = paragraph.ParagraphProperties;
        if (properties != null && !properties.HasChildren && !properties.HasAttributes)
        {
            properties.Remove();
        }
    }

    public static bool IsLastWithText(this Run run)
    {
        //If this is the last run in the paragraph holding text (or a tab) return true.
        OpenXmlElement container = (OpenXmlElement)run.Ancestors<Paragraph>().FirstOrDefault() ?? run.Parent;
        Run lastRun = container.Descendants<Run>().LastOrDefault(r => r.Elements<Text>().Any() || r.Elements<TabChar>().Any());
        return lastRun != null && lastRun.Equals(run);
    }

    #endregion

    #region RunProperties Routines

    //The order of the children of the run properties (and paragraph mark run properties) in the schema. Tracked changes of the paragraph
    //mark come before all of them and the tracked change of the properties after all of them.
    private static readonly string[] RunPropertyOrder =
    {
        "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
        "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
        "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
    };

    private static int GetRunPropertyOrder(OpenXmlElement element)
    {
        string name = element.LocalName;
        if (name == "ins" || name == "del" || name == "moveFrom" || name == "moveTo")
        {
            return -1;
        }

        int index = Array.IndexOf(RunPropertyOrder, name);
        return index >= 0 ? index : RunPropertyOrder.Length;
    }

    /// <summary>
    /// Sets one property of run properties (or paragraph mark run properties), replacing any existing property of the same kind and placing
    /// the new one where the schema requires it.
    /// </summary>
    internal static void SetRunProperty(this OpenXmlCompositeElement runProperties, OpenXmlElement property)
    {
        runProperties.RemoveRunProperty(property.LocalName);

        int order = GetRunPropertyOrder(property);
        OpenXmlElement predecessor = runProperties.ChildElements.LastOrDefault(e => GetRunPropertyOrder(e) < order);
        if (predecessor == null)
        {
            runProperties.PrependChild(property);
        }
        else
        {
            predecessor.InsertAfterSelf(property);
        }
    }

    /// <summary>
    /// Removes the properties with the given element name (for example "b" for bold) from run properties or paragraph mark run properties.
    /// </summary>
    internal static void RemoveRunProperty(this OpenXmlCompositeElement runProperties, string localName)
    {
        foreach (OpenXmlElement existing in runProperties.ChildElements.Where(e => e.LocalName == localName).ToList())
        {
            existing.Remove();
        }
    }

    internal static RunProperties GetOrAddRunProperties(this Run run)
    {
        RunProperties runProperties = run.RunProperties;
        if (runProperties == null)
        {
            runProperties = new RunProperties();
            run.PrependChild(runProperties);
        }
        return runProperties;
    }

    internal static ParagraphProperties GetOrAddParagraphProperties(this Paragraph paragraph)
    {
        ParagraphProperties paragraphProperties = paragraph.ParagraphProperties;
        if (paragraphProperties == null)
        {
            paragraphProperties = new ParagraphProperties();
            paragraph.PrependChild(paragraphProperties);
        }
        return paragraphProperties;
    }

    internal static ParagraphMarkRunProperties GetOrAddParagraphMarkRunProperties(this Paragraph paragraph)
    {
        ParagraphProperties paragraphProperties = paragraph.GetOrAddParagraphProperties();

        ParagraphMarkRunProperties paragraphMarkRunProperties = paragraphProperties.ParagraphMarkRunProperties;
        if (paragraphMarkRunProperties == null)
        {
            //The paragraph mark run properties must come before the section properties and the paragraph properties change.
            paragraphMarkRunProperties = new ParagraphMarkRunProperties();
            OpenXmlElement successor = paragraphProperties.ChildElements.FirstOrDefault(e => e is SectionProperties || e is ParagraphPropertiesChange);
            if (successor != null)
            {
                successor.InsertBeforeSelf(paragraphMarkRunProperties);
            }
            else
            {
                paragraphProperties.AppendChild(paragraphMarkRunProperties);
            }
        }
        return paragraphMarkRunProperties;
    }

    #endregion

    #region Table Routines

    public static void SetTableGrid(this Table table, TableGrid newTableGrid)
    {
        //Set the table grid to the new table grid replacing the old one if one exists.
        TableGrid oldTableGrid = table.Elements<TableGrid>().FirstOrDefault();
        if (oldTableGrid == null)
        {
            TableProperties tableProperties = table.Elements<TableProperties>().FirstOrDefault();
            if (tableProperties == null)
            {
                table.PrependChild<TableGrid>(newTableGrid);
            }
            else
            {
                tableProperties.InsertAfterSelf<TableGrid>(newTableGrid);
            }
        }
        else
        {
            table.ReplaceChild<TableGrid>(newTableGrid, oldTableGrid);
        }
    }

    public static int GetTableWidth(this Table table)
    {
        int width = 0;
        TableProperties tableProperties = table.Elements<TableProperties>().FirstOrDefault();
        if (tableProperties != null)
        {
            TableWidth tableWidth = tableProperties.Elements<TableWidth>().FirstOrDefault();
            if (tableWidth != null)
            {
                //The schema default for the width type is dxa.
                TableWidthUnitValues widthType = tableWidth.Type != null && tableWidth.Type.HasValue ? tableWidth.Type.Value : TableWidthUnitValues.Dxa;
                if (widthType == TableWidthUnitValues.Dxa)
                {
                    width = Convert.ToInt32(tableWidth.Width.Value, CultureInfo.InvariantCulture);
                }
                else if (widthType == TableWidthUnitValues.Auto)
                {
                    //The columns say how wide the table is.
                    width = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Sum(c => int.TryParse(c.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int column) ? column : 0) ?? 0;
                }
                else if (widthType == TableWidthUnitValues.Pct)
                {
                    //Find the outermost table (just in case this is a nested table).
                    Table outerTable = table.Ancestors<Table>().LastOrDefault();
                    if (outerTable == null)
                    {
                        outerTable = table;
                    }

                    //Find the next sibling section properties element from the outermost table. This provides the page size and page
                    //margins of the page that the table is on.
                    //If the table is not the last element in its section then fall back to the body level section properties.
                    SectionProperties sectionProperties = outerTable.NextSibling<SectionProperties>()
                        ?? outerTable.Ancestors<Body>().FirstOrDefault()?.Elements<SectionProperties>().LastOrDefault();
                    if (sectionProperties == null)
                    {
                        return 0;
                    }
                    PageSize pageSize = sectionProperties.Elements<PageSize>().First();
                    PageMargin pageMargin = sectionProperties.Elements<PageMargin>().First();
                    int totalPageMarginWidth = Convert.ToInt32(pageMargin.Left.Value) + Convert.ToInt32(pageMargin.Right.Value);

                    //Get the table cell margin for the table. These are required to workout the width of the table on the page.
                    TableCellMargin tableCellMargin = table.GetTableCellMargin();
                    int totalTableCellMarginWidth = Convert.ToInt32(tableCellMargin.LeftMargin.Width, CultureInfo.InvariantCulture) + Convert.ToInt32(tableCellMargin.RightMargin.Width, CultureInfo.InvariantCulture);

                    width = Convert.ToInt32(pageSize.Width.Value) - totalPageMarginWidth + totalTableCellMarginWidth;
                }
                //Any other width type (auto or nil) has no fixed width, so zero is returned.
            }
        }
        return width;
    }

    public static TableCellMargin GetTableCellMargin(this Table table)
    {
        //Assign a default table cell margin. This appears to be what Word 2007 does. I would have expected Word
        //to look at the default table style for this however when this was changed in a test document it seemed
        //to have no affect on the setting in Word?
        TableCellMargin tableCellMargin = new TableCellMargin(
           new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
           new LeftMargin() { Width = "108", Type = TableWidthUnitValues.Dxa },
           new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
           new RightMargin() { Width = "108", Type = TableWidthUnitValues.Dxa });

        TableProperties tableProperties = table.Elements<TableProperties>().FirstOrDefault();
        if (tableProperties != null)
        {
            //Try to override the table cell margin with the one defined on the tables style definition if one is set.
            TableStyle tableStyle = tableProperties.Elements<TableStyle>().FirstOrDefault();
            if (tableStyle != null)
            {
                //Get the style definition part from the document.
                StyleDefinitionsPart styleDefinitionsPart = table.GetDocument().MainDocumentPart.StyleDefinitionsPart;
                if (styleDefinitionsPart != null)
                {
                    string tableStyleId = tableStyle.Val;
                    Style style = styleDefinitionsPart.Styles.Elements<Style>().Where(e => e.Type.Value == StyleValues.Table && e.StyleId == tableStyleId).FirstOrDefault();
                    if (style != null)
                    {
                        TableProperties styelTableProperties = style.Elements<TableProperties>().FirstOrDefault();
                        if (styelTableProperties != null)
                        {
                            TableCellMargin styleTableCellMargin = styelTableProperties.Elements<TableCellMargin>().FirstOrDefault();
                            if (styleTableCellMargin != null)
                            {
                                tableCellMargin = styleTableCellMargin;
                            }
                        }
                    }
                }
            }

            //Try to override the table cell margin with the one defined on the table itself.
            TableCellMargin localTableCellMargin = tableProperties.Elements<TableCellMargin>().FirstOrDefault();
            if (localTableCellMargin != null)
            {
                tableCellMargin = localTableCellMargin;
            }
        }

        return tableCellMargin;
    }

    public static void MergeTable(this Table table1, Table table2)
    {
        List<int> widthList = table1.GetWidthList(table2);
        List<int> widthList1 = table1.GetWidthList();
        List<int> widthList2 = table2.GetWidthList();

        TableGrid newTableGrid = new TableGrid();
        foreach (int width in widthList)
        {
            newTableGrid.AppendChild<GridColumn>(new GridColumn() { Width = width.ToString(CultureInfo.InvariantCulture) });
        }

        Table table = (Table)table1.Clone();
        table.RemoveAllChildren<TableRow>();
        table.SetTableGrid(newTableGrid);

        int tableWidth = table1.GetTableWidth();
        foreach (TableRow tableRow1 in table1.Elements<TableRow>())
        {
            TableRow tableRow = (TableRow)tableRow1.Clone();
            table.AppendChild<TableRow>(tableRow);

            tableRow.SetAllTableCellWidthsGridSpansAndGridAfters(tableWidth, widthList, widthList1, tableRow1);
        }

        //Get the table borders from the table properties of the second table. These will be cloned and added to each cloned row of the second
        //table (as a table row properties exception) when it is appended to the merged table.
        TableBorders tableBorders2 = table2.GetTableBorders();
        foreach (TableRow tableRow2 in table2.Elements<TableRow>())
        {
            TableRow tableRow = (TableRow)tableRow2.Clone();
            table.AppendChild<TableRow>(tableRow);

            tableRow.SetAllTableCellWidthsGridSpansAndGridAfters(tableWidth, widthList, widthList2, tableRow2);

            //If the second table had borders set on the table element then clone these borders and apply as a table properties exception to
            //each row as it is cloned and appended to the merged table.
            if (tableBorders2 != null)
            {
                tableRow.SetTablePropertiesExceptionTableBorders((TableBorders)tableBorders2.Clone());
            }
        }

        //Insert the newly merged table and remove the original two tables.
        table1.InsertBeforeSelf<Table>(table);
        table1.Remove();
        table2.Remove();
    }

    private static List<int> GetWidthList(this Table table)
    {
        List<int> widthList = new List<int>();
        TableGrid tableGrid = table.Elements<TableGrid>().First();
        foreach (GridColumn gridColumn in tableGrid.Elements<GridColumn>())
        {
            widthList.Add(Convert.ToInt32(gridColumn.Width, CultureInfo.InvariantCulture));
        }
        return widthList;
    }

    private static List<int> GetWidthList(this Table table1, Table table2)
    {
        //Construct the new logical table grid of the combined tables
        TableGrid tableGrid1 = table1.Elements<TableGrid>().First();
        TableGrid tableGrid2 = table2.Elements<TableGrid>().First();

        List<int> offsetList = new List<int>();
        int total1 = 0;
        foreach (GridColumn gridColumn in tableGrid1.Elements<GridColumn>())
        {
            total1 += Convert.ToInt32(gridColumn.Width, CultureInfo.InvariantCulture);
            offsetList.Add(total1);
        }

        int total2 = 0;
        foreach (GridColumn gridColumn in tableGrid2.Elements<GridColumn>())
        {
            total2 += Convert.ToInt32(gridColumn.Width, CultureInfo.InvariantCulture);
            offsetList.Add(total2);
        }

        //Remove duplicates and sort ascending
        offsetList = offsetList.Distinct().ToList();
        offsetList.Sort();

        List<int> widthList = new List<int>();
        int lastOffset = 0;
        foreach (int offset in offsetList)
        {
            widthList.Add(offset - lastOffset);
            lastOffset = offset;
        }

        return widthList;
    }

    private static TableBorders GetTableBorders(this Table table)
    {
        TableBorders tableBorders = null;
        TableProperties tableProperties = table.Elements<TableProperties>().FirstOrDefault();
        if (tableProperties != null)
        {
            tableBorders = tableProperties.Elements<TableBorders>().FirstOrDefault();
        }
        return tableBorders;
    }

    #endregion

    #region TableCell Routines

    private static int GetWidth(this TableCell tableCell, List<int> widthList)
    {
        int tableGridIndex = tableCell.GetTableGridIndex();
        int span = tableCell.GetGridSpan();
        int width = 0;
        for (int i = tableGridIndex; i < tableGridIndex + span; i++)
        {
            width += widthList[i];
        }
        return width;
    }

    private static int GetTableGridIndex(this TableCell tableCell)
    {
        int tableGridIndex = 0;
        foreach (TableCell precedingTableCell in tableCell.ElementsBefore().Where(element => element is TableCell))
        {
            tableGridIndex += precedingTableCell.GetGridSpan();
        }
        return tableGridIndex;
    }

    private static int GetGridSpan(this TableCell tableCell)
    {
        int span = 1;
        TableCellProperties tableCellProperties = tableCell.Elements<TableCellProperties>().FirstOrDefault();
        if (tableCellProperties != null)
        {
            GridSpan gridSpan = tableCellProperties.Elements<GridSpan>().FirstOrDefault();
            if (gridSpan != null)
            {
                span = gridSpan.Val;
            }
        }
        return span;
    }

    private static int GetEndIndex(this TableCell tableCell, int startIndex, int width, List<int> widthList)
    {
        int index = startIndex;
        for (int i = startIndex; i < widthList.Count && width > 0; i++)
        {
            width -= widthList[i];
            index++;
        }
        return index - 1;
    }

    internal static void SetWidth(this TableCell tableCell, int width, TableWidthUnitValues unit)
    {
        TableCellProperties tableCellProperties = tableCell.Elements<TableCellProperties>().FirstOrDefault();
        if (tableCellProperties == null)
        {
            tableCellProperties = new TableCellProperties();
            tableCell.PrependChild<TableCellProperties>(tableCellProperties);
        }

        tableCellProperties.TableCellWidth = new TableCellWidth() { Width = width.ToString(CultureInfo.InvariantCulture), Type = unit };
    }

    internal static void SetGridSpan(this TableCell tableCell, int span)
    {
        if (span < 1)
        {
            throw new ArgumentException("The span argument must be greater than 0.");
        }

        TableCellProperties tableCellProperties = tableCell.Elements<TableCellProperties>().FirstOrDefault();
        GridSpan oldGridSpan = null;
        if (tableCellProperties != null)
        {
            oldGridSpan = tableCellProperties.Elements<GridSpan>().FirstOrDefault();
        }

        if (span == 1)
        {
            //If the span is one then remove the gridSpan element if one exists.
            if (oldGridSpan != null)
            {
                oldGridSpan.Remove();
            }
        }
        else // span is greater than one
        {
            if (oldGridSpan != null)
            {
                //Set the value on the existing grid span.
                oldGridSpan.Val = span;
            }
            else
            {
                //The typed property puts it where the schema requires.
                if (tableCellProperties == null)
                {
                    tableCellProperties = new TableCellProperties();
                    tableCell.PrependChild<TableCellProperties>(tableCellProperties);
                }
                tableCellProperties.GridSpan = new GridSpan() { Val = span };
            }
        }
    }

    public static void ClearText(this TableCell tableCell)
    {
        TableCellProperties tableCellProperties = tableCell.TableCellProperties;
        if (tableCellProperties != null)
        {
            //Clear any vertical merge information from the table cell properties
            VerticalMerge verticalMerge = tableCellProperties.Elements<VerticalMerge>().FirstOrDefault();
            if (verticalMerge != null)
            {
                verticalMerge.Parent.RemoveChild(verticalMerge);
            }
        }

        //Clear all children from the first paragraph but preserve the paragraph properties if one exists.
        Paragraph firstParagraph = tableCell.Elements<Paragraph>().FirstOrDefault();
        if (firstParagraph != null)
        {
            ParagraphProperties paragraphProperties = firstParagraph.ParagraphProperties;
            firstParagraph.RemoveAllChildren();
            if (paragraphProperties != null)
            {
                firstParagraph.ParagraphProperties = paragraphProperties;
            }
        }

        //Clear all children of the table cell then add back in the table cell properties iand first paragraph if we have them.
        tableCell.RemoveAllChildren();
        if (tableCellProperties != null)
        {
            tableCell.TableCellProperties = tableCellProperties;
        }
        //A table cell must always end with a paragraph so add an empty one if the cell did not have any.
        tableCell.AppendChild(firstParagraph ?? new Paragraph());
    }

    #endregion

    #region TableRow Routines

    public static void ClearText(this TableRow tableRow)
    {
        foreach (TableCell tableCell in tableRow.Elements<TableCell>())
        {
            tableCell.ClearText();
        }
    }

    public static int GetIndex(this TableRow tableRow)
    {
        return tableRow.ElementsBefore().Where(element => element is TableRow).Count();
    }

    private static void SetTablePropertiesExceptionTableBorders(this TableRow tableRow, TableBorders tableBorders)
    {
        TablePropertyExceptions tablePropertyExceptions = tableRow.Elements<TablePropertyExceptions>().FirstOrDefault();
        if (tablePropertyExceptions == null)
        {
            tablePropertyExceptions = new TablePropertyExceptions();
            tableRow.PrependChild<TablePropertyExceptions>(tablePropertyExceptions);
        }
        TableBorders oldTableBorders = tablePropertyExceptions.Elements<TableBorders>().FirstOrDefault();
        if (oldTableBorders != null)
        {
            oldTableBorders.Remove();
        }
        tablePropertyExceptions.AppendChild<TableBorders>(tableBorders);
    }

    private static void SetAllTableCellWidthsGridSpansAndGridAfters(this TableRow tableRow, int tableWidth, List<int> newWidthList, List<int> oldWidthList, TableRow oldTableRow)
    {
        int startIndex = 0;
        int endIndex = 0;

        tableRow.RemoveAllChildren<TableCell>();
        foreach (TableCell oldTableCell in oldTableRow.Elements<TableCell>())
        {
            TableCell newTableCell = (TableCell)oldTableCell.Clone();
            tableRow.AppendChild<TableCell>(newTableCell);

            int width = oldTableCell.GetWidth(oldWidthList);
            endIndex = newTableCell.GetEndIndex(startIndex, width, newWidthList);

            newTableCell.SetGridSpan((endIndex - startIndex) + 1);
            newTableCell.SetWidth(ConversionHelper.ToPct(width, tableWidth), TableWidthUnitValues.Pct);

            startIndex = endIndex + 1;
        }

        int afterSpan = (newWidthList.Count - 1) - endIndex;
        if (afterSpan > 0)
        {
            int afterWidth = newWidthList.IndexRange(startIndex, startIndex + afterSpan - 1).Sum();
            int afterWidthPct = ConversionHelper.ToPct(afterWidth, tableWidth);

            //Ensure afterWidthPct is never less than 1.
            afterWidthPct = afterWidthPct > 1 ? afterWidthPct : 1;

            tableRow.SetGridAfter(afterSpan, afterWidthPct, TableWidthUnitValues.Pct);
        }
    }

    private static void SetGridAfter(this TableRow tableRow, int span, int width, TableWidthUnitValues units)
    {
        if (span < 1)
        {
            throw new ArgumentException("The span argument must be greater than 0.");
        }

        if (width < 1)
        {
            throw new ArgumentException("The width argument must be greater than 0.");
        }

        TableRowProperties tableRowProperties = tableRow.Elements<TableRowProperties>().FirstOrDefault();
        GridAfter oldGridAfter = null;
        WidthAfterTableRow oldWidthAfterTableRow = null;
        if (tableRowProperties != null)
        {
            oldGridAfter = tableRowProperties.Elements<GridAfter>().FirstOrDefault();
            oldWidthAfterTableRow = tableRowProperties.Elements<WidthAfterTableRow>().FirstOrDefault();
        }

        GridAfter newGridAfter = null;
        if (oldGridAfter != null)
        {
            oldGridAfter.Val = span;
        }
        else
        {
            if (tableRowProperties == null)
            {
                tableRowProperties = new TableRowProperties();
                tableRow.PrependChild<TableRowProperties>(tableRowProperties);
            }
            newGridAfter = new GridAfter() { Val = span };
        }

        WidthAfterTableRow newWidthAfterTableRow = null;
        if (oldWidthAfterTableRow != null)
        {
            oldWidthAfterTableRow.Width = width.ToString(CultureInfo.InvariantCulture);
            oldWidthAfterTableRow.Type = units;
        }
        else
        {
            if (tableRowProperties == null)
            {
                tableRowProperties = new TableRowProperties();
                tableRow.PrependChild<TableRowProperties>(tableRowProperties);
            }
            newWidthAfterTableRow = new WidthAfterTableRow() { Width = width.ToString(CultureInfo.InvariantCulture), Type = units };
        }

        if (newGridAfter != null)
        {
            tableRowProperties.PrependChild<GridAfter>(newGridAfter);
            if (newWidthAfterTableRow != null)
            {
                newGridAfter.InsertAfterSelf<WidthAfterTableRow>(newWidthAfterTableRow);
            }
        }
        else if (newWidthAfterTableRow != null)
        {
            //The grid after already existed so attach the new width after it.
            oldGridAfter.InsertAfterSelf<WidthAfterTableRow>(newWidthAfterTableRow);
        }
    }

    #endregion

    #region TableCell Routines

    public static int GetIndex(this TableCell tableCell)
    {
        return tableCell.ElementsBefore().Where(element => element is TableCell).Count();
    }

    #endregion

    #region Paragraph Routines

    public static void PrependRun(this Paragraph paragraph, Run run)
    {
        if (paragraph.ParagraphProperties != null)
            paragraph.ParagraphProperties.InsertAfterSelf(run);
        else
            paragraph.PrependChild(run);
    }

    public static void SetText(this Paragraph paragraph, string text)
    {
        paragraph.RemoveAllChildren<Run>();

        text = text ?? string.Empty;
        StringBuilder sb = new StringBuilder();
        Run run = null;
        foreach (char c in text.AsEnumerable().Reverse())
        {
            switch (c)
            {
                case '\t':
                    run = new Run(new TabChar());
                    if (sb.Length > 0)
                    {
                        run.AppendChild(NewText(sb.ToString()));
                        sb.Length = 0;
                    }
                    paragraph.PrependRun(run);
                    break;

                case '\n':
                case '\v':
                    run = new Run(new Break());
                    if (sb.Length > 0)
                    {
                        run.AppendChild(NewText(sb.ToString()));
                        sb.Length = 0;
                    }
                    paragraph.PrependRun(run);
                    break;

                default:
                    //Always insert characters to the front since we are iterating the string in reverse.
                    sb.Insert(0, c);
                    break;
            }
        }
        if (sb.Length > 0)
        {
            run = new Run(NewText(sb.ToString()));
            paragraph.PrependRun(run);
        }
    }

    #endregion

    #region DateTime Routines

    /// <summary>
    /// <para>Truncates a DateTime to a specified resolution.</para>
    /// <para>A convenient source for resolution is TimeSpan.TicksPerXXXX constants.</para>
    /// </summary>
    /// <param name="date">The DateTime object to truncate</param>
    /// <param name="resolution">e.g. to round to nearest second, TimeSpan.TicksPerSecond</param>
    /// <returns>Truncated DateTime</returns>
    public static DateTime Truncate(this DateTime date, long resolution)
    {
        Guard.ThrowIfLessThanOrEqual(resolution, 0L);
        return new DateTime(date.Ticks - (date.Ticks % resolution), date.Kind);
    }

    #endregion

    public static IEnumerable<TSource> IndexRange<TSource>(this IList<TSource> source, int startIndex, int endIndex)
    {
        int index = startIndex;
        while (index <= endIndex)
        {
            yield return source[index];
            index++;
        }
    }
}
