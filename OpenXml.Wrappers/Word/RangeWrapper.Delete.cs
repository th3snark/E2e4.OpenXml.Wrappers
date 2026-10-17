using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// Deletes the content of the range.
    /// <list type="bullet">
    /// <item><description>A paragraph, table row or table is removed. Where the document needs it to stay (the only paragraph of a table cell, or
    /// the last paragraph of the body after a table) the paragraph is emptied instead, and deleting the last row of a table removes the table.</description></item>
    /// <item><description>A table cell, a header or footer, and the whole document, are emptied but remain (with one empty paragraph).</description></item>
    /// <item><description>A bookmark or comment range loses the runs inside it. The bookmark or comment itself is kept.</description></item>
    /// </list>
    /// Bookmarks and comments that are anchored in the deleted content are deleted with it, so no dangling markers are left. A comment is
    /// removed along with its text when its start marker or its reference is deleted.
    /// </summary>
    /// <exception cref="NotSupportedException">The range is of a type that does not support deleting.</exception>
    public void Delete()
    {
        if (_isSpan)
        {
            DeleteSpan();
        }
        else if (_startOpenXmlElement is Paragraph)
        {
            DeleteParagraph((Paragraph)_startOpenXmlElement);
        }
        else if (_startOpenXmlElement is TableCell)
        {
            EmptyTableCell((TableCell)_startOpenXmlElement);
        }
        else if (_startOpenXmlElement is TableRow)
        {
            DeleteTableRow((TableRow)_startOpenXmlElement);
        }
        else if (_startOpenXmlElement is Table)
        {
            DeleteElements(new List<OpenXmlElement> { _startOpenXmlElement });
        }
        else if (_startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart)
        {
            DeleteElements(_startOpenXmlElement.GetElements<Run>(_endOpenXmlElement).Cast<OpenXmlElement>().ToList());
        }
        else if (_startOpenXmlElement is Document || _startOpenXmlElement is Body)
        {
            EmptyBody(GetBody());
        }
        else if (_startOpenXmlElement is Header || _startOpenXmlElement is Footer)
        {
            EmptyHeaderFooter(_startOpenXmlElement);
        }
        else
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Deleting is not currently supported for a range starting with the type {0}", _startOpenXmlElement.GetType().Name));
        }
    }

    private static void DeleteParagraph(Paragraph paragraph)
    {
        OpenXmlElement parent = paragraph.Parent;
        bool mustStay;
        if (parent is Body)
        {
            //A table must be followed by a paragraph, so the last paragraph of the body cannot go if it is the only block or follows a table.
            bool isLastBlock = paragraph.NextSibling() == null || paragraph.NextSibling() is SectionProperties;
            OpenXmlElement previous = paragraph.PreviousSibling();
            mustStay = isLastBlock && (previous == null || previous is Table);
        }
        else
        {
            //The only paragraph of a table cell, comment or similar container must stay.
            mustStay = parent.Elements<Paragraph>().Count() == 1;
        }

        if (mustStay)
        {
            DeleteElements(paragraph.ChildElements.Where(e => !(e is ParagraphProperties)).ToList());
        }
        else
        {
            DeleteElements(new List<OpenXmlElement> { paragraph });
        }
    }

    private static void EmptyTableCell(TableCell tableCell)
    {
        //Keep the formatting of the first paragraph on the single empty paragraph that a table cell must end with.
        Paragraph first = tableCell.Elements<Paragraph>().FirstOrDefault();
        ParagraphProperties paragraphProperties = first?.ParagraphProperties == null ? null : (ParagraphProperties)first.ParagraphProperties.CloneNode(true);

        DeleteElements(tableCell.ChildElements.Where(e => !(e is TableCellProperties)).ToList());

        Paragraph paragraph = new Paragraph();
        if (paragraphProperties != null)
        {
            paragraph.AppendChild(paragraphProperties);
        }
        tableCell.AppendChild(paragraph);
    }

    private static void DeleteTableRow(TableRow tableRow)
    {
        Table table = tableRow.Parent as Table;
        if (table != null && table.Elements<TableRow>().Count() == 1)
        {
            //A table cannot have no rows so deleting the last row deletes the table.
            DeleteElements(new List<OpenXmlElement> { table });
        }
        else
        {
            DeleteElements(new List<OpenXmlElement> { tableRow });
        }
    }

    private static void EmptyHeaderFooter(OpenXmlElement headerOrFooter)
    {
        //A header or footer keeps one empty paragraph, as Word does.
        DeleteElements(headerOrFooter.ChildElements.ToList());
        headerOrFooter.AppendChild(new Paragraph());
    }

    private static void EmptyBody(Body body)
    {
        DeleteElements(body.ChildElements.Where(e => !(e is SectionProperties)).ToList());

        Paragraph paragraph = new Paragraph();
        SectionProperties sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
        if (sectionProperties != null)
        {
            sectionProperties.InsertBeforeSelf(paragraph);
        }
        else
        {
            body.AppendChild(paragraph);
        }
    }

    /// <summary>
    /// Removes the elements from the tree together with the other half of any bookmark or comment that is only partly inside them.
    /// </summary>
    private static void DeleteElements(IList<OpenXmlElement> elements)
    {
        if (elements.Count == 0)
        {
            return;
        }

        //Find where the elements are and what is anchored in them before they are removed.
        OpenXmlElement root = elements[0].GetRoot();
        List<OpenXmlElement> content = elements.SelectMany(e => e.GetElements()).ToList();
        HashSet<string> bookmarkStartIds = new HashSet<string>(content.OfType<BookmarkStart>().Select(e => e.Id.Value));
        HashSet<string> bookmarkEndIds = new HashSet<string>(content.OfType<BookmarkEnd>().Select(e => e.Id.Value));
        HashSet<string> commentIds = new HashSet<string>(content.OfType<CommentRangeStart>().Select(e => e.Id.Value));
        commentIds.UnionWith(content.OfType<CommentReference>().Select(e => e.Id.Value));

        foreach (OpenXmlElement element in elements)
        {
            element.Remove();
        }

        //A bookmark with only one marker left would be invalid so the remaining marker goes too.
        foreach (string id in bookmarkStartIds.Except(bookmarkEndIds))
        {
            RemoveAll(root.Descendants<BookmarkEnd>().Where(e => e.Id.Value == id));
        }
        foreach (string id in bookmarkEndIds.Except(bookmarkStartIds))
        {
            RemoveAll(root.Descendants<BookmarkStart>().Where(e => e.Id.Value == id));
        }

        foreach (string id in commentIds)
        {
            RemoveComment(root, id);
        }
    }

    private static void RemoveComment(OpenXmlElement root, string id)
    {
        RemoveAll(root.Descendants<CommentRangeStart>().Where(e => e.Id.Value == id));
        RemoveAll(root.Descendants<CommentRangeEnd>().Where(e => e.Id.Value == id));

        //The comment reference lives in a run of its own, which goes with it.
        foreach (CommentReference reference in root.Descendants<CommentReference>().Where(e => e.Id.Value == id).ToList())
        {
            OpenXmlElement run = reference.Parent;
            reference.Remove();
            if (run is Run && !run.Elements().Any(e => !(e is RunProperties)))
            {
                run.Remove();
            }
        }

        Comments comments = (root as Document)?.MainDocumentPart?.WordprocessingCommentsPart?.Comments;
        if (comments != null)
        {
            RemoveAll(comments.Elements<Comment>().Where(e => e.Id.Value == id));
        }
    }

    private static void RemoveAll(IEnumerable<OpenXmlElement> elements)
    {
        foreach (OpenXmlElement element in elements.ToList())
        {
            element.Remove();
        }
    }
}
