using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The information about a range that <see cref="RangeWrapper.Information"/> gives, as <c>WdInformation</c> in Word. Only the information that
/// does not depend on how the pages are laid out is here.
/// </summary>
public enum InformationType
{
    /// <summary>True if the whole range is in a table.</summary>
    WithInTable = 12,

    /// <summary>The number of the row (counting from 1) the start of the range is in; -1 if it is not in a table.</summary>
    StartOfRangeRowNumber = 13,

    /// <summary>The number of the row the end of the range is in; -1 if it is not in a table.</summary>
    EndOfRangeRowNumber = 14,

    /// <summary>The number of rows in the table the range is in; -1 if it is not in a table.</summary>
    MaximumNumberOfRows = 15,

    /// <summary>The number of the column (a cell of the row, counting from 1) the start of the range is in; -1 if it is not in a table. The end-of-row mark is one past the last cell.</summary>
    StartOfRangeColumnNumber = 16,

    /// <summary>The number of the column the end of the range is in; -1 if it is not in a table.</summary>
    EndOfRangeColumnNumber = 17,

    /// <summary>The number of cells in the widest row of the table the range is in; -1 if it is not in a table.</summary>
    MaximumNumberOfColumns = 18,

    /// <summary>True if the range is in a comment.</summary>
    InCommentPane = 26,

    /// <summary>True if the range is in a header or footer.</summary>
    InHeaderFooter = 28,

    /// <summary>The kind of header or footer the range is in: 1 for primary, 2 for the first page, 3 for even pages; -1 if it is in neither.</summary>
    HeaderFooterType = 33,

    /// <summary>True if the range is in a footnote.</summary>
    InFootnote = 42,

    /// <summary>True if the range is in an endnote.</summary>
    InEndnote = 43
}

/// <summary>
/// Information about where a range is: tables, headers and footers, bookmarks and sections.
/// </summary>
public partial class RangeWrapper
{
    /// <summary>
    /// Information about the range, as <c>Range.Information</c> in Word, for the types that do not depend on the layout of the pages. The value is a
    /// <see cref="bool"/> or an <see cref="int"/>, as the type says.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The type is not one of those that are supported.</exception>
    public object Information(InformationType type)
    {
        if (!EnumCompat.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }

        OpenXmlElement story = StoryMap.FindRoot(AnchorElement);
        if (type == InformationType.InCommentPane)
        {
            return story is Comment || story.Parent is Comments;
        }
        if (type == InformationType.InFootnote)
        {
            return story is Footnote;
        }
        if (type == InformationType.InEndnote)
        {
            return story is Endnote;
        }
        if (type == InformationType.InHeaderFooter)
        {
            return story is Header || story is Footer;
        }
        if (type == InformationType.HeaderFooterType)
        {
            return HeaderFooterTypeOf(story);
        }

        GetPositions(out StoryMap map, out int start, out int end);
        TableRow startRow = RowAt(map, start, out int startColumn, false);
        int last = end > start ? end - 1 : end;
        TableRow endRow = RowAt(map, last, out int endColumn, end > start && last > start);
        Table table = startRow?.Parent as Table;

        //The whole range is in a table when the end is in the table of the start, or in a table nested in it.
        bool within = table != null && endRow != null && (endRow.Parent == table || endRow.Ancestors<Table>().Contains(table));
        if (type == InformationType.WithInTable)
        {
            return within;
        }
        if (!within)
        {
            return -1;
        }

        switch (type)
        {
            case InformationType.StartOfRangeRowNumber:
                return IndexOfRow(startRow);
            case InformationType.EndOfRangeRowNumber:
                return IndexOfRow(endRow.Parent == table ? endRow : endRow.Ancestors<TableRow>().First(r => r.Parent == table));
            case InformationType.MaximumNumberOfRows:
                return table.Elements<TableRow>().Count();
            case InformationType.StartOfRangeColumnNumber:
                return startColumn;
            case InformationType.EndOfRangeColumnNumber:
                return endColumn;
            default:
                return table.Elements<TableRow>().Max(r => r.Elements<TableCell>().Count());
        }
    }

    //The row a position is in and its column, counting the end-of-row mark as one past the last cell (or, when asked, as the last cell, which is
    //how Word counts the end of a range that includes the mark). Null when the position is not in a table.
    private static TableRow RowAt(StoryMap map, int position, out int column, bool markIsLastCell)
    {
        column = -1;
        int index = map.IndexAt(position);
        if (index < 0)
        {
            return null;
        }

        Atom atom = map.Atoms[index];
        if (atom.Kind == AtomKind.RowMark)
        {
            TableRow mark = (TableRow)atom.Element;
            int cells = mark.Elements<TableCell>().Count();
            column = markIsLastCell ? cells : cells + 1;
            return mark;
        }

        TableCell cell = atom.Element.Ancestors<TableCell>().FirstOrDefault();
        if (cell == null)
        {
            return null;
        }
        TableRow row = (TableRow)cell.Parent;
        column = row.Elements<TableCell>().ToList().IndexOf(cell) + 1;
        return row;
    }

    private static int IndexOfRow(TableRow row)
    {
        return row.Parent.Elements<TableRow>().ToList().IndexOf(row) + 1;
    }

    private static int HeaderFooterTypeOf(OpenXmlElement story)
    {
        if (!(story is Header) && !(story is Footer))
        {
            return -1;
        }

        OpenXmlPart part = (story as Header)?.HeaderPart ?? (OpenXmlPart)(story as Footer)?.FooterPart;
        MainDocumentPart main = part?.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
        string id = main?.GetIdOfPart(part);
        HeaderFooterReferenceType reference = main?.Document?.Descendants<HeaderFooterReferenceType>().FirstOrDefault(r => r.Id?.Value == id);
        if (reference?.Type == null || !reference.Type.HasValue)
        {
            return 1;
        }
        return reference.Type.Value == HeaderFooterValues.First ? 2 : reference.Type.Value == HeaderFooterValues.Even ? 3 : 1;
    }

    /// <summary>
    /// The number of the bookmark the start of the range is in (or at the end of), as <c>Range.BookmarkID</c> in Word: bookmarks are numbered
    /// from 1 in the order they start in the document, and when bookmarks are nested it is the innermost, the one that starts last. 0 if the start
    /// is not in a bookmark.
    /// </summary>
    public int BookmarkID
    {
        get
        {
            int start = Start;
            List<BookmarkWrapper> bookmarks = BookmarksByStart();
            for (int i = bookmarks.Count - 1; i >= 0; i--)
            {
                if (bookmarks[i].Start <= start && start <= bookmarks[i].End)
                {
                    return i + 1;
                }
            }
            return 0;
        }
    }

    /// <summary>
    /// The number of the last bookmark that starts at or before the start of the range, as <c>Range.PreviousBookmarkID</c> in Word, numbered as
    /// for <see cref="BookmarkID"/>. 0 if there is none.
    /// </summary>
    public int PreviousBookmarkID
    {
        get
        {
            int start = Start;
            return BookmarksByStart().Count(b => b.Start <= start);
        }
    }

    private List<BookmarkWrapper> BookmarksByStart()
    {
        RangeWrapper story = new RangeWrapper(StoryMap.FindRoot(AnchorElement));
        return story.Bookmarks.Select((b, order) => (Bookmark: b, Start: b.Start, Order: order)).OrderBy(b => b.Start).ThenBy(b => b.Order).Select(b => b.Bookmark).ToList();
    }

    /// <summary>
    /// The sections that the range is in or touches, as <c>Range.Sections</c> in Word. Outside the main text it is one section, as measured
    /// in Word: for a footnote or endnote the section of its mark, for a comment that of the text it is on, for the text of a text box that
    /// of its anchor, and for a header or footer the first section that has it (not one that is linked to the previous section).
    /// </summary>
    /// <exception cref="NotSupportedException">The range is not in a document's main text or one of its other stories.</exception>
    public SectionWrapperCollection Sections
    {
        get
        {
            OpenXmlElement story = StoryMap.FindRoot(AnchorElement);
            if (!(story is Body body))
            {
                OpenXmlElement at = SectionAnchor(AnchorElement, 0);
                Body main = at == null ? null : StoryMap.FindRoot(at) as Body;
                if (main == null)
                {
                    throw new NotSupportedException("The range is not in the main text or another story of a document.");
                }
                int index = at is SectionProperties properties ? SectionOfProperties(main, properties) : SectionWrapperCollection.IndexOf(main, at);
                return new SectionWrapperCollection(main, index, index);
            }

            GetPositions(out StoryMap map, out int start, out int end);
            List<int> ends = map.Atoms
                .Where(a => a.Kind == AtomKind.ParagraphMark && ((Paragraph)a.Element).ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
                .Select(a => a.End)
                .ToList();
            int first = ends.Count(e => e <= start);
            int last = ends.Count(e => e <= (end > start ? end - 1 : start));
            return new SectionWrapperCollection(body, first, last);
        }
    }

    //The element of the main text whose section a story outside it has (the section properties, for a header or footer), or null.
    private static OpenXmlElement SectionAnchor(OpenXmlElement element, int depth)
    {
        OpenXmlElement story = StoryMap.FindRoot(element);
        if (story is Body)
        {
            return element;
        }
        if (story is TextBoxContent && depth <= 3)
        {
            //The text of a text box takes the place of the run it is anchored in.
            OpenXmlElement run = story.Ancestors<Run>().FirstOrDefault();
            return run == null ? null : SectionAnchor(run, depth + 1);
        }
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(element);
        Body body = main?.Document?.Body;
        if (body == null || depth > 3)
        {
            return null;
        }

        OpenXmlElement anchor = null;
        if (story is FootnoteEndnoteType note)
        {
            anchor = NoteOps.FindReference(main, note is Endnote, note);
        }
        else if (story is Comment comment)
        {
            string id = comment.Id?.Value;
            IEnumerable<OpenXmlElement> stories = new OpenXmlElement[] { body }
                .Concat(main.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<OpenXmlElement>())
                .Concat(main.EndnotesPart?.Endnotes?.Elements<Endnote>() ?? Enumerable.Empty<OpenXmlElement>());
            anchor = stories.SelectMany(s => s.Descendants())
                .FirstOrDefault(e => (e is CommentRangeStart start && start.Id?.Value == id) || (e is CommentReference reference && reference.Id?.Value == id));
        }
        else if (story is Header || story is Footer)
        {
            OpenXmlPart part = (story as OpenXmlPartRootElement)?.OpenXmlPart;
            string relationship = part == null ? null : main.GetIdOfPart(part);
            //The first section that refers to the part; the last section's properties are the body's own.
            return body.Descendants<SectionProperties>()
                .FirstOrDefault(s => s.Elements<HeaderFooterReferenceType>().Any(r => r.Id?.Value == relationship));
        }
        return anchor == null ? null : SectionAnchor(anchor, depth + 1);
    }

    private static int SectionOfProperties(Body body, SectionProperties properties)
    {
        Paragraph paragraph = properties.Parent?.Parent as Paragraph;
        if (paragraph == null)
        {
            return body.Descendants<Paragraph>().Count(p => p.ParagraphProperties?.SectionProperties != null);
        }
        return SectionWrapperCollection.IndexOf(body, paragraph);
    }
}
