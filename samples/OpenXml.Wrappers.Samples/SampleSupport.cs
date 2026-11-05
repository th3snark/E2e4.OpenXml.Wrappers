using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples;

/// <summary>
/// Creates the documents the examples start from.
/// </summary>
/// <remarks>
/// A document made by <see cref="WordprocessingDocumentWrapper.Create(string, bool, bool)"/> has Word's own styles, as a new document in
/// Word 365 does. <see cref="Create"/> formats some of them (Title, Heading 1 and 2, Header and Footer) with the library's own
/// <c>Style.Font</c> and <c>Style.ParagraphFormat</c>.
/// </remarks>
internal static class SampleDocument
{
    /// <summary>
    /// Creates a document with the standard styles and a title, and the builtin properties filled in.
    /// </summary>
    public static WordprocessingDocumentWrapper Create(string path, string title)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, true);
        AddStyles(document);

        document.BuiltInDocumentProperties.Title = title;
        document.BuiltInDocumentProperties.Author = "E2e4.OpenXml.Wrappers samples";
        document.BuiltInDocumentProperties.CreationDate = DateTime.UtcNow;
        return document;
    }

    private static void AddStyles(WordprocessingDocumentWrapper document)
    {
        //The styles are formatted through the library, the way a macro would set Style.Font in Word. Header and Footer are added on first use.
        StyleWrapperCollection styles = document.Styles;
        styles["Title"].Font.Size = 26;
        styles["Title"].Font.Color = "1F3864";
        styles["Title"].ParagraphFormat.SpaceAfter = 12;

        styles["Heading 1"].Font.Size = 16;
        styles["Heading 1"].Font.Bold = true;
        styles["Heading 1"].Font.Color = "2F5496";
        styles["Heading 1"].ParagraphFormat.SpaceBefore = 18;
        styles["Heading 1"].ParagraphFormat.SpaceAfter = 6;

        styles["Heading 2"].Font.Size = 13;
        styles["Heading 2"].Font.Bold = true;
        styles["Heading 2"].Font.Color = "2F5496";
        styles["Heading 2"].ParagraphFormat.SpaceBefore = 12;
        styles["Heading 2"].ParagraphFormat.SpaceAfter = 4;

        styles["Header"].Font.Size = 9;
        styles["Header"].Font.Color = "595959";
        styles["Footer"].Font.Size = 9;
        styles["Footer"].Font.Color = "595959";
    }
}

/// <summary>
/// Adds paragraphs and tables to the end of a document, which is the most common thing to do with a new one.
/// </summary>
/// <remarks>
/// The library has no <c>Document.Content.InsertParagraphAfter</c> that remembers where it is, so this keeps one spare empty paragraph at
/// the end: the first one in a new document, and the one after a table (a table must be followed by a paragraph). The next call fills the
/// spare paragraph instead of adding another one, so no stray empty paragraphs are left.
/// </remarks>
internal sealed class DocumentBuilder
{
    private readonly WordprocessingDocumentWrapper _document;
    private ParagraphWrapper _spare;

    public DocumentBuilder(WordprocessingDocumentWrapper document)
    {
        _document = document;
        _spare = document.Paragraphs[document.Paragraphs.Count - 1];
    }

    public WordprocessingDocumentWrapper Document
    {
        get { return _document; }
    }

    /// <summary>
    /// Adds a paragraph of text, in a style if one is given, and returns it.
    /// </summary>
    public ParagraphWrapper Paragraph(string text, string style = null)
    {
        ParagraphWrapper paragraph = TakeParagraph();
        paragraph.Range.Text = text;
        if (style != null)
        {
            paragraph.Range.Style = style;
        }
        return paragraph;
    }

    /// <summary>
    /// Adds an empty table of the given size at the end of the document and returns it. Borders are single lines.
    /// </summary>
    public TableWrapper Table(int rows, int columns)
    {
        ParagraphWrapper after = TakeParagraph();
        TableWrapper table = _document.Tables.Add(after.Range, rows, columns, DefaultTableBehavior.Word9TableBehavior);

        //The table is inserted before the paragraph, which stays after it as the spare one.
        _spare = after;
        return table;
    }

    private ParagraphWrapper TakeParagraph()
    {
        ParagraphWrapper paragraph = _spare ?? _document.Content.InsertParagraphAfter();
        _spare = null;
        return paragraph;
    }
}

/// <summary>
/// The look that the larger examples give to their tables.
/// </summary>
internal static class TableLook
{
    /// <summary>
    /// Gives a table single-line borders and, if asked, a dark header row that repeats on each page.
    /// </summary>
    public static void Grid(TableWrapper table, bool headerRow)
    {
        table.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");
        table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");
        if (headerRow)
        {
            TableRowWrapper header = table.Rows[0];
            header.HeadingFormat = true;
            header.Range.Font.Bold = true;
            header.Range.Font.Color = "FFFFFF";
            foreach (TableCellWrapper cell in header.Cells)
            {
                cell.Shading.BackgroundPatternColor = "1F3864";
            }
        }
    }

    /// <summary>
    /// Right-aligns the text of the given columns in all the rows of a table.
    /// </summary>
    public static void AlignRight(TableWrapper table, params int[] columns)
    {
        foreach (int column in columns)
        {
            foreach (TableRowWrapper row in table.Rows)
            {
                row.Cells[column].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            }
        }
    }
}
