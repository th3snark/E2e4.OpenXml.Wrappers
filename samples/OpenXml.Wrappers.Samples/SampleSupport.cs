using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples;

/// <summary>
/// Creates the documents the examples start from.
/// </summary>
/// <remarks>
/// A document made by <see cref="WordprocessingDocumentWrapper.Create(string, bool, bool)"/> has a body and nothing else, so it has no styles
/// and Word falls back on its own defaults. <see cref="Create"/> gives it the small set of styles that an ordinary Word document has (Normal,
/// Title, Heading 1 and 2, Header and Footer) and sets the formatting of them with the library's own <c>Style.Font</c> and
/// <c>Style.ParagraphFormat</c>.
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
        StyleDefinitionsPart part = document.WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        part.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "Calibri", ComplexScript = "Calibri" }, new FontSize() { Val = "22" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(new SpacingBetweenLines() { After = "120", Line = "264", LineRule = LineSpacingRuleValues.Auto }))),
            new Style(new StyleName() { Val = "Normal" }, new PrimaryStyle()) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
            CreateParagraphStyle("Title", "Title", new StyleParagraphProperties(), 10),
            CreateParagraphStyle("Heading1", "heading 1", new StyleParagraphProperties(new KeepNext(), new KeepLines(), new OutlineLevel() { Val = 0 }), 9),
            CreateParagraphStyle("Heading2", "heading 2", new StyleParagraphProperties(new KeepNext(), new KeepLines(), new OutlineLevel() { Val = 1 }), 9),
            CreateParagraphStyle("Header", "header", new StyleParagraphProperties(), 99),
            CreateParagraphStyle("Footer", "footer", new StyleParagraphProperties(), 99));

        //From here the styles are formatted through the library, the way a macro would set Style.Font in Word.
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

    private static Style CreateParagraphStyle(string id, string name, StyleParagraphProperties properties, int priority)
    {
        Style style = new Style(
            new StyleName() { Val = name },
            new BasedOn() { Val = "Normal" },
            new NextParagraphStyle() { Val = "Normal" },
            new UIPriority() { Val = priority },
            new PrimaryStyle())
        {
            Type = StyleValues.Paragraph,
            StyleId = id
        };
        if (properties.HasChildren)
        {
            style.AppendChild(properties);
        }
        return style;
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
        TableWrapper table = _document.Tables.Add(after.Range, rows, columns);

        //The table is inserted before the paragraph, which stays after it as the spare one.
        _spare = after;
        return table;
    }

    private ParagraphWrapper TakeParagraph()
    {
        ParagraphWrapper paragraph = _spare ?? _document.Range.InsertParagraphAfter();
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
                cell.BackgroundColor = "1F3864";
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
