using OpenXml.Wrappers.Samples.Advanced;
using OpenXml.Wrappers.Samples.Basic;
using OpenXml.Wrappers.Samples.Intermediate;

namespace OpenXml.Wrappers.Samples;

/// <summary>
/// One example: a name, the file it writes and the method that writes it.
/// </summary>
/// <param name="Category">Basic, Intermediate or Advanced.</param>
/// <param name="Name">What the example shows.</param>
/// <param name="FileName">The name of the .docx file the example writes.</param>
/// <param name="Create">Creates the document at the path it is given.</param>
public sealed record Example(string Category, string Name, string FileName, Action<string> Create);

/// <summary>
/// The list of all the examples, in the order of the README.
/// </summary>
public static class Examples
{
    public static IReadOnlyList<Example> All { get; } = new List<Example>
    {
        new("Basic", "Range", "basic-01-range.docx", RangeExample.Create),
        new("Basic", "Paragraph", "basic-02-paragraph.docx", ParagraphExample.Create),
        new("Basic", "Table and column", "basic-03-table.docx", TableExample.Create),
        new("Basic", "Row", "basic-04-row.docx", RowExample.Create),
        new("Basic", "Cell", "basic-05-cell.docx", CellExample.Create),
        new("Basic", "Bookmark", "basic-06-bookmark.docx", BookmarkExample.Create),
        new("Basic", "Comment", "basic-07-comment.docx", CommentExample.Create),
        new("Basic", "Section and page setup", "basic-08-section.docx", SectionExample.Create),
        new("Basic", "Style", "basic-09-style.docx", StyleExample.Create),
        new("Basic", "Field", "basic-10-field.docx", FieldExample.Create),
        new("Basic", "Header and footer", "basic-11-header-footer.docx", HeaderFooterExample.Create),
        new("Basic", "Document properties", "basic-12-properties.docx", PropertiesExample.Create),
        new("Basic", "Character ranges", "basic-13-characters.docx", CharactersExample.Create),

        new("Intermediate", "Letter from a template", "intermediate-01-letter.docx", LetterExample.Create),
        new("Intermediate", "Invoice", "intermediate-02-invoice.docx", InvoiceExample.Create),
        new("Intermediate", "Meeting minutes", "intermediate-03-minutes.docx", MinutesExample.Create),
        new("Intermediate", "Report with page numbers", "intermediate-04-report.docx", ReportExample.Create),
        new("Intermediate", "Editing an existing document", "intermediate-05-editing.docx", EditingExample.Create),

        new("Advanced", "Quarterly business report", "advanced-01-quarterly-report.docx", QuarterlyReportExample.Create),
        new("Advanced", "Product catalogue", "advanced-02-catalogue.docx", CatalogueExample.Create),
        new("Advanced", "Services agreement", "advanced-03-agreement.docx", AgreementExample.Create),
    };
}
