using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Fields: adding them, reading the code and the result, and deleting them.
/// </summary>
/// <remarks>
/// <c>Fields.Add</c> works out the result at once for the fields that do not depend on the layout of the pages (dates, document properties,
/// and so on). The others (<c>PAGE</c>, <c>NUMPAGES</c>) are marked so that Word updates them when it opens the document, and the example
/// gives them a stand-in result to show until it does.
/// </remarks>
internal static class FieldExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            document.Paragraphs[0].Range.Text = "Printed on ";
            RangeWrapper line = document.Paragraphs[0].Range;

            //Add puts the field at the end of the range. The type names the code and the text is the rest of it.
            //A date does not depend on the layout, so the result is there already.
            document.Fields.Add(line, FieldType.Date, "\\@ \"d MMMM yyyy\"");

            //A second paragraph with two fields in it, and text between them.
            ParagraphWrapper pages = document.Paragraphs.Add();
            pages.Range.Text = "Page ";
            FieldWrapper page = document.Fields.Add(pages.Range, FieldType.Page);
            page.Result.Text = "1";
            pages.Range.InsertAfter(" of ");
            FieldWrapper count = document.Fields.Add(pages.Range, FieldType.NumPages);
            count.Result.Text = "1";

            //Any other field is added from its code. TITLE shows the Title property of the document, so set that first: a TITLE field is
            //empty when Word updates it (F9) in a document without a title.
            document.BuiltInDocumentProperties.Title = "Fields";
            ParagraphWrapper custom = document.Paragraphs.Add();
            custom.Range.Text = "Title of this document: ";
            document.Fields.Add(custom.Range, "TITLE");

            //Fields can be read back, in document order, with the kind of field worked out from the code.
            ParagraphWrapper summary = document.Paragraphs.Add();
            summary.Range.Text = "Fields in this document (listed here as plain text, not as fields): " + string.Join("; ", document.Fields.Select(f => f.Type + " (" + f.Code.Text + ")"));

            //A field that is not wanted is deleted with its result.
            FieldWrapper unwanted = document.Fields.Add(document.Paragraphs.Add().Range, FieldType.Author);
            unwanted.Delete();
        }
    }
}
