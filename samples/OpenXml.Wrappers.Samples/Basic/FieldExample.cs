using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Fields: adding them, reading the code and the result, and deleting them.
/// </summary>
/// <remarks>
/// The library never calculates a field. It writes the code and a result that you give it; a field added with <c>Fields.Add</c> is also
/// marked so that Word updates it when it opens the document.
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
            FieldWrapper date = document.Fields.Add(line, FieldType.Date, "\\@ \"d MMMM yyyy\"");
            date.Result = DateTime.Today.ToString("d MMMM yyyy", System.Globalization.CultureInfo.InvariantCulture);

            //A second paragraph with two fields in it, and text between them.
            ParagraphWrapper pages = document.Paragraphs.Add();
            pages.Range.Text = "Page ";
            FieldWrapper page = document.Fields.Add(pages.Range, FieldType.Page);
            page.Result = "1";
            pages.Range.InsertAfter(" of ");
            FieldWrapper count = document.Fields.Add(pages.Range, FieldType.NumPages);
            count.Result = "1";

            //Any other field is added from its code.
            ParagraphWrapper custom = document.Paragraphs.Add();
            custom.Range.Text = "Title of this document: ";
            FieldWrapper title = document.Fields.Add(custom.Range, "TITLE");
            title.Result = "Fields";

            //Fields can be read back, in document order, with the kind of field worked out from the code.
            ParagraphWrapper summary = document.Paragraphs.Add();
            summary.Range.Text = "Fields in this document: " + string.Join("; ", document.Fields.Select(f => f.Type + " (" + f.Code + ")"));

            //A field that is not wanted is deleted with its result.
            FieldWrapper unwanted = document.Fields.Add(document.Paragraphs.Add().Range, FieldType.Author);
            unwanted.Delete();
        }
    }
}
