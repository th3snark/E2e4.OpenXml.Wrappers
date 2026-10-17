using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Bookmarks name a place in the document, so that code can find it again to read or change it.
/// </summary>
internal static class BookmarkExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            document.Paragraphs[0].Range.Text = "Dear [customer name],";
            document.Paragraphs.Add().Range.Text = "Your order will arrive on [date].";
            ParagraphWrapper tableSpot = document.Range.InsertParagraphAfter();
            TableWrapper table = document.Tables.Add(tableSpot.Range, 1, 2);
            table.Cell(0, 0).Range.Text = "Order number";
            table.Cell(0, 1).Range.Text = "[order]";

            //Bookmark a whole paragraph and a whole table cell. Names follow Word's rules: a letter first, then letters, digits, underscore.
            document.Bookmarks.Add("customer", document.Paragraphs[0].Range);
            document.Bookmarks.Add("order_number", table.Cell(0, 1).Range);
            document.Bookmarks.Add("delivery", document.Paragraphs[1].Range);

            //Look a bookmark up by name (null if it does not exist) and replace what is inside it.
            if (document.Bookmarks.Exists("customer"))
            {
                document.Bookmarks["customer"].Range.Text = "Dear Ms Rivera,";
            }
            document.Bookmarks["order_number"].Range.Text = "A-10442";

            //Text can also be added inside the bookmark, and a bookmark that is no longer wanted is deleted.
            document.Bookmarks["delivery"].Range.Find.Execute("[date]", replaceWith: "Friday 14 March", replace: ReplaceMode.All);
            document.Bookmarks["delivery"].Range.InsertAfter(" Thank you for waiting.");
            document.Bookmarks["delivery"].Delete();

            //The paragraph after the table.
            tableSpot.Range.Text = string.Format(CultureInfo.InvariantCulture, "Bookmarks left in this document: {0}.", document.Bookmarks.Count);
        }
    }
}
