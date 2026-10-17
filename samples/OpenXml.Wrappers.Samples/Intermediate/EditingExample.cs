using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Intermediate;

/// <summary>
/// Editing a document that already exists: open it, find and replace, fix a table, comment on what needs a person's attention, remove
/// what is out of date, and save. The "existing" document is made first so that the example is self-contained.
/// </summary>
internal static class EditingExample
{
    public static void Create(string path)
    {
        CreateOldVersion(path);

        //Open takes a path and, as the second argument, whether to allow changes. Changes are saved when the document is disposed.
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(path, true))
        {
            //1. Rename the product everywhere, in the text and in the table.
            document.Content.Find.Execute("Quillpad", matchCase: true, replaceWith: "Quillpad Pro", replace: ReplaceMode.All);

            //2. Remove the paragraphs that were written as drafts. ToList takes a copy first, so deleting while looping is safe.
            foreach (ParagraphWrapper paragraph in document.Paragraphs.ToList())
            {
                if (paragraph.Text.StartsWith("DRAFT:", StringComparison.Ordinal))
                {
                    paragraph.Delete();
                }
            }

            //3. Put a comment on every paragraph that still has a TODO in it.
            int todos = 0;
            foreach (ParagraphWrapper paragraph in document.Paragraphs.ToList())
            {
                if (paragraph.Text.Contains("TODO", StringComparison.Ordinal))
                {
                    CommentWrapper comment = document.Comments.Add(paragraph.Range, "Someone needs to finish this before the document goes out.");
                    comment.Author = "Editing example";
                    comment.Initials = "EE";
                    todos++;
                }
            }

            //4. Update the price table: change a price and add a row at the end.
            TableWrapper prices = document.Tables[0];
            prices.Cell(1, 1).Range.Text = "49.00";
            TableRowWrapper added = prices.Rows.Add();
            added.Cells[0].Range.Text = "Team (10 users)";
            added.Cells[1].Range.Text = "399.00";
            added.Range.Font.Bold = true;

            //5. Add text inside a bookmark that the old version left empty (the label before it stays), and note what was done at the end.
            document.Bookmarks["release_date"].Range.InsertAfter("14 March 2026");
            document.Paragraphs.Add().Range.Text = string.Format(CultureInfo.InvariantCulture, "Edited: {0} comment(s) added, {1} paragraph(s) now.", todos, document.Paragraphs.Count);

            document.BuiltInDocumentProperties.LastAuthor = "Editing example";
            document.BuiltInDocumentProperties.Title = "Quillpad Pro user guide";
        }
    }

    private static void CreateOldVersion(string path)
    {
        using (WordprocessingDocumentWrapper old = SampleDocument.Create(path, "Quillpad user guide"))
        {
            DocumentBuilder builder = new DocumentBuilder(old);
            builder.Paragraph("Quillpad user guide", "Title");
            ParagraphWrapper release = builder.Paragraph("Released on: ");
            old.Bookmarks.Add("release_date", release.Range);
            builder.Paragraph("DRAFT: this section will be rewritten by marketing.");
            builder.Paragraph("Getting started", "Heading 1");
            builder.Paragraph("Install Quillpad, sign in, and create your first notebook. TODO: add screenshots.");
            builder.Paragraph("Sync", "Heading 1");
            builder.Paragraph("Quillpad keeps your notebooks in step on all your devices. TODO: explain the offline mode.");
            builder.Paragraph("Prices", "Heading 1");

            TableWrapper table = builder.Table(2, 2);
            table.Cell(0, 0).Range.Text = "Plan";
            table.Cell(0, 1).Range.Text = "Price per year (GBP)";
            table.Cell(1, 0).Range.Text = "Personal";
            table.Cell(1, 1).Range.Text = "39.00";
            table.Rows[0].Range.Font.Bold = true;
            table.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");
            table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");
        }
    }
}
