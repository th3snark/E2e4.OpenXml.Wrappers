using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Intermediate;

/// <summary>
/// Meeting minutes: headings, a details table, a decision that is bookmarked and referred to by a field, an action table with a repeating
/// header row, review comments, and a footer with page numbers.
/// </summary>
internal static class MinutesExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Minutes of the project board, 14 March 2026"))
        {
            DocumentBuilder builder = new DocumentBuilder(document);
            builder.Paragraph("Minutes of the project board", "Title");

            TableWrapper details = builder.Table(3, 2);
            string[,] facts =
            {
                { "Date", "14 March 2026, 10:00 to 11:15" },
                { "Place", "Meeting room 2 and online" },
                { "Present", "Ana Silva (chair), Ben Okafor, Cy Duran, Dee Moreau (minutes)" },
            };
            for (int row = 0; row < 3; row++)
            {
                details.Cell(row, 0).Range.Text = facts[row, 0];
                details.Cell(row, 0).Range.Font.Bold = true;
                details.Cell(row, 0).BackgroundColor = "D9E2F3";
                details.Cell(row, 1).Range.Text = facts[row, 1];
            }
            details.Columns[0].Width = 1500;
            details.Columns[1].Width = 7000;
            details.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");
            details.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");

            builder.Paragraph("1. Progress", "Heading 1");
            ParagraphWrapper progress = builder.Paragraph("The build is on schedule. The first release candidate is expected on 28 March and testing will start the same day.");

            builder.Paragraph("2. Decisions", "Heading 1");
            ParagraphWrapper decision = builder.Paragraph("The board decided to ship the release candidate to the pilot customers before the public launch.");

            //Bookmark the decision so that other places in the document can refer to it.
            document.Bookmarks.Add("decision_pilot", decision.Range);

            builder.Paragraph("3. Actions", "Heading 1");
            ParagraphWrapper reference = builder.Paragraph("Actions that follow from the decision are listed below. The decision itself: ");

            //REF is the field behind Insert, Cross-reference. Its result is the text of the bookmark. Word refreshes it when it updates fields.
            FieldWrapper cross = document.Fields.Add(reference.Range, FieldType.Ref, "decision_pilot \\h");
            cross.Result.Text = decision.Range.Text.TrimEnd('\r');

            string[][] actions =
            {
                new[] { "Action", "Owner", "Due" },
                new[] { "Prepare the pilot customer list", "Ben", "21 March" },
                new[] { "Write the release notes", "Cy", "26 March" },
                new[] { "Book the pilot kick-off call", "Dee", "28 March" },
            };
            TableWrapper table = builder.Table(actions.Length, 3);
            for (int row = 0; row < actions.Length; row++)
            {
                for (int column = 0; column < 3; column++)
                {
                    table.Cell(row, column).Range.Text = actions[row][column];
                }
            }
            table.Rows[0].Range.Font.Bold = true;
            table.Rows[0].HeadingFormat = true;
            for (int column = 0; column < 3; column++)
            {
                table.Cell(0, column).BackgroundColor = "D9E2F3";
            }
            table.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");
            table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");

            builder.Paragraph("Next meeting: 11 April 2026, 10:00.").Format.SpaceBefore = 12;

            //Review comments, from two people. Comments can be added to a paragraph or a bookmark.
            CommentWrapper question = document.Comments.Add(progress.Range, "Is 28 March still realistic after the security review?");
            question.Author = "Ana Silva";
            question.Initial = "AS";
            CommentWrapper note = document.Comments.Add(document.Bookmarks["decision_pilot"].Range, "Please add the names of the pilot customers.");
            note.Author = "Ben Okafor";
            note.Initial = "BO";

            //A footer with the page number.
            RangeWrapper footer = document.Sections[0].Footers[HeaderFooterIndex.Primary].Range;
            footer.Text = "Project board minutes - page ";
            footer.Style = "Footer";
            document.Fields.Add(footer, FieldType.Page).Result.Text = "1";
        }
    }
}
