using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Intermediate;

/// <summary>
/// A letter made from a template, the way a mail merge works: the template has placeholders and bookmarks, and a program fills them in
/// from data. Shows creating a template, opening it again, find and replace, bookmarks and document properties together.
/// </summary>
internal static class LetterExample
{
    public static void Create(string path)
    {
        CreateTemplate(path);

        //Open the template as editable and merge the data into it.
        using (WordprocessingDocumentWrapper letter = WordprocessingDocumentWrapper.Open(path, true))
        {
            Dictionary<string, string> data = new Dictionary<string, string>
            {
                ["{{name}}"] = "Ms Rivera",
                ["{{product}}"] = "the Atlas planner",
                ["{{date}}"] = "14 March 2026",
                ["{{discount}}"] = "15%",
            };
            foreach (KeyValuePair<string, string> pair in data)
            {
                //ReplaceMode.All changes every occurrence in the range, including those in tables.
                letter.Content.Find.Execute(pair.Key, matchCase: true, replaceWith: pair.Value, replace: ReplaceMode.All);
            }

            //A bookmark marks the place for text that is more than one word. "\n" in the text becomes a line break.
            letter.Bookmarks["recipient_address"].Range.Text = "Ms Elena Rivera\n12 Harbour Road\nPortsmouth PO1 2AB";

            //The table cell that holds the account number is also bookmarked.
            letter.Bookmarks["account_number"].Range.Text = "ACC-004917";

            letter.BuiltInDocumentProperties.Title = "Letter to Ms Rivera";
            letter.BuiltInDocumentProperties.Subject = "Welcome letter";

            //Placeholders that were not in the data would be a mistake in the letter, so check that none are left.
            if (letter.Content.Find.Execute("{{"))
            {
                throw new InvalidOperationException("The letter still has placeholders.");
            }
        }
    }

    private static void CreateTemplate(string path)
    {
        using (WordprocessingDocumentWrapper template = SampleDocument.Create(path, "Letter template"))
        {
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Paragraph("Atlas Stationery Ltd", "Title");
            builder.Paragraph("1 Paper Lane, Leeds LS1 4AB").Format.SpaceAfter = 24;

            ParagraphWrapper address = builder.Paragraph("{{recipient address}}");
            address.Format.SpaceAfter = 18;
            template.Bookmarks.Add("recipient_address", address.Range);

            builder.Paragraph("{{date}}").Format.SpaceAfter = 18;
            builder.Paragraph("Dear {{name}},");
            builder.Paragraph("Thank you for choosing {{product}}. As a new customer you are entitled to a {{discount}} discount on your next order, which you can claim by quoting the account number below.");

            TableWrapper details = builder.Table(1, 2);
            details.Cell(0, 0).Range.Text = "Your account number";
            details.Cell(0, 1).Range.Text = "{{account}}";
            details.Cell(0, 0).BackgroundColor = "D9E2F3";
            details.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");
            details.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "808080");
            template.Bookmarks.Add("account_number", details.Cell(0, 1).Range);

            builder.Paragraph("We hope you enjoy using it.").Format.SpaceBefore = 12;
            builder.Paragraph("Yours sincerely,").Format.SpaceAfter = 36;
            builder.Paragraph("Sam Okafor");
            builder.Paragraph("Customer care, Atlas Stationery Ltd");
        }
    }
}
