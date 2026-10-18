using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Advanced;

/// <summary>
/// A services agreement made from a template with placeholders: numbered clauses (SEQ fields), cross references to bookmarked clauses
/// (REF fields), defined terms in a table, a fee schedule with merged cells in a section of its own, a signature block, comments from
/// reviewers, a header and footer with page numbers, and custom properties. The placeholders are filled in at the end with find and replace.
/// </summary>
internal static class AgreementExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Services agreement"))
        {
            DocumentBuilder builder = new DocumentBuilder(document);

            builder.Paragraph("Services agreement", "Title");
            builder.Paragraph("This agreement is made on {{date}} between {{supplier}} (the Supplier) and {{customer}} (the Customer).");

            //1. Definitions: a table of terms, with the terms in bold.
            AddClause(document, builder, 1, "Definitions", null);
            (string Term, string Meaning)[] terms =
            {
                ("Services", "the services described in the Schedule."),
                ("Fees", "the fees set out in the Schedule."),
                ("Term", "the period of {{term}} months starting on the date of this agreement."),
            };
            TableWrapper definitions = builder.Table(terms.Length, 2);
            for (int i = 0; i < terms.Length; i++)
            {
                definitions.Cell(i, 0).Range.Text = terms[i].Term;
                definitions.Cell(i, 0).Range.Font.Bold = true;
                definitions.Cell(i, 1).Range.Text = terms[i].Meaning;
            }
            definitions.Columns[0].Width = 1500;
            definitions.Columns[1].Width = 7000;
            TableLook.Grid(definitions, false);

            //2 to 5: the clauses. The heading of each is bookmarked so that other clauses can refer to it.
            ParagraphWrapper services = AddClause(document, builder, 2, "Services", "The Supplier will provide the Services to the Customer during the Term with reasonable skill and care.");
            ParagraphWrapper fees = AddClause(document, builder, 3, "Fees and payment", "The Customer will pay the Fees within 30 days of each invoice. Late payments carry interest at 4% a year above the base rate.");
            ParagraphWrapper confidentiality = AddClause(document, builder, 4, "Confidentiality", "Each party will keep the other's confidential information secret and use it only for this agreement, for the Term and for two years after it.");
            ParagraphWrapper termination = AddClause(document, builder, 5, "Termination", "Either party may end this agreement on 60 days' written notice. Payment obligations continue as set out in ");

            //A cross reference: a REF field to the bookmark of the heading of the fees clause.
            FieldWrapper reference = document.Fields.Add(termination.Range, FieldType.Ref, "clause_fees \\h");
            reference.Result.Text = "3. Fees and payment";
            termination.Range.InsertAfter(", up to the date on which this agreement ends.");

            //Comments from the reviewers.
            CommentWrapper legal = document.Comments.Add(fees.Range, "Legal: check the interest rate against the late payment rules.");
            legal.Author = "Legal team";
            legal.Initial = "LT";
            CommentWrapper sales = document.Comments.Add(services.Range, "Sales: the customer asked for a service level to be added here.");
            sales.Author = "Sales";
            sales.Initial = "SA";

            //Signature block: two columns, a line to sign on above each name.
            builder.Paragraph("Signed by the parties", "Heading 1");
            TableWrapper signatures = builder.Table(3, 2);
            signatures.Cell(0, 0).Range.Text = "For the Supplier";
            signatures.Cell(0, 1).Range.Text = "For the Customer";
            signatures.Rows[0].Range.Font.Bold = true;
            signatures.Rows[1].Height = 48;
            signatures.Rows[1].HeightRule = RowHeightRule.AtLeast;
            signatures.Cell(1, 0).Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.SingleLine;
            signatures.Cell(1, 1).Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.SingleLine;
            signatures.Cell(2, 0).Range.Text = "{{supplier signatory}}, {{supplier title}}";
            signatures.Cell(2, 1).Range.Text = "{{customer signatory}}, {{customer title}}";
            signatures.Borders.SetOutside(BorderLineStyle.None, 0.5, "auto");
            signatures.Borders.SetInside(BorderLineStyle.None, 0.5, "auto");
            document.Bookmarks.Add("signature_block", signatures.Range);

            //The schedule goes in a section of its own.
            ParagraphWrapper end = builder.Paragraph(string.Empty);
            end.InsertSectionBreak(SectionStart.NewPage);
            builder.Paragraph("Schedule: Services and Fees", "Heading 1");
            TableWrapper schedule = builder.Table(5, 4);
            schedule.Cell(0, 0).Range.Text = "Services and Fees";
            string[][] rows =
            {
                new[] { "Service", "Unit", "Quantity", "Fee (GBP)" },
                new[] { "Platform subscription", "month", "{{term}}", "{{fee platform}}" },
                new[] { "Onboarding and training", "day", "5", "{{fee training}}" },
                new[] { "Priority support", "month", "{{term}}", "{{fee support}}" },
            };
            for (int row = 0; row < rows.Length; row++)
            {
                for (int column = 0; column < 4; column++)
                {
                    schedule.Cell(row + 1, column).Range.Text = rows[row][column];
                }
            }

            //The first row is one cell across the table, made by merging the four cells.
            schedule.Cell(0, 0).Merge(schedule.Cell(0, 3));
            schedule.Cell(0, 0).Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            schedule.Cell(0, 0).BackgroundColor = "1F3864";
            schedule.Rows[0].Range.Font.Bold = true;
            schedule.Rows[0].Range.Font.Color = "FFFFFF";
            for (int column = 0; column < 4; column++)
            {
                schedule.Cell(1, column).BackgroundColor = "D9E2F3";
            }
            schedule.Rows[1].Range.Font.Bold = true;
            schedule.Borders.SetOutside(BorderLineStyle.SingleLine, 1, "1F3864");
            schedule.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");

            //The numbers are right-aligned. The first row is skipped: after the merge it has only one cell.
            for (int row = 1; row < schedule.Rows.Count; row++)
            {
                schedule.Rows[row].Cells[2].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                schedule.Rows[row].Cells[3].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            }

            //Header and footer, for the whole agreement.
            SectionWrapper first = document.Sections[0];
            RangeWrapper header = first.Headers[HeaderFooterIndex.Primary].Range;
            header.Text = "CONFIDENTIAL - Services agreement";
            header.Style = "Header";
            header.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            RangeWrapper footer = first.Footers[HeaderFooterIndex.Primary].Range;
            footer.Text = "Agreement {{reference}} - page ";
            footer.Style = "Footer";
            footer.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            document.Fields.Add(footer, FieldType.Page).Result.Text = "1";
            footer.InsertAfter(" of ");
            document.Fields.Add(footer, FieldType.NumPages).Result.Text = "3";

            //Fill in the placeholders: the same Find that was used to look for text replaces it in every part of the document. The
            //header and footer are separate ranges, so they are replaced too.
            Dictionary<string, string> values = new Dictionary<string, string>
            {
                ["{{date}}"] = "14 March 2026",
                ["{{supplier}}"] = "Atlas Software Ltd",
                ["{{customer}}"] = "Eastside Offices Ltd",
                ["{{term}}"] = "12",
                ["{{reference}}"] = "SA-2026-0217",
                ["{{supplier signatory}}"] = "Sam Okafor",
                ["{{supplier title}}"] = "Managing director",
                ["{{customer signatory}}"] = "Elena Rivera",
                ["{{customer title}}"] = "Finance director",
                ["{{fee platform}}"] = "1,200.00",
                ["{{fee training}}"] = "3,500.00",
                ["{{fee support}}"] = "300.00",
            };
            foreach (KeyValuePair<string, string> pair in values)
            {
                document.Content.Find.Execute(pair.Key, matchCase: true, replaceWith: pair.Value, replace: ReplaceMode.All);
                footer.Find.Execute(pair.Key, matchCase: true, replaceWith: pair.Value, replace: ReplaceMode.All);
            }

            document.BuiltInDocumentProperties.Subject = "Services agreement";
            document.BuiltInDocumentProperties.Company = "Atlas Software Ltd";
            document.CustomDocumentProperties.Add("Reference", DocumentPropertyType.Text, values["{{reference}}"]);
            document.CustomDocumentProperties.Add("Term (months)", DocumentPropertyType.Number, 12);
            document.CustomDocumentProperties.Add("Signed", DocumentPropertyType.Boolean, false);
        }
    }

    //A numbered clause: a heading line with a SEQ field for the number, then the text. The heading is bookmarked as clause_ plus the
    //first word of its title, in lower case.
    private static ParagraphWrapper AddClause(WordprocessingDocumentWrapper document, DocumentBuilder builder, int number, string title, string text)
    {
        ParagraphWrapper heading = builder.Paragraph(string.Empty, "Heading 2");
        FieldWrapper numberField = document.Fields.Add(heading.Range, FieldType.Sequence, "clause \\* ARABIC");
        numberField.Result.Text = number.ToString(CultureInfo.InvariantCulture);
        heading.Range.InsertAfter(". " + title);
        string name = "clause_" + title.Split(' ')[0].ToLowerInvariant();
        document.Bookmarks.Add(name, heading.Range);

        if (text == null)
        {
            return heading;
        }
        return builder.Paragraph(text);
    }
}
