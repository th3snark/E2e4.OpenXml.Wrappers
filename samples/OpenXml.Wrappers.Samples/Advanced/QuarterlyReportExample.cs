using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Advanced;

/// <summary>
/// A quarterly business report, made from data: a cover page, a table of contents, headings and styled call-outs, tables worked out from
/// the data (revenue, milestones, risks), a cross reference to a bookmarked figure, review comments, a landscape appendix, headers and
/// footers with page numbers, and document properties. It uses most of the types of the library.
/// </summary>
internal static class QuarterlyReportExample
{
    private sealed record Region(string Name, decimal Previous, decimal Current);

    private sealed record Milestone(string Name, string Owner, string Due, string Status);

    private sealed record Risk(string Description, string Severity, string Mitigation);

    public static void Create(string path)
    {
        Region[] regions =
        {
            new("North", 1_240_000m, 1_385_000m),
            new("South", 980_000m, 1_012_000m),
            new("East", 1_105_000m, 1_098_000m),
            new("West", 1_430_000m, 1_612_000m),
            new("International", 640_000m, 802_000m),
        };
        Milestone[] milestones =
        {
            new("Launch of the mobile app", "Ana Silva", "15 July", "Done"),
            new("New distribution centre open", "Ben Okafor", "30 August", "Done"),
            new("Customer portal, phase 2", "Cy Duran", "30 September", "Late"),
            new("Supplier consolidation", "Dee Moreau", "30 September", "On track"),
        };
        Risk[] risks =
        {
            new("Supplier concentration: two suppliers provide 60% of components", "High", "Qualify a third supplier by the end of Q4"),
            new("Delay of the customer portal pushes launch into the holiday season", "Medium", "Add two contractors; weekly steering review"),
            new("Currency movements reduce international margin", "Medium", "Hedge 50% of the forecast exposure"),
            new("Warehouse capacity at peak", "Low", "Short-term lease agreed as an option"),
        };

        decimal previousTotal = regions.Sum(r => r.Previous);
        decimal currentTotal = regions.Sum(r => r.Current);

        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Quarterly business report, Q3 2026"))
        {
            AddCustomStyles(document);
            DocumentBuilder builder = new DocumentBuilder(document);

            //Cover page.
            builder.Paragraph("Quarterly business report", "Title").Format.SpaceBefore = 150;
            builder.Paragraph("Third quarter 2026").Range.Font.Size = 18;
            builder.Paragraph("Prepared by the finance team for the board of directors").Format.SpaceAfter = 18;
            ParagraphWrapper confidential = builder.Paragraph("Confidential - not for distribution outside the company.");
            confidential.Range.Font.Italic = true;
            confidential.InsertBreak(BreakType.PageBreak);

            //Contents.
            builder.Paragraph("Contents", "Heading 1");
            ParagraphWrapper contents = builder.Paragraph(string.Empty);
            FieldWrapper toc = document.Fields.Add(contents.Range, FieldType.TableOfContents, "\\o \"1-2\" \\h \\z \\u");
            toc.Result.Text = "Update this field to build the table of contents.";
            contents.InsertBreak(BreakType.PageBreak);

            //1. Executive summary. The revenue figure is bookmarked so that other parts of the document can quote it.
            builder.Paragraph("1. Executive summary", "Heading 1");
            ParagraphWrapper figure = builder.Paragraph(string.Format(CultureInfo.InvariantCulture, "Revenue for the quarter was {0}, {1:+0.0;-0.0}% on the same quarter last year.", Money(currentTotal), Percent(previousTotal, currentTotal)));
            document.Bookmarks.Add("revenue_total", figure.Range);
            builder.Paragraph("Four of the five regions grew. The West region and the international business were the strongest, helped by the new distribution centre and the launch of the mobile app. The East region was slightly down.");
            builder.Paragraph("Key point: the customer portal is late, and is the main risk to the plan for the fourth quarter.", "Key Point");

            //2. Financial results.
            builder.Paragraph("2. Financial results", "Heading 1");
            builder.Paragraph("Revenue by region", "Heading 2");
            TableWrapper revenue = builder.Table(1 + regions.Length + 1, 4);
            string[] heading = { "Region", "Q3 2025", "Q3 2026", "Change" };
            for (int column = 0; column < heading.Length; column++)
            {
                revenue.Cell(0, column).Range.Text = heading[column];
            }
            for (int i = 0; i < regions.Length; i++)
            {
                revenue.Cell(i + 1, 0).Range.Text = regions[i].Name;
                revenue.Cell(i + 1, 1).Range.Text = Money(regions[i].Previous);
                revenue.Cell(i + 1, 2).Range.Text = Money(regions[i].Current);
                revenue.Cell(i + 1, 3).Range.Text = string.Format(CultureInfo.InvariantCulture, "{0:+0.0;-0.0}%", Percent(regions[i].Previous, regions[i].Current));
                if (regions[i].Current < regions[i].Previous)
                {
                    revenue.Cell(i + 1, 3).Range.Font.Color = "C00000";
                }
            }
            int totalRow = regions.Length + 1;
            revenue.Cell(totalRow, 0).Range.Text = "Total";
            revenue.Cell(totalRow, 1).Range.Text = Money(previousTotal);
            revenue.Cell(totalRow, 2).Range.Text = Money(currentTotal);
            revenue.Cell(totalRow, 3).Range.Text = string.Format(CultureInfo.InvariantCulture, "{0:+0.0;-0.0}%", Percent(previousTotal, currentTotal));
            TableLook.Grid(revenue, true);
            TableLook.AlignRight(revenue, 1, 2, 3);
            revenue.Rows[totalRow].Range.Font.Bold = true;
            foreach (TableCellWrapper cell in revenue.Rows[totalRow].Cells)
            {
                cell.Shading.BackgroundPatternColor = "D9E2F3";
            }
            revenue.Alignment = RowAlignment.Center;

            ParagraphWrapper refer = builder.Paragraph("As the summary says: ");
            refer.Format.SpaceBefore = 10;
            FieldWrapper quote = document.Fields.Add(refer.Range, FieldType.Ref, "revenue_total \\h");
            quote.Result.Text = figure.Range.Text.TrimEnd('\r');

            //3. Operations.
            builder.Paragraph("3. Operations", "Heading 1");
            builder.Paragraph("Milestones", "Heading 2");
            TableWrapper plan = builder.Table(1 + milestones.Length, 4);
            string[] planHeading = { "Milestone", "Owner", "Due", "Status" };
            for (int column = 0; column < planHeading.Length; column++)
            {
                plan.Cell(0, column).Range.Text = planHeading[column];
            }
            for (int i = 0; i < milestones.Length; i++)
            {
                plan.Cell(i + 1, 0).Range.Text = milestones[i].Name;
                plan.Cell(i + 1, 1).Range.Text = milestones[i].Owner;
                plan.Cell(i + 1, 2).Range.Text = milestones[i].Due;
                plan.Cell(i + 1, 3).Range.Text = milestones[i].Status;
                plan.Cell(i + 1, 3).Shading.BackgroundPatternColor = milestones[i].Status == "Late" ? "F8CBAD" : milestones[i].Status == "Done" ? "C6E0B4" : "FFE699";
                plan.Cell(i + 1, 3).VerticalAlignment = CellVerticalAlignment.Center;
            }
            TableLook.Grid(plan, true);
            plan.Columns[0].Width = 3600;
            plan.Columns[1].Width = 1900;
            plan.Columns[2].Width = 1700;
            plan.Columns[3].Width = 1300;

            //4. Risks.
            builder.Paragraph("4. Risks", "Heading 1");
            TableWrapper risk = builder.Table(1 + risks.Length, 3);
            risk.Cell(0, 0).Range.Text = "Risk";
            risk.Cell(0, 1).Range.Text = "Severity";
            risk.Cell(0, 2).Range.Text = "Mitigation";
            for (int i = 0; i < risks.Length; i++)
            {
                risk.Cell(i + 1, 0).Range.Text = risks[i].Description;
                risk.Cell(i + 1, 1).Range.Text = risks[i].Severity;
                risk.Cell(i + 1, 2).Range.Text = risks[i].Mitigation;
                risk.Cell(i + 1, 1).Shading.BackgroundPatternColor = risks[i].Severity == "High" ? "F8CBAD" : risks[i].Severity == "Medium" ? "FFE699" : "C6E0B4";
                risk.Rows[i + 1].AllowBreakAcrossPages = false;
            }
            TableLook.Grid(risk, true);
            risk.Columns[0].Width = 3900;
            risk.Columns[1].Width = 1100;
            risk.Columns[2].Width = 3500;

            //5. Outlook.
            builder.Paragraph("5. Outlook", "Heading 1");
            builder.Paragraph("The plan for the fourth quarter assumes that the customer portal goes live by the end of November and that the third supplier is qualified in time for the holiday peak. If either slips, the board will be asked to approve the contingency in section 4.");
            ParagraphWrapper endOfMain = builder.Paragraph("Appendix A follows on landscape pages.");

            //Appendix A: monthly revenue by region, in a section of its own on landscape pages.
            endOfMain.InsertSectionBreak(SectionStart.NewPage);
            builder.Paragraph("Appendix A: monthly revenue (thousands)", "Heading 1");
            string[] months = { "July", "August", "September" };
            TableWrapper monthly = builder.Table(1 + regions.Length, 1 + months.Length + 1);
            monthly.Cell(0, 0).Range.Text = "Region";
            for (int m = 0; m < months.Length; m++)
            {
                monthly.Cell(0, m + 1).Range.Text = months[m];
            }
            monthly.Cell(0, months.Length + 1).Range.Text = "Quarter";
            for (int i = 0; i < regions.Length; i++)
            {
                monthly.Cell(i + 1, 0).Range.Text = regions[i].Name;
                decimal[] share = { 0.31m, 0.33m, 0.36m };
                for (int m = 0; m < months.Length; m++)
                {
                    monthly.Cell(i + 1, m + 1).Range.Text = (regions[i].Current * share[m] / 1000).ToString("N0", CultureInfo.InvariantCulture);
                }
                monthly.Cell(i + 1, months.Length + 1).Range.Text = (regions[i].Current / 1000).ToString("N0", CultureInfo.InvariantCulture);
            }
            TableLook.Grid(monthly, true);
            TableLook.AlignRight(monthly, 1, 2, 3, 4);

            document.Sections[1].PageSetup.Orientation = PageOrientation.Landscape;

            //Comments from two reviewers.
            CommentWrapper question = document.Comments.Add(figure.Range, "Does this include the one-off licence income from the West region?");
            question.Author = "Priya Nair";
            question.Initial = "PN";
            CommentWrapper delay = document.Comments.Add(document.Paragraphs.First(p => p.Text.StartsWith("Key point", StringComparison.Ordinal)).Range, "Please add the revised date for the customer portal.");
            delay.Author = "Tom Hale";
            delay.Initial = "TH";

            //Header and footer. The cover page has neither.
            SectionWrapper first = document.Sections[0];
            first.PageSetup.DifferentFirstPageHeaderFooter = true;
            first.Headers[HeaderFooterIndex.FirstPage].Range.Text = string.Empty;
            first.Footers[HeaderFooterIndex.FirstPage].Range.Text = string.Empty;

            RangeWrapper header = first.Headers[HeaderFooterIndex.Primary].Range;
            header.Text = "Quarterly business report - Q3 2026 - Confidential";
            header.Style = "Header";
            header.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            RangeWrapper footer = first.Footers[HeaderFooterIndex.Primary].Range;
            footer.Text = "Page ";
            footer.Style = "Footer";
            footer.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            document.Fields.Add(footer, FieldType.Page).Result.Text = "2";
            footer.InsertAfter(" of ");
            document.Fields.Add(footer, FieldType.NumPages).Result.Text = "7";

            //Properties: the built-in ones, and some of the company's own.
            DocumentPropertiesWrapper properties = document.BuiltInDocumentProperties;
            properties.Subject = "Third quarter results";
            properties.Company = "Example Ltd";
            properties.Keywords = "quarterly; board; revenue";
            document.CustomDocumentProperties.Add("Quarter", DocumentPropertyType.Text, "Q3 2026");
            document.CustomDocumentProperties.Add("Classification", DocumentPropertyType.Text, "Confidential");
            document.CustomDocumentProperties.Add("Revenue", DocumentPropertyType.FloatingPoint, (double)currentTotal);
        }
    }

    private static void AddCustomStyles(WordprocessingDocumentWrapper document)
    {
        StyleWrapper keyPoint = document.Styles.Add("Key Point", StyleType.Paragraph);
        keyPoint.Font.Bold = true;
        keyPoint.Font.Color = "7F3F00";
        keyPoint.ParagraphFormat.LeftIndent = 14;
        keyPoint.ParagraphFormat.RightIndent = 14;
        keyPoint.ParagraphFormat.SpaceBefore = 8;
        keyPoint.ParagraphFormat.SpaceAfter = 8;
    }

    private static string Money(decimal value)
    {
        return "GBP " + value.ToString("N0", CultureInfo.InvariantCulture);
    }

    private static double Percent(decimal previous, decimal current)
    {
        return (double)((current - previous) / previous * 100);
    }
}
