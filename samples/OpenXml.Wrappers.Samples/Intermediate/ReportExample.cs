using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Intermediate;

/// <summary>
/// A short report with a cover page, a table of contents field, headings, a header and a "Page x of y" footer that leave out the cover
/// page, and a landscape appendix in a section of its own.
/// </summary>
internal static class ReportExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Water use report 2025"))
        {
            DocumentBuilder builder = new DocumentBuilder(document);

            //Cover page.
            builder.Paragraph("Water use report 2025", "Title").Format.SpaceBefore = 160;
            builder.Paragraph("Facilities department, March 2026").Format.SpaceAfter = 24;
            ParagraphWrapper cover = builder.Paragraph("Prepared for the sustainability committee.");
            cover.InsertBreak(BreakType.PageBreak);

            //A table of contents is a field. Word fills it in (from the headings) when the fields are updated.
            builder.Paragraph("Contents", "Heading 1");
            ParagraphWrapper contents = builder.Paragraph(string.Empty);
            FieldWrapper toc = document.Fields.Add(contents.Range, FieldType.TableOfContents, "\\o \"1-2\" \\h \\z \\u");
            toc.Result.Text = "Right-click and choose Update Field to build the table of contents.";
            contents.InsertBreak(BreakType.PageBreak);

            //The body: headings in the styles of the document and some text under each one.
            string[] sections = { "Summary", "Method", "Findings", "Recommendations" };
            foreach (string title in sections)
            {
                builder.Paragraph(title, "Heading 1");
                builder.Paragraph(Filler(title));
                builder.Paragraph("Detail", "Heading 2");
                builder.Paragraph(Filler(title + " detail"));
            }

            //The appendix has a wide table, so it is on landscape pages in a section of its own.
            ParagraphWrapper endOfBody = builder.Paragraph("This is the end of the main report.");
            endOfBody.InsertSectionBreak(SectionStart.NewPage);
            builder.Paragraph("Appendix: monthly readings", "Heading 1");
            TableWrapper table = builder.Table(13, 6);
            string[] heading = { "Month", "Building A", "Building B", "Building C", "Total (m3)", "Change on 2024" };
            for (int column = 0; column < heading.Length; column++)
            {
                table.Cell(0, column).Range.Text = heading[column];
                table.Cell(0, column).Shading.BackgroundPatternColor = "D9E2F3";
            }
            table.Rows[0].Range.Font.Bold = true;
            table.Rows[0].HeadingFormat = true;
            string[] months = CultureInfo.InvariantCulture.DateTimeFormat.MonthNames.Take(12).ToArray();
            for (int month = 0; month < 12; month++)
            {
                int a = 410 + (month * 37) % 90;
                int b = 380 + (month * 53) % 110;
                int c = 295 + (month * 29) % 70;
                table.Cell(month + 1, 0).Range.Text = months[month];
                table.Cell(month + 1, 1).Range.Text = a.ToString(CultureInfo.InvariantCulture);
                table.Cell(month + 1, 2).Range.Text = b.ToString(CultureInfo.InvariantCulture);
                table.Cell(month + 1, 3).Range.Text = c.ToString(CultureInfo.InvariantCulture);
                table.Cell(month + 1, 4).Range.Text = (a + b + c).ToString(CultureInfo.InvariantCulture);
                table.Cell(month + 1, 5).Range.Text = string.Format(CultureInfo.InvariantCulture, "{0:+0.0;-0.0}%", ((month % 5) - 2) * 1.3);
            }
            for (int column = 1; column < 6; column++)
            {
                for (int row = 0; row < table.Rows.Count; row++)
                {
                    table.Cell(row, column).Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                }
            }
            table.Borders.SetOutside(BorderLineStyle.SingleLine, 1, "1F3864");
            table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");

            //The last section is the landscape one. The first section keeps the portrait page the document started with.
            PageSetupWrapper landscape = document.Sections[1].PageSetup;
            landscape.Orientation = PageOrientation.Landscape;

            //Header and footer of the first section: no header or page number on the cover page.
            SectionWrapper first = document.Sections[0];
            first.PageSetup.DifferentFirstPageHeaderFooter = true;
            first.Headers[HeaderFooterIndex.FirstPage].Range.Text = string.Empty;
            first.Footers[HeaderFooterIndex.FirstPage].Range.Text = string.Empty;

            RangeWrapper header = first.Headers[HeaderFooterIndex.Primary].Range;
            header.Text = "Water use report 2025";
            header.Style = "Header";
            header.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            RangeWrapper footer = first.Footers[HeaderFooterIndex.Primary].Range;
            footer.Text = "Page ";
            footer.Style = "Footer";
            footer.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            document.Fields.Add(footer, FieldType.Page).Result.Text = "2";
            footer.InsertAfter(" of ");
            document.Fields.Add(footer, FieldType.NumPages).Result.Text = "6";
        }
    }

    private static string Filler(string topic)
    {
        return string.Format(
            CultureInfo.InvariantCulture,
            "This part of the report covers {0}. Water use across the three buildings fell over the year, with the largest savings in the second half after the leak survey. The figures are taken from the meter readings in the appendix and have not been adjusted for occupancy. Readers who want the readings by month should turn to the appendix at the end of the report. This paragraph is made to be long enough to fill several lines on the page.",
            topic.ToLowerInvariant());
    }
}
