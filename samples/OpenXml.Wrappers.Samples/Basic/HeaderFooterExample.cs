using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Headers and footers: text in them, page number fields, and a different first page.
/// </summary>
internal static class HeaderFooterExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Headers and footers"))
        {
            //Three pages of text, with a page break between them.
            for (int page = 1; page <= 3; page++)
            {
                ParagraphWrapper paragraph = page == 1 ? document.Paragraphs[0] : document.Paragraphs.Add();
                paragraph.Range.Text = string.Format(CultureInfo.InvariantCulture, "Page {0} of the example", page);
                paragraph.Range.Style = "Heading 1";
                if (page < 3)
                {
                    paragraph.InsertBreak(BreakType.PageBreak);
                }
            }

            SectionWrapper section = document.Sections[0];

            //Reading Range creates the header or footer if it does not exist yet. Setting Text replaces its content.
            RangeWrapper header = section.Headers[HeaderFooterIndex.Primary].Range;
            header.Text = "E2e4.OpenXml.Wrappers - headers and footers";
            header.Style = "Header";
            header.ParagraphFormat.Alignment = ParagraphAlignment.Right;

            //A footer that reads "Page x of y": text and fields can be added one after the other at the end of the range.
            RangeWrapper footer = section.Footers[HeaderFooterIndex.Primary].Range;
            footer.Text = "Page ";
            footer.Style = "Footer";
            footer.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            footer.Fields.Add(footer, FieldType.Page).Result.Text = "1";
            footer.InsertAfter(" of ");
            footer.Fields.Add(footer, FieldType.NumPages).Result.Text = "3";

            //The first page has its own header and footer. This is switched on in the page setup of the section.
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
            RangeWrapper firstHeader = section.Headers[HeaderFooterIndex.FirstPage].Range;
            firstHeader.Text = "Cover page - no page number";
            firstHeader.Style = "Header";
            section.Footers[HeaderFooterIndex.FirstPage].Range.Text = string.Empty;
        }
    }
}
