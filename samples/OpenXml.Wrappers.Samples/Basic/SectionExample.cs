using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Sections and their page setup: page size, orientation and margins, with a landscape section after a portrait one.
/// </summary>
internal static class SectionExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            document.Paragraphs[0].Range.Text = "This page is portrait, A4, with 2 cm margins.";

            //A section break after the first paragraph starts a new page. The paragraph that holds the break ends the first section.
            document.Paragraphs[0].InsertSectionBreak("nextPage");
            document.Paragraphs.Add().Range.Text = "This page is in a second section, in landscape, with narrower margins.";

            //Sections are listed in order, and are 0-based like every other collection. Sizes are in points.
            PageSetupWrapper first = document.Sections[0].PageSetup;
            first.PageWidth = 595.3;
            first.PageHeight = 841.9;
            first.TopMargin = 56.7;
            first.BottomMargin = 56.7;
            first.LeftMargin = 56.7;
            first.RightMargin = 56.7;

            //Changing the orientation swaps the page width and height, as it does in Word.
            PageSetupWrapper second = document.Sections[1].PageSetup;
            second.PageWidth = 595.3;
            second.PageHeight = 841.9;
            second.Orientation = PageOrientation.Landscape;
            second.TopMargin = 36;
            second.BottomMargin = 36;
            second.LeftMargin = 36;
            second.RightMargin = 36;

            document.Paragraphs.Add().Range.Text = string.Format(CultureInfo.InvariantCulture, "This page is {0:0.#} by {1:0.#} points, and the document has {2} sections.", second.PageWidth, second.PageHeight, document.Sections.Count);
        }
    }
}
