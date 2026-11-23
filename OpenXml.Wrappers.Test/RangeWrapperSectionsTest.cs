using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>Range.Sections</c> outside the main text gives one section, as measured in Word 16 (#108): for a footnote or endnote (wherever the
/// endnotes go) the section of its mark, for a comment that of its text, for a text box that of its anchor, and for a header or footer the
/// first section that has it, not one linked to the previous section.
/// </summary>
public class RangeWrapperSectionsTest
{
    [Theory]
    [InlineData(EndnoteLocation.EndOfDocument)]
    [InlineData(EndnoteLocation.EndOfSection)]
    public void SectionsOfOtherStoriesAreThoseWordGives(EndnoteLocation location)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "One\rTwo\rThree";
        foreach (string text in new[] { "One", "Two" })
        {
            ParagraphWrapper paragraph = document.Paragraphs.First(p => p.Range.Text.StartsWith(text, System.StringComparison.Ordinal));
            RangeWrapper at = document.Range(paragraph.Range.End - 1, paragraph.Range.End - 1);
            at.InsertBreak(BreakType.SectionBreakNextPage);
        }
        Assert.Equal(3, document.Sections.Count);

        RangeWrapper second = document.Range(document.Sections[1].Range.Start, document.Sections[1].Range.Start);
        FootnoteWrapper footnote = document.Footnotes.Add(second, null, "F");
        document.Endnotes.Location = location;
        EndnoteWrapper endnote = document.Endnotes.Add(document.Range(document.Sections[1].Range.Start, document.Sections[1].Range.Start), null, "E");

        document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "H1";
        document.Sections[1].Headers[HeaderFooterIndex.Primary].LinkToPrevious = false;
        document.Sections[1].Headers[HeaderFooterIndex.Primary].Range.Text = "H2";

        ShapeWrapper box = document.Shapes.AddTextbox(TextOrientation.Horizontal, 100, 100, 100, 50,
            document.Range(document.Sections[1].Range.Start, document.Sections[1].Range.Start));
        box.TextFrame.TextRange.Text = "Box";

        RangeWrapper third = document.Sections[2].Range;
        CommentWrapper comment = document.Comments.Add(document.Range(third.Start, third.End - 1), "C");

        AssertSection(document, 1, footnote.Range);
        AssertSection(document, 1, footnote.Reference);
        AssertSection(document, 1, endnote.Range);
        AssertSection(document, 0, document.Sections[0].Headers[HeaderFooterIndex.Primary].Range);
        AssertSection(document, 1, document.Sections[1].Headers[HeaderFooterIndex.Primary].Range);
        AssertSection(document, 1, document.Sections[2].Headers[HeaderFooterIndex.Primary].Range);
        AssertSection(document, 0, document.Sections[0].Footers[HeaderFooterIndex.Primary].Range);
        AssertSection(document, 0, document.Sections[2].Footers[HeaderFooterIndex.Primary].Range);
        AssertSection(document, 1, box.TextFrame.TextRange);
        AssertSection(document, 2, comment.Range);
        AssertSection(document, 2, comment.Scope);
    }

    private static void AssertSection(WordprocessingDocumentWrapper document, int index, RangeWrapper range)
    {
        SectionWrapper section = Assert.Single(range.Sections);
        Assert.Equal(document.Sections[index], section);
    }
}
