using System.IO;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The text of a text box is a story of its own, as in Word (#107): <c>TextFrame.TextRange.Find</c> searches it, with positions counted
/// in it, and a search of the main text leaves it out. What Word 16 gave is in the comments.
/// </summary>
public class FindTextBoxTest
{
    private static WordprocessingDocumentWrapper Create(Stream stream, out ShapeWrapper box)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "Main text here";
        box = document.Shapes.AddTextbox(TextOrientation.Horizontal, 20, 30, 120, 60, document.Range(0, 0));
        box.TextFrame.TextRange.Text = "Inside the box\rSecond box line";
        return document;
    }

    [Fact]
    public void FindInATextBoxSearchesItsText()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out ShapeWrapper box);

        //Word: the text range is 0 to 31, and "box" is found at 11 to 14 and then at 22 to 25.
        RangeWrapper range = box.TextFrame.TextRange;
        Assert.Equal(0, range.Start);
        Assert.Equal(31, range.End);
        Assert.True(range.Find.Execute("box"));
        Assert.Equal("box", range.Text);
        Assert.Equal(11, range.Start);
        Assert.Equal(14, range.End);
        Assert.True(range.Find.Execute("box"));
        Assert.Equal(22, range.Start);
        Assert.Equal(25, range.End);
    }

    [Fact]
    public void FindInATextBoxDoesNotFindTheMainText()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out ShapeWrapper box);

        RangeWrapper range = box.TextFrame.TextRange;
        Assert.False(range.Find.Execute("Main"));
        Assert.Equal(0, range.Start);
        Assert.Equal(31, range.End);
    }

    [Theory]
    [InlineData(FindWrap.Stop)]
    [InlineData(FindWrap.Continue)]
    public void FindInTheMainTextDoesNotFindTheTextOfATextBox(FindWrap wrap)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out ShapeWrapper _);

        RangeWrapper range = document.Content;
        Assert.False(range.Find.Execute("box", wrap: wrap));
    }

    [Fact]
    public void ReplaceAllInATextBoxChangesOnlyItsText()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out ShapeWrapper box);

        Assert.True(box.TextFrame.TextRange.Find.Execute("box", replaceWith: "BOX", replace: ReplaceMode.All));

        Assert.Equal("Inside the BOX\rSecond BOX line\r", box.TextFrame.TextRange.Text);
        Assert.Equal("Main text here\r", document.Content.Text);
    }

    [Fact]
    public void ReplaceAllInTheMainTextLeavesATextBoxAlone()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out ShapeWrapper box);
        box.TextFrame.TextRange.Text = "Some text in the box";

        Assert.True(document.Content.Find.Execute("text", replaceWith: "TEXT", replace: ReplaceMode.All));

        Assert.Equal("Main TEXT here\r", document.Content.Text);
        Assert.Equal("Some text in the box\r", box.TextFrame.TextRange.Text);
    }

    [Fact]
    public void TheSectionOfATextBoxIsThatOfItsAnchor()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, out ShapeWrapper box);
        FieldWrapper field = document.Fields.Add(box.TextFrame.TextRange, FieldType.Empty, "SECTION", false);

        Assert.True(field.Update());
        Assert.Equal("1", field.Result.Text);
        Assert.Equal(document.Sections[0], Assert.Single(box.TextFrame.TextRange.Sections));
    }
}
