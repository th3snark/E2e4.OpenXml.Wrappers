using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;
using OpenXml.Wrappers.Test.Utils;

namespace OpenXml.Wrappers.Test;

public class ParagraphWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public ParagraphWrapperTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void WrappersOfTheSameParagraphAreEqual()
    {
        Paragraph paragraph = new Paragraph();
        ParagraphWrapper first = new ParagraphWrapper(paragraph);
        ParagraphWrapper second = new ParagraphWrapper(paragraph);

        Assert.True(first == second);
        Assert.False(first != second);
        Assert.True(first.Equals((object)second));
        Assert.Equal(first.GetHashCode(), second.GetHashCode());
        Assert.Contains(second, new HashSet<ParagraphWrapper> { first });
    }

    [Fact]
    public void WrappersOfDifferentParagraphsAreNotEqual()
    {
        ParagraphWrapper first = new ParagraphWrapper(new Paragraph());
        ParagraphWrapper second = new ParagraphWrapper(new Paragraph());

        Assert.False(first == second);
        Assert.True(first != second);
        Assert.False(first.Equals((object)"text"));
    }

    [Fact]
    public void ParagraphWrapperComparesCorrectlyWithNull()
    {
        ParagraphWrapper wrapper = new ParagraphWrapper(new Paragraph());
        ParagraphWrapper none = null;

        Assert.False(wrapper == null);
        Assert.False(null == wrapper);
        Assert.True(none == null);
        Assert.True(wrapper != null);
        Assert.False(wrapper.Equals(null));
    }

    [Fact]
    public void ParagraphsFromTheSameDocumentPositionAreEqual()
    {
        Assert.True(_document1.Paragraphs[0] == _document1.Paragraphs[0]);
        Assert.True(_document1.Paragraphs[0] != _document1.Paragraphs[1]);
    }

    [Fact]
    public void RangeTextPropertyIsCorrect()
    {
        Assert.Equal("Test Document\r", _document1.Content.FirstParagraph.Range.Text);
        Assert.Equal("Test Document\r", _document1.Content.FirstParagraph.Text);
    }

    [Fact]
    public void RangeFirstParagraphRangeTextPropertyIsCorrect()
    {
        Assert.Equal("Test Document\r", _document1.Content.FirstParagraph.Range.Text);
        Assert.Equal("Test Document\r", _document1.Content.FirstParagraph.Text);
    }

    [Fact]
    public void SetFontBoldOnParagraphWithoutPropertiesUsesParagraphMarkRunProperties()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("abc")));
        new ParagraphWrapper(paragraph).Range.Font.Bold = true;

        Assert.NotNull(paragraph.ParagraphProperties.ParagraphMarkRunProperties);
        Assert.Empty(paragraph.ParagraphProperties.Elements<RunProperties>());
        Assert.Single(paragraph.ParagraphProperties.ParagraphMarkRunProperties.Elements<Bold>());
        Assert.Single(paragraph.Elements<Run>().First().RunProperties.Elements<Bold>());
    }

    [Fact]
    public void SetFontBoldTwiceDoesNotDuplicateTheParagraphMarkRunProperties()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("abc")));
        ParagraphWrapper wrapper = new ParagraphWrapper(paragraph);
        wrapper.Range.Font.Bold = true;
        wrapper.Range.Font.Bold = true;

        Assert.Single(paragraph.ParagraphProperties.Elements<ParagraphMarkRunProperties>());
        Assert.Single(paragraph.ParagraphProperties.ParagraphMarkRunProperties.Elements<Bold>());
    }

    [Fact]
    public void SetFontBoldFalseTurnsOffExistingBold()
    {
        Paragraph paragraph = new Paragraph(new Run(new RunProperties(new Bold()), new Text("abc")));
        new ParagraphWrapper(paragraph).Range.Font.Bold = false;

        Bold bold = paragraph.Elements<Run>().First().RunProperties.Elements<Bold>().Single();
        Assert.False(bold.Val.Value);
    }

    [Fact]
    public void SetFontColorKeepsSchemaOrderAfterBold()
    {
        Paragraph paragraph = new Paragraph(new Run(new RunProperties(new Bold()), new Text("abc")));
        new ParagraphWrapper(paragraph).Range.Font.Color = "FF0000";

        RunProperties runProperties = paragraph.Elements<Run>().First().RunProperties;
        Assert.IsType<Bold>(runProperties.ChildElements[0]);
        Assert.IsType<Color>(runProperties.ChildElements[1]);
    }

    [Fact]
    public void ParagraphMarkRunPropertiesAreInsertedBeforeSectionProperties()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new SectionProperties()), new Run(new Text("abc")));
        new ParagraphWrapper(paragraph).Range.Font.Color = "FF0000";

        Assert.IsType<ParagraphMarkRunProperties>(paragraph.ParagraphProperties.ChildElements[0]);
        Assert.IsType<SectionProperties>(paragraph.ParagraphProperties.ChildElements[1]);
    }

    [Fact]
    public void CopyPropertiesToWithoutSourceRunPropertiesDoesNotThrow()
    {
        ParagraphWrapper from = new ParagraphWrapper(new Paragraph(new ParagraphProperties(), new Run(new Text("a"))));
        ParagraphWrapper to = new ParagraphWrapper(new Paragraph(new Run(new Text("b"))));
        from.CopyPropertiesTo(to);

        Assert.NotNull(to.Paragraph.ParagraphProperties);
    }

    [Theory]
    [InlineData("Column")]
    [InlineData("TextWrapping")]
    [InlineData("Page")]
    public void InsertBreakAcceptsAllBreakTypes(string name)
    {
        BreakType type = name == "Column" ? BreakType.ColumnBreak : name == "Page" ? BreakType.PageBreak : BreakType.TextWrappingBreak;
        BreakValues expected = name == "Column" ? BreakValues.Column : name == "Page" ? BreakValues.Page : BreakValues.TextWrapping;
        Paragraph paragraph = new Paragraph();
        new ParagraphWrapper(paragraph).InsertBreak(type);

        Assert.Equal(expected, paragraph.Descendants<Break>().Single().Type.Value);
    }

    [Theory]
    [InlineData(BreakType.SectionBreakNextPage)]
    [InlineData((BreakType)99)]
    public void InsertBreakRejectsSectionBreaksAndUnknownTypes(BreakType type)
    {
        Assert.Throws<ArgumentException>(() => new ParagraphWrapper(new Paragraph()).InsertBreak(type));
    }

    [Fact]
    public void ParagraphTextAndRangeTextSettersAgreeOnLineBreaks()
    {
        Paragraph viaParagraph = new Paragraph();
        Paragraph viaRange = new Paragraph();
        new ParagraphWrapper(viaParagraph).Text = "a\nb\tc";
        new ParagraphWrapper(viaRange).Range.Text = "a\nb\tc";

        Assert.Equal(new ParagraphWrapper(viaRange).Range.Text, new ParagraphWrapper(viaParagraph).Range.Text);
        Assert.Equal("a\vb\tc\r", new ParagraphWrapper(viaParagraph).Range.Text);
    }

    [Fact]
    public void SettingTextToNullClearsTheParagraph()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("abc")));
        new ParagraphWrapper(paragraph).Text = null;

        Assert.Empty(paragraph.Elements<Run>());
    }

    [Fact]
    public void SettingTextPreservesLeadingAndTrailingSpaces()
    {
        Paragraph paragraph = new Paragraph();
        new ParagraphWrapper(paragraph).Text = " abc ";

        Assert.Equal(SpaceProcessingModeValues.Preserve, paragraph.Descendants<Text>().Single().Space.Value);
    }

    [Fact]
    public void EmptyParagraphRangeTextIsTheParagraphMarker()
    {
        Assert.Equal("\r", new ParagraphWrapper(new Paragraph()).Range.Text);
        Assert.Equal("\r", new ParagraphWrapper(new Paragraph()).Text);
    }

    [Fact]
    public void RangeTextIncludesRunsNestedInHyperlinks()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a")), new Hyperlink(new Run(new Text("b"))), new Run(new Text("c")));

        Assert.Equal("abc\r", new ParagraphWrapper(paragraph).Range.Text);
    }

    [Fact]
    public void RangeTextPutsTheParagraphMarkerAfterATrailingTab()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a")), new Run(new TabChar()));

        Assert.Equal("a\t\r", new ParagraphWrapper(paragraph).Range.Text);
    }

    [Fact]
    public void LineBreaksRoundTripThroughRangeText()
    {
        Paragraph paragraph = new Paragraph();
        RangeWrapper range = new ParagraphWrapper(paragraph).Range;
        range.Text = "a\nb";

        Assert.Equal("a\vb\r", range.Text);
    }

    [Fact]
    public void ParagraphRangeCanEnumerateItsOwnParagraph()
    {
        ParagraphWrapper wrapper = new ParagraphWrapper(new Paragraph(new Run(new Text("a"))));

        Assert.Equal(1, wrapper.Range.Paragraphs.Count);
        Assert.Equal("a\r", wrapper.Range.FirstParagraph.Text);
    }

    [Fact]
    public void RunPropertiesOfNewRunsAreRunPropertiesNotParagraphMarkRunProperties()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphMarkRunProperties(new Bold())));
        new ParagraphWrapper(paragraph).Range.Text = "abc";

        Run run = paragraph.Elements<Run>().Single();
        Assert.NotNull(run.RunProperties);
        Assert.Empty(run.Elements<ParagraphMarkRunProperties>());
        Assert.Single(run.RunProperties.Elements<Bold>());
    }

    [Fact]
    public void TextEndsWithTheParagraphMark()
    {
        Assert.Equal("a\r", new ParagraphWrapper(new Paragraph(new Run(new Text("a")))).Text);
    }

    [Fact]
    public void SettingTextKeepsTheParagraphMarkSoTheTextCanBeSetBackUnchanged()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a")));
        ParagraphWrapper wrapper = new ParagraphWrapper(paragraph);

        wrapper.Text = wrapper.Text;
        Assert.Equal("a\r", wrapper.Text);

        wrapper.Text = "b\r";
        Assert.Equal("b\r", wrapper.Text);
        Assert.Single(paragraph.Elements<Run>());
    }
}
