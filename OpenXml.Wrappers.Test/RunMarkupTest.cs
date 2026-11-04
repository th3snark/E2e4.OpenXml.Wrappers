using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The markup of runs and paragraphs, read directly. <see cref="MarkupConformanceTest"/> compares with Word through the normaliser, which puts
/// adjacent runs of the same formatting together on both sides, because Word keeps such runs apart when their revision ids differ.
/// </summary>
public class RunMarkupTest : IDisposable
{
    private readonly MemoryStream _stream = new MemoryStream();
    private readonly WordprocessingDocumentWrapper _document;

    public RunMarkupTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_stream, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
        GC.SuppressFinalize(this);
    }

    private Paragraph FirstParagraph
    {
        get { return _document.Paragraphs[0].Paragraph; }
    }

    [Fact]
    public void RunsOfTheSameFormattingAreMergedAfterFormatting()
    {
        _document.Content.InsertAfter("One two three");

        _document.Range(0, 3).Font.Bold = true;
        _document.Range(3, 7).Font.Bold = true;

        Run[] runs = FirstParagraph.Elements<Run>().ToArray();
        Assert.Equal(2, runs.Length);
        Assert.Equal("One two", runs[0].InnerText);

        _document.Range(0, 7).Font.Bold = false;
        Assert.Equal("One two three", FirstParagraph.Elements<Run>().Single().InnerText);
    }

    [Fact]
    public void RunsAreMergedAfterTextIsInsertedDeletedOrReplaced()
    {
        _document.Content.InsertAfter("One two three");

        _document.Range(3, 3).InsertAfter(" and");
        Assert.Equal("One and two three", FirstParagraph.Elements<Run>().Single().InnerText);

        _document.Range(3, 7).Delete();
        Assert.Equal("One two three", FirstParagraph.Elements<Run>().Single().InnerText);

        _document.Range(4, 7).Text = "2";
        Assert.Equal("One 2 three", FirstParagraph.Elements<Run>().Single().InnerText);
    }

    [Fact]
    public void RangesStayOnTheirTextWhenRunsAreMerged()
    {
        _document.Content.InsertAfter("One two three");
        RangeWrapper three = _document.Range(8, 13);
        RangeWrapper two = _document.Range(4, 7);

        _document.Range(0, 3).Font.Bold = true;
        _document.Range(0, 3).Font.Bold = false;

        Assert.Single(FirstParagraph.Elements<Run>());
        Assert.Equal("two", two.Text);
        Assert.Equal("three", three.Text);

        _document.Range(0, 0).InsertBefore("Zero ");
        Assert.Equal("two", two.Text);
        Assert.Equal(9, two.Start);
        Assert.Equal("three", three.Text);
    }

    [Fact]
    public void FieldsTabsAndMarkersAreNotMergedAway()
    {
        _document.Content.InsertAfter("a\tb");
        _document.Bookmarks.Add("Mark", _document.Range(0, 1));
        _document.Range(0, 3).Font.Italic = true;

        Assert.Single(FirstParagraph.Elements<BookmarkStart>());
        Assert.Equal(2, FirstParagraph.Elements<Run>().Count());
        Assert.Equal("a\tb\r", _document.Paragraphs[0].Text);
    }

    [Theory]
    [InlineData("word", false)]
    [InlineData("two words", false)]
    [InlineData(" word", true)]
    [InlineData("word ", true)]
    public void SpaceIsPreservedOnlyWhereItIsNeeded(string text, bool preserved)
    {
        _document.Content.InsertAfter(text);

        Text element = FirstParagraph.Descendants<Text>().Single();
        Assert.Equal(preserved, element.Space != null);
    }

    [Fact]
    public void SplittingATextOnlyPreservesSpaceOnThePartThatNeedsIt()
    {
        _document.Content.InsertAfter("Bold text");

        _document.Range(0, 4).Font.Bold = true;

        Text[] texts = FirstParagraph.Descendants<Text>().ToArray();
        Assert.Null(texts[0].Space);
        Assert.Equal(SpaceProcessingModeValues.Preserve, texts[1].Space.Value);
    }

    [Fact]
    public void OnIsTheBareElementWithoutItsComplexScriptPartner()
    {
        _document.Content.InsertAfter("Bold");

        _document.Range(0, 4).Font.Bold = true;

        OpenXmlElement bold = FirstParagraph.Elements<Run>().Single().RunProperties.Single();
        Assert.IsType<Bold>(bold);
        Assert.False(bold.HasAttributes);
    }

    [Fact]
    public void OffUnderABoldStyleIsWrittenAsWordWritesIt()
    {
        _document.Content.InsertAfter("Strong");
        _document.Range(0, 6).Style = "Strong";

        _document.Range(0, 6).Font.Bold = false;

        Assert.Contains("w:val=\"0\"", FirstParagraph.Elements<Run>().Single().RunProperties.Elements<Bold>().Single().OuterXml, StringComparison.Ordinal);
        Assert.False(_document.Range(0, 6).Font.Bold);
    }

    [Fact]
    public void AValueTheTextInheritsIsNotWritten()
    {
        _document.Content.InsertAfter("Text");
        double inherited = _document.Range(0, 4).Font.Size.Value;

        _document.Range(0, 4).Font.Size = inherited + 2;
        Assert.NotNull(FirstParagraph.Elements<Run>().Single().RunProperties);

        _document.Range(0, 4).Font.Size = inherited;
        Assert.Null(FirstParagraph.Elements<Run>().Single().RunProperties);
    }

    [Fact]
    public void AParagraphPropertyEqualToItsStyleIsNotWritten()
    {
        _document.Content.InsertAfter("Text");
        ParagraphFormatWrapper format = _document.Paragraphs[0].Range.ParagraphFormat;

        format.Alignment = ParagraphAlignment.Center;
        format.Alignment = ParagraphAlignment.Left;
        format.WidowControl = false;
        format.WidowControl = true;
        format.SpaceAfter = format.SpaceAfter;

        Assert.Null(FirstParagraph.ParagraphProperties);
    }

    [Fact]
    public void AnOffValueOfAParagraphPropertyIsWrittenWhereTheStyleTurnsItOn()
    {
        _document.Content.InsertAfter("Heading");
        _document.Paragraphs[0].Style = "Heading 1";

        _document.Paragraphs[0].Range.ParagraphFormat.KeepWithNext = false;

        Assert.Contains("w:val=\"0\"", FirstParagraph.ParagraphProperties.KeepNext.OuterXml, StringComparison.Ordinal);
        Assert.False(_document.Paragraphs[0].Range.ParagraphFormat.KeepWithNext);
    }

    [Fact]
    public void TheDefaultStyleIsNotNamed()
    {
        _document.Content.InsertAfter("Text");

        _document.Paragraphs[0].Style = "Heading 1";
        _document.Paragraphs[0].Style = "Normal";

        Assert.Null(FirstParagraph.ParagraphProperties);
        Assert.Equal("Normal", _document.Paragraphs[0].Style);
    }

    [Fact]
    public void AddingAParagraphAfterALastParagraphContinuesItAndItsMarkTakesTheFormattingOfTheText()
    {
        _document.Content.InsertAfter("Text");
        _document.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        _document.Range(0, 4).Font.Bold = true;

        ParagraphWrapper added = _document.Paragraphs[0].Range.InsertParagraphAfter();

        Assert.Equal(JustificationValues.Center, added.Paragraph.ParagraphProperties.Justification.Val.Value);
        Assert.Null(added.Paragraph.ParagraphProperties.ParagraphMarkRunProperties);
        Assert.NotNull(FirstParagraph.ParagraphProperties.ParagraphMarkRunProperties.GetFirstChild<Bold>());
    }

    [Fact]
    public void AddingAParagraphAfterAMiddleParagraphTakesTheFormattingOfTheNextOne()
    {
        _document.Content.InsertAfter("One");
        _document.Content.InsertParagraphAfter();
        _document.Content.InsertAfter("Two");
        _document.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        _document.Paragraphs[1].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        ParagraphWrapper added = _document.Paragraphs[0].Range.InsertParagraphAfter();

        Assert.Equal(1, added.Range.Paragraphs.Count);
        Assert.Equal(added, _document.Paragraphs[1]);
        Assert.Equal(ParagraphAlignment.Right, added.Range.ParagraphFormat.Alignment);
    }

    [Fact]
    public void ASectionBreakIsNotCopiedToANewParagraph()
    {
        _document.Content.InsertAfter("First");
        _document.Sections.Add();
        Paragraph ending = _document.Paragraphs[0].Paragraph;
        Assert.NotNull(ending.ParagraphProperties.GetFirstChild<SectionProperties>());

        new RangeWrapper(ending).InsertParagraphBefore();

        Assert.Equal(2, _document.Sections.Count);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.Document.Body.Elements<Paragraph>(), p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null);
    }

    [Fact]
    public void JoiningParagraphsKeepsTheFormattingOfTheFirstAndTheSectionBreakOfTheSecond()
    {
        _document.Content.InsertAfter("First");
        _document.Content.InsertParagraphAfter();
        _document.Content.InsertAfter("Second");
        _document.Paragraphs[1].Range.InsertParagraphAfter();
        _document.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        new ParagraphWrapper(_document.Paragraphs[1].Paragraph).InsertSectionBreak();
        Assert.Equal(2, _document.Sections.Count);

        TextUnitCollection characters = _document.Paragraphs[0].Range.Characters;
        characters[characters.Count - 1].Delete();

        Assert.Equal(2, _document.Sections.Count);
        Assert.Equal("FirstSecond\f", _document.Paragraphs[0].Text);
        Assert.Equal(ParagraphAlignment.Center, _document.Paragraphs[0].Alignment);
    }

    [Fact]
    public void TheMarkOfASectionBreakReadsAsAPageBreak()
    {
        _document.Content.InsertAfter("First");

        _document.Sections.Add();

        Assert.StartsWith("First\f", _document.Content.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void ACarriageReturnInTheTextOfAParagraphStartsAParagraph()
    {
        _document.Paragraphs[0].Range.Text = "One\rTwo\rThree";

        Assert.Equal(3, _document.Paragraphs.Count);
        Assert.Equal("One\r", _document.Paragraphs[0].Text);
        Assert.Equal("Three\r", _document.Paragraphs[2].Text);
        Assert.DoesNotContain(FirstParagraph.Descendants<Text>(), t => t.Text.Contains('\r'));
    }

    [Fact]
    public void ACarriageReturnInsertedIntoAParagraphStartsAParagraph()
    {
        _document.Content.InsertAfter("Text");

        _document.Paragraphs[0].Range.InsertBefore("A\rB");

        Assert.Equal("A\rBText\r", _document.Content.Text);
        Assert.Single(_document.Paragraphs[1].Paragraph.Elements<Run>());
    }

    [Fact]
    public void SettingTheTextOfAParagraphLeavesNoEmptyHyperlinkAndStaysInItsBookmark()
    {
        _document.Content.InsertAfter("Text");
        Paragraph paragraph = FirstParagraph;
        Run run = paragraph.Elements<Run>().Single();
        Hyperlink hyperlink = new Hyperlink(new Run(new Text("link"))) { Anchor = "somewhere" };
        run.InsertAfterSelf(hyperlink);
        _document.Bookmarks.Add("Whole", new RangeWrapper(paragraph));

        _document.Paragraphs[0].Text = "New";

        Assert.Empty(paragraph.Elements<Hyperlink>());
        Assert.Equal("New", _document.Bookmarks["Whole"].Range.Text.TrimEnd('\r'));
    }

    [Fact]
    public void DeletingBreaksLeavesNoEmptyRuns()
    {
        _document.Content.InsertAfter("One\vTwo");

        _document.Paragraphs[0].DeleteBreaks();

        Assert.Equal("OneTwo\r", _document.Paragraphs[0].Text);
        Assert.Single(FirstParagraph.Elements<Run>());
    }

    [Fact]
    public void AFieldThatCanBeWorkedOutIsOneElementUnlessTheTextIsFormatted()
    {
        _document.Content.InsertAfter("Plain bold");
        _document.Range(6, 10).Font.Bold = true;

        FieldWrapper plain = _document.Fields.Add(_document.Range(0, 0), FieldType.Date);
        Assert.Equal(FieldType.Date, plain.Type);
        Assert.Single(FirstParagraph.Elements<SimpleField>());
        Assert.Empty(FirstParagraph.Descendants<FieldChar>());

        int end = _document.Paragraphs[0].Range.End - 1;
        FieldWrapper bold = _document.Fields.Add(_document.Range(end, end), FieldType.Date);
        Assert.Equal(FieldType.Date, bold.Type);
        Assert.Equal(3, FirstParagraph.Descendants<FieldChar>().Count());
        Assert.All(FirstParagraph.Descendants<FieldChar>(), c => Assert.NotNull(((Run)c.Parent).RunProperties.GetFirstChild<Bold>()));
        Assert.Equal(2, _document.Fields.Count);
    }

    [Fact]
    public void AFieldThatNeedsLayoutIsLeftForWord()
    {
        FieldWrapper page = _document.Fields.Add(_document.Range(0, 0), FieldType.Page);

        Assert.Equal(FieldType.Page, page.Type);
        FieldChar begin = FirstParagraph.Descendants<FieldChar>().First();
        Assert.True(begin.Dirty.Value);
        Assert.Empty(FirstParagraph.Descendants<Text>());
        Assert.Equal(1, _document.Fields.Count);
    }

    [Fact]
    public void AFontNameLeavesTheOtherFontsAlone()
    {
        _document.Content.InsertAfter("Text");
        Run run = FirstParagraph.Elements<Run>().Single();
        run.PrependChild(new RunProperties(new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, EastAsia = "MS Gothic", ComplexScript = "Arial" }));

        _document.Range(0, 4).Font.Name = "Courier New";

        RunFonts fonts = run.RunProperties.RunFonts;
        Assert.Equal("Courier New", fonts.Ascii.Value);
        Assert.Equal("Courier New", fonts.HighAnsi.Value);
        Assert.Null(fonts.AsciiTheme);
        Assert.Equal("MS Gothic", fonts.EastAsia.Value);
        Assert.Equal("Arial", fonts.ComplexScript.Value);
    }
}
