using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>Range.TextRetrievalMode</c>, <c>Range.Document</c> and <c>Range.Parent</c>. The text is what Word 16 gave for the same content
/// (a paragraph "before " + a DATE field with the result 05/11/2026 + " after", 33 positions, and "visible hidden end" with "hidden" hidden).
/// </summary>
public class RangeWrapperRetrievalTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly Body _body;

    public RangeWrapperRetrievalTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        _body.RemoveAllChildren<Paragraph>();
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Run Mark(FieldCharValues type)
    {
        return new Run(new FieldChar() { FieldCharType = type });
    }

    private RangeWrapper ComplexField()
    {
        _body.InsertAt(
            new Paragraph(
                Plain("before "),
                Mark(FieldCharValues.Begin),
                new Run(new FieldCode(" DATE ") { Space = SpaceProcessingModeValues.Preserve }),
                Mark(FieldCharValues.Separate),
                Plain("05/11/2026"),
                Mark(FieldCharValues.End),
                Plain(" after")),
            0);
        return new RangeWrapper(_body);
    }

    private RangeWrapper SimpleFieldParagraph()
    {
        _body.InsertAt(new Paragraph(Plain("before "), new SimpleField(Plain("05/11/2026")) { Instruction = " DATE " }, Plain(" after")), 0);
        return new RangeWrapper(_body);
    }

    private RangeWrapper HiddenText()
    {
        Run hidden = Plain("hidden");
        hidden.PrependChild(new RunProperties(new Vanish()));
        _body.InsertAt(new Paragraph(Plain("visible "), hidden, Plain(" end")), 0);
        return new RangeWrapper(_body);
    }

    [Fact]
    public void Defaults_AreWordsDefaults()
    {
        RangeWrapper range = ComplexField();

        Assert.False(range.TextRetrievalMode.IncludeFieldCodes);
        Assert.True(range.TextRetrievalMode.IncludeHiddenText);
        Assert.Equal(ViewType.Print, range.TextRetrievalMode.ViewType);
        Assert.Equal("before 05/11/2026 after\r", range.Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IncludeFieldCodes_ShowsTheCodeInsteadOfTheResult(bool simple)
    {
        RangeWrapper range = simple ? SimpleFieldParagraph() : ComplexField();
        range.TextRetrievalMode.IncludeFieldCodes = true;

        Assert.Equal("before \u0013 DATE \u0015 after\r", range.Text);
        Assert.Equal(33, range.End);
    }

    [Theory]
    [InlineData(false, 0, 20, "before \u0013 DATE ")]
    [InlineData(false, 7, 12, "\u0013 DAT")]
    [InlineData(false, 15, 25, "")]
    [InlineData(false, 25, 26, "\u0015")]
    [InlineData(true, 0, 20, "before \u0013 DATE ")]
    [InlineData(true, 7, 12, "\u0013 DAT")]
    [InlineData(true, 15, 25, "")]
    [InlineData(true, 25, 26, "\u0015")]
    public void IncludeFieldCodes_PartOfAField(bool simple, int start, int end, string expected)
    {
        RangeWrapper whole = simple ? SimpleFieldParagraph() : ComplexField();
        whole.SetRange(start, end);
        whole.TextRetrievalMode.IncludeFieldCodes = true;

        Assert.Equal(expected, whole.Text);
    }

    [Fact]
    public void FieldCodesAreLeftOutOfARangeInsideTheCodeUnlessAsked()
    {
        RangeWrapper range = ComplexField();
        range.SetRange(7, 12);

        Assert.Equal(string.Empty, range.Text);
        range.TextRetrievalMode.IncludeFieldCodes = true;
        Assert.Equal("\u0013 DAT", range.Text);
        range.TextRetrievalMode.IncludeFieldCodes = false;
        Assert.Equal(string.Empty, range.Text);
    }

    [Fact]
    public void IncludeHiddenText_False_LeavesHiddenTextOut()
    {
        RangeWrapper range = HiddenText();

        Assert.Equal("visible hidden end\r", range.Text);
        range.TextRetrievalMode.IncludeHiddenText = false;
        Assert.Equal("visible  end\r", range.Text);
        range.TextRetrievalMode.IncludeHiddenText = true;
        Assert.Equal("visible hidden end\r", range.Text);
    }

    [Fact]
    public void HiddenText_InAStyle_IsHidden()
    {
        RangeWrapper range = HiddenText();
        Paragraph paragraph = _body.GetFirstChild<Paragraph>();
        paragraph.RemoveAllChildren<Run>();
        paragraph.AppendChild(Plain("text"));
        Style style = new Style(new StyleName() { Val = "Secret" }, new StyleRunProperties(new Vanish())) { Type = StyleValues.Character, StyleId = "Secret" };
        _document.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles.AppendChild(style);
        paragraph.GetFirstChild<Run>().PrependChild(new RunProperties(new RunStyle() { Val = "Secret" }));

        range.TextRetrievalMode.IncludeHiddenText = false;

        Assert.Equal("\r", range.Text);
    }

    [Fact]
    public void ARangeMadeFromAnotherKeepsTheSettings_ANewOneDoesNot()
    {
        RangeWrapper range = ComplexField();
        range.TextRetrievalMode.IncludeFieldCodes = true;
        range.TextRetrievalMode.ViewType = ViewType.Web;

        RangeWrapper copy = range.Duplicate();
        Assert.True(copy.TextRetrievalMode.IncludeFieldCodes);
        Assert.Equal(ViewType.Web, copy.TextRetrievalMode.ViewType);

        copy.TextRetrievalMode.IncludeFieldCodes = false;
        Assert.True(range.TextRetrievalMode.IncludeFieldCodes);

        range.Collapse(CollapseDirection.Start);
        range.MoveEnd(TextUnit.Character, 3);
        Assert.True(range.TextRetrievalMode.IncludeFieldCodes);
        Assert.Equal("bef", range.Text);

        Assert.False(new RangeWrapper(_body).TextRetrievalMode.IncludeFieldCodes);
    }

    [Fact]
    public void ViewType_RejectsAValueThatIsNotAView()
    {
        RangeWrapper range = ComplexField();

        Assert.Throws<ArgumentOutOfRangeException>(() => range.TextRetrievalMode.ViewType = (ViewType)99);
    }

    [Fact]
    public void Duplicate_OfTheSettings_IsIndependent()
    {
        TextRetrievalMode mode = new TextRetrievalMode() { IncludeFieldCodes = true, IncludeHiddenText = false, ViewType = ViewType.Outline };

        TextRetrievalMode copy = mode.Duplicate();
        copy.IncludeFieldCodes = false;

        Assert.True(mode.IncludeFieldCodes);
        Assert.False(copy.IncludeHiddenText);
        Assert.Equal(ViewType.Outline, copy.ViewType);
    }

    [Fact]
    public void Document_IsTheWrapperTheDocumentWasCreatedWith()
    {
        RangeWrapper range = ComplexField();

        Assert.Same(_document, range.Document);
        Assert.Same(_document, range.Parent);
        Assert.Same(_document, _document.Range(0, 3).Document);
    }

    [Fact]
    public void Document_OfAHeaderRange_IsTheDocument()
    {
        _document.Content.Text = "x";
        SectionWrapper section = _document.Content.Sections[0];
        HeaderFooterWrapper header = section.Headers[HeaderFooterIndex.Primary];
        header.Range.Text = "head";

        Assert.Same(_document, header.Range.Document);
    }

    [Fact]
    public void Document_OfAnotherWrapperOverTheSamePackage_IsTheFirst()
    {
        WordprocessingDocumentWrapper second = new WordprocessingDocumentWrapper(_document.WordprocessingDocument);

        Assert.Same(_document, new RangeWrapper(_body).Document);
        Assert.NotSame(second, new RangeWrapper(_body).Document);
    }

    [Fact]
    public void Document_OfElementsNotInADocument_IsNull()
    {
        Body loose = new Body(new Paragraph(Plain("x")));
        RangeWrapper range = new RangeWrapper(loose);

        Assert.Null(range.Document);
        Assert.Null(range.Parent);
    }
}
