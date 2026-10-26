using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Members that follow Word's object model: breaks, the formatting of paragraphs and fonts, and tab stops.
/// </summary>
public class WordMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public WordMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "Hello world";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        string[] errors = new OpenXmlValidator().Validate(_document.WordprocessingDocument).Select(e => e.Description + " at " + e.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, string.Join("; ", errors));
    }

    #region Breaks

    [Theory]
    [InlineData(BreakType.PageBreak, "\f", "page")]
    [InlineData(BreakType.ColumnBreak, "\u000e", "column")]
    [InlineData(BreakType.LineBreak, "\v", null)]
    public void ABreakReplacesASpanAndIsReadBackAsItsCharacter(BreakType type, string character, string xmlType)
    {
        RangeWrapper range = _document.Range(5, 6);

        range.InsertBreak(type);

        Assert.Equal("Hello" + character + "world" + "\r", _document.Paragraphs[0].Text);
        Break inserted = _document.Paragraphs[0].Paragraph.Descendants<Break>().Single();
        Assert.Equal(xmlType, inserted.Type?.InnerText);
        AssertValid();
    }

    [Fact]
    public void ACollapsedSpanKeepsItsTextWhenABreakIsInserted()
    {
        RangeWrapper range = _document.Range(5, 5);

        range.InsertBreak(BreakType.PageBreak);

        Assert.Equal("Hello\f world\r", _document.Paragraphs[0].Text);
    }

    [Fact]
    public void ASectionBreakSplitsTheParagraphAndEndsTheSectionBeforeTheRange()
    {
        _document.Range(5, 5).InsertBreak(BreakType.SectionBreakContinuous);

        Assert.Equal(2, _document.Sections.Count);
        Assert.Equal(2, _document.Paragraphs.Count);
        Assert.Equal("Hello\r", _document.Paragraphs[0].Text);
        Assert.Equal(" world\r", _document.Paragraphs[1].Text);
        Assert.NotNull(_document.Paragraphs[0].Paragraph.ParagraphProperties.SectionProperties);
        Assert.Equal(SectionStart.Continuous, _document.Sections[1].PageSetup.SectionStart);
        AssertValid();
    }

    [Fact]
    public void ASectionBreakBeforeAParagraphGetsAParagraphOfItsOwn()
    {
        _document.Paragraphs[0].Range.InsertBreak(BreakType.SectionBreakNextPage);

        Assert.Equal(2, _document.Sections.Count);
        Assert.Equal("\r", _document.Paragraphs[0].Text);
        Assert.Equal("Hello world\r", _document.Paragraphs[1].Text);
        AssertValid();
    }

    [Fact]
    public void AnUndefinedBreakTypeIsRejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Range(0, 0).InsertBreak((BreakType)99));
    }

    #endregion

    #region Paragraph

    [Fact]
    public void AParagraphHasTheFormattingOfItsFormat()
    {
        ParagraphWrapper paragraph = _document.Paragraphs[0];

        paragraph.Alignment = ParagraphAlignment.Center;
        paragraph.LeftIndent = 36;
        paragraph.RightIndent = 18;
        paragraph.FirstLineIndent = -9;
        paragraph.LineSpacingRule = LineSpacingRule.DoubleSpacing;
        paragraph.KeepTogether = true;
        paragraph.PageBreakBefore = true;
        paragraph.WidowControl = true;

        Assert.Equal(ParagraphAlignment.Center, paragraph.Alignment);
        Assert.Equal(36, paragraph.Format.LeftIndent);
        Assert.Equal(18, paragraph.RightIndent);
        Assert.Equal(-9, paragraph.FirstLineIndent);
        Assert.Equal(LineSpacingRule.DoubleSpacing, paragraph.LineSpacingRule);
        Assert.Equal(24, paragraph.LineSpacing);
        Assert.True(paragraph.KeepTogether);
        Assert.True(paragraph.PageBreakBefore);
        Assert.True(paragraph.WidowControl);
        AssertValid();
    }

    [Fact]
    public void AParagraphStyleCanBeSetOnAParagraphAndAllParagraphs()
    {
        _document.Paragraphs.Add().Text = "Second";

        _document.Paragraphs[0].Style = "Heading 1";
        Assert.Equal("Heading 1", _document.Paragraphs[0].Style);
        Assert.Equal("Normal", _document.Paragraphs[1].Style);
        Assert.Null(_document.Paragraphs.Style);

        _document.Paragraphs.Style = "Heading 2";
        Assert.Equal("Heading 2", _document.Paragraphs.Style);
        Assert.Equal(ParagraphOutlineLevel.Level2, _document.Paragraphs.Format.OutlineLevel);
    }

    [Fact]
    public void ParagraphsHaveAFirstAndLastAndNeighboursByCount()
    {
        _document.Paragraphs.Add().Text = "Second";
        _document.Paragraphs.Add().Text = "Third";

        Assert.Equal("Hello world\r", _document.Paragraphs.First.Text);
        Assert.Equal("Third\r", _document.Paragraphs.Last.Text);
        Assert.Equal("Third\r", _document.Paragraphs[0].Next(2).Text);
        Assert.Equal("Hello world\r", _document.Paragraphs[2].Previous(2).Text);
        Assert.Null(_document.Paragraphs[2].Next());
        Assert.Null(_document.Paragraphs[0].Previous(1));
    }

    #endregion

    #region Font

    [Fact]
    public void UnderlineHasAStyleAndAColour()
    {
        FontWrapper font = _document.Paragraphs[0].Range.Font;
        Assert.Equal(UnderlineStyle.None, font.Underline);

        font.Underline = UnderlineStyle.Wavy;
        font.UnderlineColor = "FF0000";
        Assert.Equal(UnderlineStyle.Wavy, font.Underline);
        Assert.Equal("FF0000", font.UnderlineColor);

        font.Underline = UnderlineStyle.DoubleLine;
        Assert.Equal("FF0000", font.UnderlineColor);

        font.Underline = UnderlineStyle.None;
        Assert.Equal(UnderlineStyle.None, font.Underline);
        font.UnderlineColor = null;
        AssertValid();
    }

    [Fact]
    public void SuperscriptAndSubscriptExcludeEachOther()
    {
        FontWrapper font = _document.Range(0, 5).Font;

        font.Superscript = true;
        Assert.True(font.Superscript);
        Assert.False(font.Subscript);

        font.Subscript = true;
        Assert.False(font.Superscript);
        Assert.True(font.Subscript);

        font.Superscript = false;
        Assert.True(font.Subscript);
        font.Subscript = false;
        Assert.False(font.Subscript);
        Assert.Null(_document.Paragraphs[0].Paragraph.Descendants<VerticalTextAlignment>().FirstOrDefault());
    }

    [Fact]
    public void TheOnAndOffPropertiesAreStoredAndRead()
    {
        FontWrapper font = _document.Paragraphs[0].Range.Font;

        font.AllCaps = true;
        font.SmallCaps = true;
        font.Hidden = true;
        font.Emboss = true;
        font.Engrave = true;
        font.Outline = true;
        font.Shadow = true;

        Assert.True(font.AllCaps);
        Assert.True(font.SmallCaps);
        Assert.True(font.Hidden);
        Assert.True(font.Emboss);
        Assert.True(font.Engrave);
        Assert.True(font.Outline);
        Assert.True(font.Shadow);
        AssertValid();

        font.Reset();
        Assert.False(font.AllCaps);
        Assert.False(font.Shadow);
        Assert.Empty(_document.Paragraphs[0].Paragraph.Descendants<RunProperties>());
    }

    [Fact]
    public void DoubleStrikeThroughReplacesStrikeThrough()
    {
        FontWrapper font = _document.Paragraphs[0].Range.Font;

        font.StrikeThrough = true;
        font.DoubleStrikeThrough = true;

        Assert.False(font.StrikeThrough);
        Assert.True(font.DoubleStrikeThrough);
        AssertValid();
    }

    [Fact]
    public void SpacingScalingPositionAndKerningAreInTheirUnits()
    {
        FontWrapper font = _document.Paragraphs[0].Range.Font;
        Assert.Equal(0, font.Spacing);
        Assert.Equal(100, font.Scaling);
        Assert.Equal(0, font.Position);
        Assert.Equal(0, font.Kerning);

        font.Spacing = 1.5;
        font.Scaling = 150;
        font.Position = -3;
        font.Kerning = 14;

        Run run = _document.Paragraphs[0].Paragraph.Elements<Run>().First();
        Assert.Equal(30, run.RunProperties.GetFirstChild<Spacing>().Val.Value);
        Assert.Equal(150, run.RunProperties.GetFirstChild<CharacterScale>().Val.Value);
        Assert.Equal("-6", run.RunProperties.GetFirstChild<Position>().Val.Value);
        Assert.Equal(28u, run.RunProperties.GetFirstChild<Kern>().Val.Value);
        Assert.Equal(1.5, font.Spacing);
        Assert.Equal(150, font.Scaling);
        Assert.Equal(-3, font.Position);
        Assert.Equal(14, font.Kerning);
        AssertValid();

        font.Spacing = 0;
        font.Scaling = 100;
        Assert.Null(run.RunProperties.GetFirstChild<Spacing>());
        Assert.Null(run.RunProperties.GetFirstChild<CharacterScale>());
        Assert.Throws<ArgumentOutOfRangeException>(() => font.Scaling = 0);
    }

    #endregion

    #region ParagraphFormat

    [Fact]
    public void TheOtherParagraphFormatPropertiesAreStoredAndRead()
    {
        ParagraphFormatWrapper format = _document.Paragraphs[0].Format;
        Assert.Equal(ParagraphOutlineLevel.BodyText, format.OutlineLevel);
        Assert.True(format.Hyphenation);

        format.OutlineLevel = ParagraphOutlineLevel.Level3;
        format.NoLineNumber = true;
        format.MirrorIndents = true;
        format.Hyphenation = false;
        format.SpaceBeforeAuto = true;
        format.SpaceAfterAuto = true;

        Assert.Equal(ParagraphOutlineLevel.Level3, format.OutlineLevel);
        Assert.Equal(2, _document.Paragraphs[0].Paragraph.ParagraphProperties.OutlineLevel.Val.Value);
        Assert.True(format.NoLineNumber);
        Assert.True(format.MirrorIndents);
        Assert.False(format.Hyphenation);
        Assert.True(format.SpaceBeforeAuto);
        Assert.True(format.SpaceAfterAuto);
        AssertValid();

        format.Reset();
        Assert.Equal(ParagraphOutlineLevel.BodyText, format.OutlineLevel);
        Assert.True(format.Hyphenation);
        Assert.False(format.NoLineNumber);
    }

    [Fact]
    public void TabStopsAreAddedInOrderWithAlignmentAndLeader()
    {
        TabStopWrapperCollection tabs = _document.Paragraphs[0].Format.TabStops;
        Assert.Equal(0, tabs.Count);

        tabs.Add(144, TabAlignment.Right, TabLeader.Dots);
        tabs.Add(72);

        Assert.Equal(2, tabs.Count);
        Assert.Equal(72, tabs[0].Position);
        Assert.Equal(TabAlignment.Left, tabs[0].Alignment);
        Assert.Equal(144, tabs[1].Position);
        Assert.Equal(TabAlignment.Right, tabs[1].Alignment);
        Assert.Equal(TabLeader.Dots, tabs[1].Leader);
        AssertValid();

        tabs[1].Alignment = TabAlignment.Center;
        tabs[1].Leader = TabLeader.Lines;
        tabs[0].Position = 200;
        Assert.Equal(new double[] { 144, 200 }, tabs.Select(t => t.Position).ToArray());
        Assert.Equal(TabAlignment.Center, tabs[0].Alignment);

        tabs[0].Clear();
        Assert.Equal(1, tabs.Count);
        tabs.ClearAll();
        Assert.Equal(0, tabs.Count);
        Assert.Null(_document.Paragraphs[0].Paragraph.ParagraphProperties?.Tabs);
    }

    [Fact]
    public void TabStopsOfAStyleAreInheritedAndClearedWithAClearTab()
    {
        StyleWrapper style = _document.Styles.Add("Tabbed", StyleType.Paragraph);
        style.ParagraphFormat.TabStops.Add(100, TabAlignment.DecimalPoint);
        _document.Paragraphs[0].Style = "Tabbed";
        TabStopWrapperCollection tabs = _document.Paragraphs[0].Format.TabStops;

        Assert.Equal(1, tabs.Count);
        Assert.Equal(TabAlignment.DecimalPoint, tabs[0].Alignment);
        Assert.Equal(0, _document.Paragraphs[0].Format.Direct.TabStops.Count);

        tabs.Add(50);
        Assert.Equal(new double[] { 50, 100 }, tabs.Select(t => t.Position).ToArray());

        tabs.ClearAll();
        Assert.Equal(0, tabs.Count);
        Assert.Equal(1, style.ParagraphFormat.TabStops.Count);
        Assert.Contains(_document.Paragraphs[0].Paragraph.ParagraphProperties.Tabs.Elements<TabStop>(), t => t.Val.Value == TabStopValues.Clear);
        AssertValid();
    }

    #endregion
}
