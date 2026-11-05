using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The spacing and indent shortcuts of <see cref="ParagraphFormatWrapper"/> and <see cref="ParagraphWrapper"/>, the East Asian and reading
/// order switches, and Duplicate. The numbers are those Word 16 gave and the markup is what it saved (the paragraph-* scenarios of
/// MarkupConformanceTest).
/// </summary>
public class ParagraphFormatShortcutsTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ParagraphFormatShortcutsTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "One\rTwo\rThree\rFour";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private ParagraphWrapper Paragraph(int index)
    {
        return _document.Content.Paragraphs[index];
    }

    private Body Body
    {
        get { return _document.WordprocessingDocument.MainDocumentPart.Document.Body; }
    }

    [Fact]
    public void OpenUpAndCloseUp_PutTwelvePointsBeforeAndTakeThemAway()
    {
        ParagraphWrapper paragraph = Paragraph(0);

        paragraph.OpenUp();
        Assert.Equal(12, paragraph.Format.SpaceBefore);
        Assert.Equal("240", Body.Descendants<SpacingBetweenLines>().Single().Before.Value);

        paragraph.CloseUp();
        Assert.Equal(0, paragraph.Format.SpaceBefore);
        Assert.Empty(Body.Descendants<SpacingBetweenLines>());
    }

    [Fact]
    public void OpenOrCloseUp_Toggles()
    {
        ParagraphWrapper paragraph = Paragraph(0);

        paragraph.OpenOrCloseUp();
        Assert.Equal(12, paragraph.Format.SpaceBefore);
        paragraph.OpenOrCloseUp();
        Assert.Equal(0, paragraph.Format.SpaceBefore);

        paragraph.Format.SpaceBefore = 6;
        paragraph.OpenOrCloseUp();
        Assert.Equal(0, paragraph.Format.SpaceBefore);
        paragraph.OpenUp();
        Assert.Equal(12, paragraph.Format.SpaceBefore);
    }

    [Fact]
    public void Space1_Space15_Space2_SetTheRuleAndTheMultiple()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;

        format.Space1();
        Assert.Equal(LineSpacingRule.SingleSpacing, format.LineSpacingRule);
        Assert.Equal(12, format.LineSpacing);

        format.Space15();
        Assert.Equal(LineSpacingRule.OnePointFive, format.LineSpacingRule);
        Assert.Equal(18, format.LineSpacing);

        format.Space2();
        Assert.Equal(LineSpacingRule.DoubleSpacing, format.LineSpacingRule);
        Assert.Equal(24, format.LineSpacing);
    }

    [Fact]
    public void IndentAndOutdent_MoveBetweenDefaultTabStops()
    {
        ParagraphWrapper paragraph = Paragraph(0);

        paragraph.Indent();
        Assert.Equal(36, paragraph.LeftIndent);
        paragraph.Indent();
        Assert.Equal(72, paragraph.LeftIndent);
        paragraph.Outdent();
        Assert.Equal(36, paragraph.LeftIndent);
        paragraph.Outdent();
        paragraph.Outdent();
        Assert.Equal(0, paragraph.LeftIndent);

        paragraph.LeftIndent = 20;
        paragraph.Indent();
        Assert.Equal(36, paragraph.LeftIndent);
        paragraph.Outdent();
        Assert.Equal(0, paragraph.LeftIndent);
    }

    [Fact]
    public void Indent_KeepsTheFirstLineIndent()
    {
        ParagraphWrapper paragraph = Paragraph(0);
        paragraph.FirstLineIndent = 18;

        paragraph.Indent();

        Assert.Equal(36, paragraph.LeftIndent);
        Assert.Equal(18, paragraph.FirstLineIndent);
    }

    [Fact]
    public void Indent_UsesTheDefaultTabStopOfTheDocument()
    {
        _document.WordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<DefaultTabStop>().Val = 567;
        ParagraphWrapper paragraph = Paragraph(0);

        paragraph.Indent();
        paragraph.Indent();

        Assert.Equal(56.7, paragraph.LeftIndent.Value, 1);
    }

    [Fact]
    public void TabIndent_MovesByStops_NeverBelowZero()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;

        format.TabIndent(2);
        Assert.Equal(72, format.LeftIndent);
        format.TabIndent(-1);
        Assert.Equal(36, format.LeftIndent);
        format.TabIndent(-5);
        Assert.Equal(0, format.LeftIndent);
    }

    [Fact]
    public void TabHangingIndent_IndentsAndHangsTheFirstLineOut()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;

        format.TabHangingIndent(1);
        Assert.Equal(36, format.LeftIndent);
        Assert.Equal(-36, format.FirstLineIndent);
        format.TabHangingIndent(1);
        Assert.Equal(72, format.LeftIndent);
        Assert.Equal(-72, format.FirstLineIndent);
        format.TabHangingIndent(-1);
        Assert.Equal(36, format.LeftIndent);
        Assert.Equal(-36, format.FirstLineIndent);
        format.TabHangingIndent(-1);
        Assert.Equal(0, format.LeftIndent);
        Assert.Equal(0, format.FirstLineIndent);
    }

    [Fact]
    public void ParagraphsShortcuts_ApplyToEveryParagraphOfARange()
    {
        RangeWrapper range = _document.Range(0, 8);

        range.ParagraphFormat.OpenUp();

        Assert.Equal(12, Paragraph(0).Format.SpaceBefore);
        Assert.Equal(12, Paragraph(1).Format.SpaceBefore);
        Assert.Equal(0, Paragraph(2).Format.SpaceBefore);
    }

    [Fact]
    public void Switches_HaveWordsDefaults()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;

        Assert.True(format.WordWrap);
        Assert.True(format.AutoAdjustRightIndent);
        Assert.True(format.AddSpaceBetweenFarEastAndAlpha);
        Assert.True(format.AddSpaceBetweenFarEastAndDigit);
        Assert.True(format.FarEastLineBreakControl);
        Assert.True(format.HangingPunctuation);
        Assert.False(format.HalfWidthPunctuationOnTopOfLine);
        Assert.False(format.DisableLineHeightGrid);
        Assert.Equal(ReadingOrder.LeftToRight, format.ReadingOrder);
        Assert.Equal(BaselineAlignment.Auto, format.BaseLineAlignment);
        Assert.Equal(TextboxTightWrap.None, format.TextboxTightWrap);
    }

    [Fact]
    public void Switches_AreWrittenAsWordWritesThem()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;

        format.WordWrap = false;
        format.HalfWidthPunctuationOnTopOfLine = true;
        format.DisableLineHeightGrid = true;
        format.BaseLineAlignment = BaselineAlignment.Far;
        format.TextboxTightWrap = TextboxTightWrap.FirstAndLastLine;
        format.ReadingOrder = ReadingOrder.RightToLeft;

        Assert.False(format.WordWrap);
        Assert.True(format.HalfWidthPunctuationOnTopOfLine);
        Assert.True(format.DisableLineHeightGrid);
        Assert.Equal(BaselineAlignment.Far, format.BaseLineAlignment);
        Assert.Equal(TextboxTightWrap.FirstAndLastLine, format.TextboxTightWrap);
        Assert.Equal(ReadingOrder.RightToLeft, format.ReadingOrder);
        Assert.False(Body.Descendants<WordWrap>().Single().Val.Value);
        Assert.False(Body.Descendants<SnapToGrid>().Single().Val.Value);
        Assert.Null(Body.Descendants<TopLinePunctuation>().Single().Val);
        Assert.Null(Body.Descendants<BiDi>().Single().Val);
    }

    [Fact]
    public void Switches_SetToTheirDefault_LeaveNothing()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;
        format.WordWrap = false;
        format.DisableLineHeightGrid = true;
        format.ReadingOrder = ReadingOrder.RightToLeft;
        format.BaseLineAlignment = BaselineAlignment.Top;

        format.WordWrap = true;
        format.DisableLineHeightGrid = false;
        format.ReadingOrder = ReadingOrder.LeftToRight;
        format.BaseLineAlignment = BaselineAlignment.Auto;

        Assert.Empty(Body.Descendants<ParagraphProperties>());
    }

    [Fact]
    public void Switches_Null_RemoveTheDirectValue()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;
        format.FarEastLineBreakControl = false;

        format.FarEastLineBreakControl = null;

        Assert.True(format.FarEastLineBreakControl);
        Assert.Empty(Body.Descendants<Kinsoku>());
    }

    [Fact]
    public void Enums_RejectValuesThatAreNotMembers()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;

        Assert.Throws<ArgumentOutOfRangeException>(() => format.ReadingOrder = (ReadingOrder)5);
        Assert.Throws<ArgumentOutOfRangeException>(() => format.BaseLineAlignment = (BaselineAlignment)7);
        Assert.Throws<ArgumentOutOfRangeException>(() => format.TextboxTightWrap = (TextboxTightWrap)9);
    }

    [Fact]
    public void Switches_OfMixedParagraphs_AreNull()
    {
        Paragraph(0).Format.WordWrap = false;

        Assert.Null(_document.Range(0, 8).ParagraphFormat.WordWrap);
    }

    [Fact]
    public void Duplicate_KeepsTheFormattingTheParagraphHasNow_AndCanBeAssigned()
    {
        ParagraphFormatWrapper format = Paragraph(0).Format;
        format.LeftIndent = 36;
        format.SpaceBefore = 12;
        format.Space15();
        format.KeepWithNext = true;
        format.WordWrap = false;
        format.Alignment = ParagraphAlignment.Right;

        ParagraphFormatWrapper copy = format.Duplicate();
        format.LeftIndent = 72;
        Assert.Equal(36, copy.LeftIndent);

        Paragraph(2).Format = copy;

        ParagraphFormatWrapper target = Paragraph(2).Format;
        Assert.Equal(36, target.LeftIndent);
        Assert.Equal(12, target.SpaceBefore);
        Assert.Equal(LineSpacingRule.OnePointFive, target.LineSpacingRule);
        Assert.True(target.KeepWithNext);
        Assert.False(target.WordWrap);
        Assert.Equal(ParagraphAlignment.Right, target.Alignment);
        Assert.Equal(0, Paragraph(1).Format.LeftIndent);
    }

    [Fact]
    public void ParagraphForwardsTheRestOfTheFormat()
    {
        ParagraphWrapper paragraph = Paragraph(0);

        paragraph.OutlineLevel = ParagraphOutlineLevel.Level2;
        paragraph.NoLineNumber = true;
        paragraph.MirrorIndents = true;
        paragraph.Hyphenation = false;
        paragraph.SpaceBeforeAuto = true;
        paragraph.SpaceAfterAuto = true;
        paragraph.ReadingOrder = ReadingOrder.RightToLeft;
        paragraph.TabStops.Add(72);

        Assert.Equal(ParagraphOutlineLevel.Level2, paragraph.Format.OutlineLevel);
        Assert.True(paragraph.Format.NoLineNumber);
        Assert.True(paragraph.Format.MirrorIndents);
        Assert.False(paragraph.Format.Hyphenation);
        Assert.True(paragraph.Format.SpaceBeforeAuto);
        Assert.True(paragraph.Format.SpaceAfterAuto);
        Assert.Equal(ReadingOrder.RightToLeft, paragraph.Format.ReadingOrder);
        Assert.Equal(1, paragraph.TabStops.Count);

        paragraph.Reset();

        Assert.Equal(ParagraphOutlineLevel.BodyText, paragraph.OutlineLevel);
        Assert.Equal(ReadingOrder.LeftToRight, paragraph.ReadingOrder);
        Assert.False(paragraph.NoLineNumber);
    }
}
