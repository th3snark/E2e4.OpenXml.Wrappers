using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperParagraphFormatTest
{
    private static Paragraph CreateParagraph(string text = "a")
    {
        return new Paragraph(new Run(new Text(text)));
    }

    private static string[] ChildNames(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        return element.ChildElements.Select(e => e.LocalName).ToArray();
    }

    [Theory]
    [InlineData(ParagraphAlignment.Left)]
    [InlineData(ParagraphAlignment.Center)]
    [InlineData(ParagraphAlignment.Right)]
    [InlineData(ParagraphAlignment.Justify)]
    public void AlignmentRoundTrips(ParagraphAlignment alignment)
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        format.Alignment = alignment;

        Assert.Equal(alignment, format.Alignment);
    }

    [Fact]
    public void AlignmentValuesMatchWord()
    {
        Assert.Equal(0, (int)ParagraphAlignment.Left);
        Assert.Equal(1, (int)ParagraphAlignment.Center);
        Assert.Equal(2, (int)ParagraphAlignment.Right);
        Assert.Equal(3, (int)ParagraphAlignment.Justify);
        Assert.Equal(0, (int)LineSpacingRule.SingleSpacing);
        Assert.Equal(1, (int)LineSpacingRule.OnePointFive);
        Assert.Equal(2, (int)LineSpacingRule.DoubleSpacing);
        Assert.Equal(3, (int)LineSpacingRule.AtLeast);
        Assert.Equal(4, (int)LineSpacingRule.Exactly);
        Assert.Equal(5, (int)LineSpacingRule.Multiple);
    }

    [Fact]
    public void AlignmentWritesTheMatchingJustification()
    {
        Paragraph paragraph = CreateParagraph();
        new RangeWrapper(paragraph).ParagraphFormat.Alignment = ParagraphAlignment.Justify;

        Assert.Equal(JustificationValues.Both, paragraph.ParagraphProperties.Justification.Val.Value);
    }

    [Fact]
    public void AlignmentDefaultsToLeftAndNullRemovesTheFormatting()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        Assert.Equal(ParagraphAlignment.Left, format.Alignment);

        format.Alignment = ParagraphAlignment.Center;
        format.Alignment = null;

        Assert.Null(paragraph.ParagraphProperties);
        Assert.Equal(ParagraphAlignment.Left, format.Alignment);
    }

    [Fact]
    public void AlignmentRejectsAnUndefinedValue()
    {
        ParagraphFormatWrapper format = new RangeWrapper(CreateParagraph()).ParagraphFormat;

        Assert.Throws<ArgumentOutOfRangeException>(() => format.Alignment = (ParagraphAlignment)99);
    }

    [Fact]
    public void RightAndEndAlignmentBothReadAsRight()
    {
        Paragraph start = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.End }));

        Assert.Equal(ParagraphAlignment.Right, new RangeWrapper(start).ParagraphFormat.Alignment);
    }

    [Fact]
    public void IndentsAreInPointsAndStoredInTwips()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        Assert.Equal(0, format.LeftIndent);

        format.LeftIndent = 36;
        format.RightIndent = 10.5;

        Assert.Equal(36, format.LeftIndent);
        Assert.Equal(10.5, format.RightIndent);
        Assert.Equal("720", paragraph.ParagraphProperties.Indentation.Left.Value);
        Assert.Equal("210", paragraph.ParagraphProperties.Indentation.Right.Value);
    }

    [Fact]
    public void ANegativeIndentIsAllowedAndClearingTheLastAttributeRemovesTheElement()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        format.LeftIndent = -18;
        Assert.Equal(-18, format.LeftIndent);

        format.LeftIndent = null;
        Assert.Null(paragraph.ParagraphProperties);
    }

    [Fact]
    public void SettingTheLeftIndentReplacesAStartIndent()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new Indentation() { Start = "100" }));
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        Assert.Equal(5, format.LeftIndent);

        format.LeftIndent = 20;

        Assert.Null(paragraph.ParagraphProperties.Indentation.Start);
        Assert.Equal(20, format.LeftIndent);
    }

    [Fact]
    public void AFirstLineIndentIsPositiveAndAHangingIndentIsNegative()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;

        format.FirstLineIndent = 18;
        Assert.Equal(18, format.FirstLineIndent);
        Assert.Equal("360", paragraph.ParagraphProperties.Indentation.FirstLine.Value);
        Assert.Null(paragraph.ParagraphProperties.Indentation.Hanging);

        format.FirstLineIndent = -18;
        Assert.Equal(-18, format.FirstLineIndent);
        Assert.Equal("360", paragraph.ParagraphProperties.Indentation.Hanging.Value);
        Assert.Null(paragraph.ParagraphProperties.Indentation.FirstLine);

        format.FirstLineIndent = 0;
        Assert.Equal(0, format.FirstLineIndent);
    }

    [Fact]
    public void SpacingIsInPointsAndTheOtherFormsOfSpacingAreCleared()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new SpacingBetweenLines() { BeforeAutoSpacing = true, AfterLines = 100 }));
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        format.SpaceBefore = 6;
        format.SpaceAfter = 12.5;

        Assert.Equal(6, format.SpaceBefore);
        Assert.Equal(12.5, format.SpaceAfter);
        SpacingBetweenLines spacing = paragraph.ParagraphProperties.SpacingBetweenLines;
        Assert.Equal("120", spacing.Before.Value);
        Assert.Equal("250", spacing.After.Value);
        Assert.Null(spacing.BeforeAutoSpacing);
        Assert.Null(spacing.AfterLines);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(1585)]
    public void SpacingOutOfRangeIsRejected(double points)
    {
        ParagraphFormatWrapper format = new RangeWrapper(CreateParagraph()).ParagraphFormat;

        Assert.Throws<ArgumentOutOfRangeException>(() => format.SpaceBefore = points);
        Assert.Throws<ArgumentOutOfRangeException>(() => format.SpaceAfter = points);
    }

    [Fact]
    public void ParagraphWrapperSpacingAndTheFormatAgree()
    {
        ParagraphWrapper wrapper = new ParagraphWrapper(CreateParagraph());
        wrapper.SpaceBefore = 12;

        Assert.Equal(12, wrapper.Format.SpaceBefore);
        Assert.Equal("240", wrapper.Paragraph.ParagraphProperties.SpacingBetweenLines.Before.Value);
    }

    [Fact]
    public void TheLineSpacingRuleSetsTheSpacingOfTheMatchingNumberOfLines()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        Assert.Equal(LineSpacingRule.SingleSpacing, format.LineSpacingRule);
        Assert.Equal(12, format.LineSpacing);

        format.LineSpacingRule = LineSpacingRule.OnePointFive;
        Assert.Equal(LineSpacingRule.OnePointFive, format.LineSpacingRule);
        Assert.Equal(18, format.LineSpacing);

        format.LineSpacingRule = LineSpacingRule.DoubleSpacing;
        Assert.Equal(LineSpacingRule.DoubleSpacing, format.LineSpacingRule);
        Assert.Equal(24, format.LineSpacing);
        Assert.Equal("480", paragraph.ParagraphProperties.SpacingBetweenLines.Line.Value);
        Assert.Equal(LineSpacingRuleValues.Auto, paragraph.ParagraphProperties.SpacingBetweenLines.LineRule.Value);

        format.LineSpacingRule = LineSpacingRule.SingleSpacing;
        Assert.Equal("240", paragraph.ParagraphProperties.SpacingBetweenLines.Line.Value);
    }

    [Fact]
    public void SettingTheLineSpacingMakesItAMultipleUnlessTheRuleIsExactOrAtLeast()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;

        format.LineSpacing = 15;
        Assert.Equal(LineSpacingRule.Multiple, format.LineSpacingRule);
        Assert.Equal(15, format.LineSpacing);
        Assert.Equal("300", paragraph.ParagraphProperties.SpacingBetweenLines.Line.Value);

        format.LineSpacingRule = LineSpacingRule.Exactly;
        Assert.Equal(LineSpacingRuleValues.Exact, paragraph.ParagraphProperties.SpacingBetweenLines.LineRule.Value);
        Assert.Equal(15, format.LineSpacing);

        format.LineSpacing = 20;
        Assert.Equal(LineSpacingRule.Exactly, format.LineSpacingRule);
        Assert.Equal(20, format.LineSpacing);

        format.LineSpacingRule = LineSpacingRule.AtLeast;
        Assert.Equal(LineSpacingRule.AtLeast, format.LineSpacingRule);
    }

    [Fact]
    public void ChangingFromExactBackToAMultipleResetsToSingleSpacing()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        format.LineSpacingRule = LineSpacingRule.Exactly;
        format.LineSpacing = 30;

        format.LineSpacingRule = LineSpacingRule.Multiple;

        Assert.Equal(LineSpacingRule.SingleSpacing, format.LineSpacingRule);
    }

    [Fact]
    public void NullLineSpacingRemovesTheLineAndRule()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        format.LineSpacing = 18;
        format.LineSpacing = null;

        Assert.Null(paragraph.ParagraphProperties);
    }

    [Fact]
    public void LineSpacingMustBeGreaterThanZero()
    {
        ParagraphFormatWrapper format = new RangeWrapper(CreateParagraph()).ParagraphFormat;

        Assert.Throws<ArgumentOutOfRangeException>(() => format.LineSpacing = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => format.LineSpacing = 2000);
    }

    [Fact]
    public void TheOnOffPropertiesRoundTrip()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        Assert.False(format.KeepWithNext);

        format.KeepWithNext = true;
        format.KeepTogether = true;
        format.PageBreakBefore = true;
        format.WidowControl = true;

        Assert.True(format.KeepWithNext);
        Assert.True(format.KeepTogether);
        Assert.True(format.PageBreakBefore);
        Assert.True(format.WidowControl);

        format.KeepWithNext = false;
        Assert.False(format.KeepWithNext);
        Assert.False(paragraph.ParagraphProperties.KeepNext.Val.Value);

        format.KeepWithNext = null;
        Assert.Null(paragraph.ParagraphProperties.KeepNext);
    }

    [Fact]
    public void PropertiesAreWrittenInSchemaOrder()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" }, new ParagraphMarkRunProperties()), CreateParagraph().FirstChild.CloneNode(true));
        ParagraphFormatWrapper format = new RangeWrapper(paragraph).ParagraphFormat;
        format.Alignment = ParagraphAlignment.Center;
        format.SpaceAfter = 6;
        format.LeftIndent = 10;
        format.WidowControl = true;
        format.KeepWithNext = true;

        Assert.Equal(new[] { "pStyle", "keepNext", "widowControl", "spacing", "ind", "jc", "rPr" }, ChildNames(paragraph.ParagraphProperties));
    }

    [Fact]
    public void ARangeOfSeveralParagraphsFormatsAllOfThem()
    {
        Paragraph first = CreateParagraph();
        Paragraph second = CreateParagraph();
        Paragraph outside = CreateParagraph();
        Body body = new Body(new Table(new TableRow(new TableCell(first, second))), outside);

        new RangeWrapper(first.Parent).ParagraphFormat.Alignment = ParagraphAlignment.Right;

        Assert.Equal(ParagraphAlignment.Right, new RangeWrapper(first).ParagraphFormat.Alignment);
        Assert.Equal(ParagraphAlignment.Right, new RangeWrapper(second).ParagraphFormat.Alignment);
        Assert.Null(outside.ParagraphProperties);
    }

    [Fact]
    public void ThePropertiesAreNullWhenTheParagraphsDiffer()
    {
        Paragraph first = CreateParagraph();
        Paragraph second = CreateParagraph();
        TableCell cell = new TableCell(first, second);
        new RangeWrapper(first).ParagraphFormat.SpaceAfter = 6;
        RangeWrapper range = new RangeWrapper(cell);

        Assert.Null(range.ParagraphFormat.SpaceAfter);
        Assert.Equal(0, range.ParagraphFormat.SpaceBefore);
    }

    [Fact]
    public void ABookmarkRangeFormatsTheParagraphsItTouches()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph first = new Paragraph(new Run(new Text("a")), start, new Run(new Text("b")));
        Paragraph second = new Paragraph(new Run(new Text("c")), end);
        Paragraph outside = CreateParagraph();
        Body body = new Body(first, second, outside);

        new RangeWrapper(start, end).ParagraphFormat.PageBreakBefore = true;

        Assert.NotNull(first.ParagraphProperties.PageBreakBefore);
        Assert.NotNull(second.ParagraphProperties.PageBreakBefore);
        Assert.Null(outside.ParagraphProperties);
    }

    [Fact]
    public void ParagraphWrapperFormatIsTheFormatOfItsRange()
    {
        Paragraph paragraph = CreateParagraph();
        ParagraphWrapper wrapper = new ParagraphWrapper(paragraph);
        wrapper.Format.Alignment = ParagraphAlignment.Center;

        Assert.Equal(ParagraphAlignment.Center, wrapper.Range.ParagraphFormat.Alignment);
    }

    [Fact]
    public void AnEmptyBookmarkRangeIsFormattedThroughItsParagraph()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Body body = new Body(new Paragraph(start, end));

        ParagraphFormatWrapper format = new RangeWrapper(start, end).ParagraphFormat;
        format.Alignment = ParagraphAlignment.Center;
        Assert.Equal(ParagraphAlignment.Center, format.Alignment);
    }
}
