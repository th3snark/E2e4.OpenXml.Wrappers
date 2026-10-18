using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperFontTest
{
    private static Paragraph CreateParagraph(params string[] runTexts)
    {
        return new Paragraph(runTexts.Select(t => new Run(new Text(t))));
    }

    private static string[] ChildNames(DocumentFormat.OpenXml.OpenXmlElement element)
    {
        return element.ChildElements.Select(e => e.LocalName).ToArray();
    }

    [Fact]
    public void SettingBoldFormatsEveryRunAndTheParagraphMark()
    {
        Paragraph paragraph = CreateParagraph("a", "b");
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Font.Bold = true;

        Assert.All(paragraph.Elements<Run>(), run => Assert.Equal(new[] { "b", "bCs" }, ChildNames(run.RunProperties)));
        Assert.Equal(new[] { "b", "bCs" }, ChildNames(paragraph.ParagraphProperties.ParagraphMarkRunProperties));
        Assert.True(range.Font.Bold);
    }

    [Fact]
    public void SettingBoldToFalseSwitchesItOffExplicitly()
    {
        Paragraph paragraph = CreateParagraph("a");
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Font.Bold = true;
        range.Font.Bold = false;

        Assert.False(range.Font.Bold);
        Assert.False(paragraph.Elements<Run>().Single().RunProperties.Elements<Bold>().Single().Val.Value);
    }

    [Fact]
    public void SettingBoldToNullRemovesTheFormattingAndTidiesUp()
    {
        Paragraph paragraph = CreateParagraph("a");
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Font.Bold = true;
        range.Font.Bold = null;

        Assert.Null(paragraph.Elements<Run>().Single().RunProperties);
        Assert.Null(paragraph.ParagraphProperties.ParagraphMarkRunProperties);
        Assert.False(range.Font.Bold);
    }

    [Fact]
    public void AnExistingBoldElementWithNoValueReadsAsTrue()
    {
        Paragraph paragraph = new Paragraph(new Run(new RunProperties(new Bold()), new Text("a")));

        Assert.True(new RangeWrapper(paragraph).Font.Bold);
    }

    [Fact]
    public void ABooleanIsFalseWhenNothingIsSetAndNullWhenTheRunsDiffer()
    {
        Paragraph paragraph = CreateParagraph("a", "b");

        Assert.False(new RangeWrapper(paragraph).Font.Italic);

        new RangeWrapper(paragraph).Font.Italic = true;
        paragraph.Elements<Run>().Last().RunProperties.RemoveAllChildren();

        Assert.Null(new RangeWrapper(paragraph).Font.Italic);
    }

    [Fact]
    public void ItalicStrikeThroughAndUnderlineCanBeSetAndRead()
    {
        Paragraph paragraph = CreateParagraph("a");
        FontWrapper font = new RangeWrapper(paragraph).Font;
        font.Italic = true;
        font.StrikeThrough = true;
        font.Underline = UnderlineStyle.SingleLine;

        Assert.True(font.Italic);
        Assert.True(font.StrikeThrough);
        Assert.Equal(UnderlineStyle.SingleLine, font.Underline);
        Assert.Equal(UnderlineValues.Single, paragraph.Elements<Run>().Single().RunProperties.Underline.Val.Value);

        font.Underline = UnderlineStyle.None;
        Assert.Equal(UnderlineStyle.None, font.Underline);
        Assert.Equal(UnderlineValues.None, paragraph.Elements<Run>().Single().RunProperties.Underline.Val.Value);
    }

    [Fact]
    public void TheUnderlineStyleIsRead()
    {
        Paragraph paragraph = new Paragraph(new Run(new RunProperties(new Underline() { Val = UnderlineValues.Double }), new Text("a")));

        Assert.Equal(UnderlineStyle.DoubleLine, new RangeWrapper(paragraph).Font.Underline);
    }

    [Fact]
    public void SizeIsInPointsAndStoredInHalfPoints()
    {
        Paragraph paragraph = CreateParagraph("a");
        FontWrapper font = new RangeWrapper(paragraph).Font;
        Assert.Null(font.Direct.Size);
        Assert.Equal(10, font.Size);

        font.Size = 10.5;

        Assert.Equal(10.5, font.Size);
        RunProperties properties = paragraph.Elements<Run>().Single().RunProperties;
        Assert.Equal("21", properties.FontSize.Val.Value);
        Assert.Equal("21", properties.FontSizeComplexScript.Val.Value);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1639)]
    public void SizeOutOfRangeIsRejected(double size)
    {
        FontWrapper font = new RangeWrapper(CreateParagraph("a")).Font;

        Assert.Throws<ArgumentOutOfRangeException>(() => font.Size = size);
    }

    [Fact]
    public void SizeIsNullWhenTheRunsDiffer()
    {
        Paragraph paragraph = CreateParagraph("a", "b");
        new RangeWrapper(paragraph.Elements<Run>().First().Parent).Font.Size = 12;
        paragraph.Elements<Run>().Last().PrependChild(new RunProperties(new FontSize() { Val = "40" }));

        Assert.Null(new RangeWrapper(paragraph).Font.Size);
    }

    [Fact]
    public void NameSetsEveryFontSlot()
    {
        Paragraph paragraph = CreateParagraph("a");
        FontWrapper font = new RangeWrapper(paragraph).Font;
        font.Name = "Arial";

        Assert.Equal("Arial", font.Name);
        RunFonts fonts = paragraph.Elements<Run>().Single().RunProperties.RunFonts;
        Assert.Equal("Arial", fonts.Ascii.Value);
        Assert.Equal("Arial", fonts.HighAnsi.Value);
        Assert.Equal("Arial", fonts.EastAsia.Value);
        Assert.Equal("Arial", fonts.ComplexScript.Value);

        font.Name = null;
        Assert.Null(font.Direct.Name);
        Assert.Equal("Times New Roman", font.Name);
    }

    [Theory]
    [InlineData("FF0000")]
    [InlineData("00ff7F")]
    [InlineData("auto")]
    public void ColorAcceptsHexadecimalAndAuto(string color)
    {
        FontWrapper font = new RangeWrapper(CreateParagraph("a")).Font;
        font.Color = color;

        Assert.Equal(color, font.Color);
    }

    [Theory]
    [InlineData("red")]
    [InlineData("#FF0000")]
    [InlineData("FF00")]
    [InlineData("")]
    public void ColorRejectsAnythingElse(string color)
    {
        FontWrapper font = new RangeWrapper(CreateParagraph("a")).Font;

        Assert.Throws<ArgumentException>(() => font.Color = color);
    }

    [Fact]
    public void SettingSeveralPropertiesKeepsTheSchemaOrder()
    {
        Paragraph paragraph = new Paragraph(new Run(new RunProperties(new RunStyle() { Val = "Strong" }), new Text("a")));
        FontWrapper font = new RangeWrapper(paragraph).Font;
        font.Size = 12;
        font.Color = "FF0000";
        font.Underline = UnderlineStyle.SingleLine;
        font.Bold = true;
        font.Name = "Arial";

        Assert.Equal(
            new[] { "rStyle", "rFonts", "b", "bCs", "color", "sz", "szCs", "u" },
            ChildNames(paragraph.Elements<Run>().Single().RunProperties));
    }

    [Fact]
    public void ParagraphMarkTrackedChangesStayBeforeTheFormatting()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphMarkRunProperties(new Inserted() { Id = "1", Author = "a" })), new Run(new Text("x")));
        new RangeWrapper(paragraph).Font.Bold = true;

        Assert.Equal(new[] { "ins", "b", "bCs" }, ChildNames(paragraph.ParagraphProperties.ParagraphMarkRunProperties));
    }

    [Fact]
    public void ParagraphMarkPropertiesGoBeforeTheSectionProperties()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new SectionProperties()), new Run(new Text("x")));
        new RangeWrapper(paragraph).Font.Bold = true;

        Assert.Equal(new[] { "rPr", "sectPr" }, ChildNames(paragraph.ParagraphProperties));
    }

    [Fact]
    public void AnEmptyParagraphIsFormattedAndReadThroughItsParagraphMark()
    {
        Paragraph paragraph = new Paragraph();
        FontWrapper font = new RangeWrapper(paragraph).Font;
        font.Bold = true;

        Assert.True(font.Bold);
        Assert.Single(paragraph.ParagraphProperties.ParagraphMarkRunProperties.Elements<Bold>());
    }

    [Fact]
    public void RunsInsideHyperlinksAreIncluded()
    {
        Run inside = new Run(new Text("b"));
        Paragraph paragraph = new Paragraph(new Run(new Text("a")), new Hyperlink(inside));
        new RangeWrapper(paragraph).Font.Italic = true;

        Assert.NotNull(inside.RunProperties.Italic);
    }

    [Fact]
    public void ABookmarkRangeFormatsItsRunsButNotTheRestOrTheParagraphMark()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Run before = new Run(new Text("x"));
        Run inside = new Run(new Text("b"));
        Run after = new Run(new Text("y"));
        Paragraph paragraph = new Paragraph(before, start, inside, end, after);
        Body body = new Body(paragraph);

        new RangeWrapper(start, end).Font.Bold = true;

        Assert.NotNull(inside.RunProperties.Bold);
        Assert.Null(before.RunProperties);
        Assert.Null(after.RunProperties);
        Assert.Null(paragraph.ParagraphProperties);
    }

    [Fact]
    public void ATableCellFormatsAllOfItsParagraphs()
    {
        Paragraph first = CreateParagraph("a");
        Paragraph second = CreateParagraph("b");
        Paragraph outside = CreateParagraph("c");
        Body body = new Body(new Table(new TableRow(new TableCell(first, second))), outside);

        new RangeWrapper(first.Parent).Font.Color = "00FF00";

        Assert.Equal("00FF00", first.Elements<Run>().Single().RunProperties.Color.Val.Value);
        Assert.Equal("00FF00", second.Elements<Run>().Single().RunProperties.Color.Val.Value);
        Assert.Equal("00FF00", second.ParagraphProperties.ParagraphMarkRunProperties.Elements<Color>().Single().Val.Value);
        Assert.Null(outside.Elements<Run>().Single().RunProperties);
    }

    [Fact]
    public void TheDocumentFormatsEverythingIncludingTables()
    {
        Document document = new Document(new Body(CreateParagraph("a"), new Table(new TableRow(new TableCell(CreateParagraph("b")))), new SectionProperties()));
        FontWrapper font = new RangeWrapper(document).Font;
        font.Size = 14;

        Assert.All(document.Descendants<Run>(), run => Assert.Equal("28", run.RunProperties.FontSize.Val.Value));
        Assert.Equal(14, font.Size);
    }

    [Fact]
    public void AnEmptyRangeHasNoFont()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Body body = new Body(new Paragraph(start, end));
        FontWrapper font = new RangeWrapper(start, end).Font;

        Assert.Null(font.Bold);
        Assert.Null(font.Size);
        Assert.Null(font.Name);
        Assert.Null(font.Color);
        font.Bold = true;
    }

    [Fact]
    public void AFontOnARunRangeIsNotSupported()
    {
        Assert.Throws<NotSupportedException>(() => new RangeWrapper(new Run(new Text("a"))).Font.Bold = true);
    }
}
