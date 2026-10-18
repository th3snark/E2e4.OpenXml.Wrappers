using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class WordColorTest
{
    private static RangeWrapper CreateRange()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("text")));
        return new RangeWrapper(paragraph);
    }

    [Fact]
    public void RgbPacksTheComponentsAsVbaDoes()
    {
        Assert.Equal(255, ColorHelper.RGB(255, 0, 0));
        Assert.Equal(65280, ColorHelper.RGB(0, 255, 0));
        Assert.Equal(16711680, ColorHelper.RGB(0, 0, 255));
        Assert.Equal((int)NamedColor.Orange, ColorHelper.RGB(255, 102, 0));
    }

    [Theory]
    [InlineData(-1, 0, 0)]
    [InlineData(0, 256, 0)]
    [InlineData(0, 0, 300)]
    public void RgbRejectsComponentsOutOfRange(int red, int green, int blue)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ColorHelper.RGB(red, green, blue));
    }

    [Fact]
    public void AllTheWaysOfMakingARedGiveTheSameColour()
    {
        WordColor fromHex = "FF0000";
        WordColor fromRgb = ColorHelper.RGB(255, 0, 0);
        WordColor fromName = NamedColor.Red;

        Assert.Equal(fromHex, fromRgb);
        Assert.Equal(fromHex, fromName);
        Assert.Equal("FF0000", fromRgb.Hex);
        Assert.Equal(255, fromHex.Red);
        Assert.Equal(0, fromHex.Green);
        Assert.Equal(0, fromHex.Blue);
        Assert.Equal(255, fromHex.Value);
    }

    [Fact]
    public void HexIsReadInTheOrderRedGreenBlue()
    {
        WordColor color = "1F3864";

        Assert.Equal(0x1F, color.Red);
        Assert.Equal(0x38, color.Green);
        Assert.Equal(0x64, color.Blue);
        Assert.Equal(0x64 << 16 | 0x38 << 8 | 0x1F, color.Value);
        Assert.Equal("1F3864", color.ToString());
    }

    [Fact]
    public void TheAutomaticColourIsTheDefaultAndHasWordsValue()
    {
        Assert.True(default(WordColor).IsAutomatic);
        Assert.Equal(WordColor.Automatic, default(WordColor));
        Assert.Equal("auto", WordColor.Automatic.Hex);
        Assert.Equal(-16777216, WordColor.Automatic.Value);
        Assert.Equal(WordColor.Automatic, new WordColor("AUTO"));
        Assert.Equal(WordColor.Automatic, WordColor.FromValue(-16777216));
        Assert.Equal(WordColor.Automatic, (WordColor)NamedColor.Automatic);
        Assert.NotEqual(WordColor.Automatic, (WordColor)"000000");
    }

    [Theory]
    [InlineData("")]
    [InlineData("red")]
    [InlineData("FF00")]
    [InlineData("FF00000")]
    [InlineData("GG0000")]
    [InlineData("-F0000")]
    public void HexRejectsAnythingElse(string text)
    {
        Assert.Throws<ArgumentException>(() => new WordColor(text));
    }

    [Fact]
    public void ValuesOutOfRangeAreRejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => WordColor.FromValue(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => WordColor.FromValue(0x1000000));
    }

    [Fact]
    public void ColorIndexRoundTripsForAllSixteenColours()
    {
        for (int index = (int)ColorIndex.Black; index <= (int)ColorIndex.Gray25; index++)
        {
            WordColor color = WordColor.FromColorIndex((ColorIndex)index);
            Assert.Equal((ColorIndex)index, color.ToColorIndex());
        }
        Assert.Equal(ColorIndex.Auto, WordColor.Automatic.ToColorIndex());
        Assert.Throws<ArgumentOutOfRangeException>(() => WordColor.FromColorIndex(ColorIndex.ByAuthor));
    }

    [Fact]
    public void ColorIndexGivesTheNearestColour()
    {
        Assert.Equal(ColorIndex.Red, ((WordColor)"F01010").ToColorIndex());
        Assert.Equal(ColorIndex.DarkBlue, ((WordColor)"000070").ToColorIndex());
    }

    [Fact]
    public void FontColorTakesEveryKindOfValue()
    {
        RangeWrapper range = CreateRange();

        range.Font.Color = ColorHelper.RGB(255, 0, 0);
        Assert.Equal(NamedColor.Red, (NamedColor)range.Font.Color.Value.Value);
        Assert.Equal("FF0000", range.StartElement.Elements<Run>().First().RunProperties.Color.Val.Value);

        range.Font.Color = NamedColor.Blue;
        Assert.Equal("0000FF", range.StartElement.Elements<Run>().First().RunProperties.Color.Val.Value);

        range.Font.Color = WordColor.Automatic;
        Assert.Equal("auto", range.StartElement.Elements<Run>().First().RunProperties.Color.Val.Value);

        range.Font.Color = null;
        Assert.Null(range.StartElement.Elements<Run>().First().RunProperties?.Color);
    }

    [Fact]
    public void FontColorIndexSetsAndReadsTheNearestColour()
    {
        RangeWrapper range = CreateRange();

        range.Font.ColorIndex = ColorIndex.Red;
        Assert.Equal("FF0000", range.StartElement.Elements<Run>().First().RunProperties.Color.Val.Value);
        Assert.Equal(ColorIndex.Red, range.Font.ColorIndex);

        range.Font.ColorIndex = ColorIndex.Auto;
        Assert.Equal(ColorIndex.Auto, range.Font.ColorIndex);
    }

    [Fact]
    public void FontUnderlineColorTakesAColour()
    {
        RangeWrapper range = CreateRange();

        range.Font.UnderlineColor = NamedColor.Green;

        Assert.Equal(new WordColor("008000"), range.Font.UnderlineColor);
        Assert.Equal("008000", range.StartElement.Elements<Run>().First().RunProperties.Underline.Color.Value);
    }

    [Fact]
    public void FontColorReadsNullWhenTheTextDiffers()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("one")), new Run(new Text("two")));
        RangeWrapper range = new RangeWrapper(paragraph);
        range.Font.Color = "FF0000";
        ((Run)paragraph.Elements<Run>().Last()).RunProperties.Color.Val = "00FF00";

        Assert.Null(range.Font.Color);
        Assert.Null(range.Font.ColorIndex);
    }

    [Fact]
    public void BorderColorIndexAndColourAgree()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        BorderWrapper border = new TableWrapper(table).Borders[BorderPosition.Top];

        border.ColorIndex = ColorIndex.Blue;

        Assert.Equal(new WordColor("0000FF"), border.Color);
        Assert.Equal(ColorIndex.Blue, border.ColorIndex);
        Assert.Throws<ArgumentNullException>(() => border.ColorIndex = null);
    }

    [Fact]
    public void BordersSetOutsideTakesAnRgbValue()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        BordersWrapper borders = new TableWrapper(table).Borders;

        borders.SetOutside(BorderLineStyle.SingleLine, 1, ColorHelper.RGB(0, 128, 0));

        Assert.Equal(new WordColor("008000"), borders.OutsideColor);
        Assert.Equal(new WordColor("008000"), borders[BorderPosition.Left].Color);
    }

    [Fact]
    public void ShadingColoursAndIndexesAgree()
    {
        TableCell cell = new TableCell(new Paragraph());
        Assert.NotNull(new Table(new TableRow(cell)));
        ShadingWrapper shading = new TableCellWrapper(cell).Shading;

        shading.BackgroundPatternColorIndex = ColorIndex.Yellow;
        shading.ForegroundPatternColor = NamedColor.Black;

        Assert.Equal(new WordColor("FFFF00"), shading.BackgroundPatternColor);
        Assert.Equal(ColorIndex.Yellow, shading.BackgroundPatternColorIndex);
        Assert.Equal(ColorIndex.Black, shading.ForegroundPatternColorIndex);
        Assert.Equal("FFFF00", cell.TableCellProperties.Shading.Fill.Value);
    }
}
