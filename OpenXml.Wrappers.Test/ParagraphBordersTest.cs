using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The borders and shading of paragraphs and of characters. The default widths and distances, and the markup, are those Word 16 gave
/// (the paragraph-borders-shading and font-border-shading scenarios of MarkupConformanceTest).
/// </summary>
public class ParagraphBordersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ParagraphBordersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "One\rTwo\rThree\rFour text";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private Body Body
    {
        get { return _document.WordprocessingDocument.MainDocumentPart.Document.Body; }
    }

    private ParagraphWrapper Paragraph(int index)
    {
        return _document.Content.Paragraphs[index];
    }

    [Fact]
    public void ANewParagraph_HasFourBordersWithoutLines()
    {
        BordersWrapper borders = Paragraph(0).Borders;

        Assert.Equal(4, borders.Count);
        Assert.False(borders.Enable);
        Assert.Equal(BorderLineStyle.None, borders.OutsideLineStyle);
        Assert.Equal(BorderLineStyle.None, borders[BorderPosition.Top].LineStyle);
        Assert.Empty(Body.Descendants<ParagraphBorders>());
    }

    [Fact]
    public void ASetBorder_HasWordsDefaults()
    {
        BordersWrapper borders = Paragraph(0).Borders;

        borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        BorderWrapper top = borders[BorderPosition.Top];
        Assert.True(top.Visible);
        Assert.Equal(0.5, top.LineWidth.Points);
        Assert.Equal(WordColor.Automatic, top.Color);
        TopBorder element = Body.Descendants<TopBorder>().Single();
        Assert.Equal(1u, element.Space.Value);
        Assert.Equal(BorderLineStyle.None, borders[BorderPosition.Left].LineStyle);
    }

    [Fact]
    public void TheSidesAreFourPointsFromTheText_AboveAndBelowOne()
    {
        BordersWrapper borders = Paragraph(0).Borders;

        borders.OutsideLineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(1u, Body.Descendants<TopBorder>().Single().Space.Value);
        Assert.Equal(4u, Body.Descendants<LeftBorder>().Single().Space.Value);
        Assert.Equal(1u, Body.Descendants<BottomBorder>().Single().Space.Value);
        Assert.Equal(4u, Body.Descendants<RightBorder>().Single().Space.Value);
    }

    [Fact]
    public void OutsideWidthAndColor_AreSetOnAllFour()
    {
        BordersWrapper borders = Paragraph(1).Borders;

        borders.OutsideLineStyle = BorderLineStyle.SingleLine;
        borders.OutsideLineWidth = 1.5;
        borders.OutsideColor = ColorHelper.RGB(255, 0, 0);

        Assert.Equal(BorderLineStyle.SingleLine, borders.OutsideLineStyle);
        Assert.Equal(1.5, borders.OutsideLineWidth.Points);
        Assert.Equal(ColorHelper.RGB(255, 0, 0), (int)borders.OutsideColor.Value.Value);
        Assert.True(borders.Enable);
        Assert.All(Body.Descendants<BorderType>(), b => Assert.Equal("FF0000", b.Color.Value));
    }

    [Fact]
    public void OneBorderOnly_MakesTheGroupMixed()
    {
        BordersWrapper borders = Paragraph(2).Borders;

        borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(BorderLineStyle.Other, borders.OutsideLineStyle);
        Assert.False(borders.Enable);
    }

    [Fact]
    public void BordersOfARange_GoToEveryParagraphItTouches()
    {
        RangeWrapper range = _document.Range(0, 8);

        range.ParagraphFormat.Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.DoubleLine;

        Assert.Equal(BorderLineStyle.DoubleLine, Paragraph(0).Borders[BorderPosition.Bottom].LineStyle);
        Assert.Equal(BorderLineStyle.DoubleLine, Paragraph(1).Borders[BorderPosition.Bottom].LineStyle);
        Assert.Equal(BorderLineStyle.None, Paragraph(2).Borders[BorderPosition.Bottom].LineStyle);
        Assert.Equal(BorderLineStyle.DoubleLine, range.ParagraphFormat.Borders[BorderPosition.Bottom].LineStyle);
    }

    [Fact]
    public void BordersOfParagraphsThatDiffer_ReadAsOther()
    {
        Paragraph(0).Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(BorderLineStyle.Other, _document.Range(0, 8).ParagraphFormat.Borders[BorderPosition.Top].LineStyle);
    }

    [Fact]
    public void Clear_RemovesTheBordersOfAllTheParagraphs()
    {
        _document.Range(0, 8).ParagraphFormat.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;

        _document.Range(0, 8).ParagraphFormat.Borders.Clear();

        Assert.Empty(Body.Descendants<ParagraphBorders>());
        Assert.Equal(BorderLineStyle.None, Paragraph(2).Borders[BorderPosition.Top].LineStyle);
    }

    [Theory]
    [InlineData(BorderPosition.Horizontal)]
    [InlineData(BorderPosition.Vertical)]
    [InlineData(BorderPosition.DiagonalDown)]
    public void AParagraphHasTheFourEdgesOnly(BorderPosition position)
    {
        Assert.Throws<NotSupportedException>(() => Paragraph(0).Borders[position]);
    }

    [Fact]
    public void ParagraphShading_IsWrittenAsWordWritesIt()
    {
        ShadingWrapper shading = Paragraph(3).Shading;
        Assert.Equal(ShadingTexture.None, shading.Texture);

        shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
        shading.Texture = ShadingTexture.Percent10;

        Shading element = Body.Descendants<Shading>().Single();
        Assert.Equal("FF0000", element.Fill.Value);
        Assert.Equal("auto", element.Color.Value);
        Assert.Equal(ShadingPatternValues.Percent10, element.Val.Value);
        Assert.Equal(ShadingTexture.Percent10, Paragraph(3).Shading.Texture);
        Assert.Equal(ShadingTexture.None, Paragraph(0).Shading.Texture);
    }

    [Fact]
    public void ParagraphShading_OfARange_GoesToEveryParagraph()
    {
        _document.Range(0, 8).ParagraphFormat.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);

        Assert.Equal(2, Body.Descendants<Shading>().Count());
        Assert.Equal(ColorHelper.RGB(0, 255, 0), (int)Paragraph(1).Shading.BackgroundPatternColor.Value.Value);

        Paragraph(0).Shading.Clear();
        Assert.Single(Body.Descendants<Shading>());
    }

    [Fact]
    public void ACharacterBorder_IsOneBorderAllRoundTheText()
    {
        BordersWrapper borders = _document.Range(0, 3).Font.Borders;

        Assert.Equal(1, borders.Count);
        borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Border element = Body.Descendants<Border>().Single();
        Assert.Equal(BorderValues.Single, element.Val.Value);
        Assert.Equal(4u, element.Size.Value);
        Assert.Equal(0u, element.Space.Value);
        Assert.Equal("auto", element.Color.Value);
        Assert.Equal(BorderLineStyle.SingleLine, _document.Range(0, 3).Font.Borders.OutsideLineStyle);
        Assert.Equal(BorderLineStyle.None, _document.Range(4, 7).Font.Borders.OutsideLineStyle);
    }

    [Theory]
    [InlineData(BorderPosition.Left)]
    [InlineData(BorderPosition.Bottom)]
    [InlineData(BorderPosition.Right)]
    [InlineData(BorderPosition.Horizontal)]
    public void ACharacterHasNoBorderButTheTop(BorderPosition position)
    {
        Assert.Throws<NotSupportedException>(() => _document.Range(0, 3).Font.Borders[position]);
    }

    [Fact]
    public void CharacterBorders_ApplyToEveryRunOfTheRange_AndClear()
    {
        _document.Range(0, 3).Font.Bold = true;
        RangeWrapper range = _document.Range(0, 7);

        range.Font.Borders.OutsideLineStyle = BorderLineStyle.DoubleLine;

        Assert.True(Body.Descendants<Border>().Count() >= 2);
        Assert.All(Body.Descendants<Border>(), b => Assert.Equal(BorderValues.Double, b.Val.Value));

        range.Font.Borders.Clear();
        Assert.Empty(Body.Descendants<Border>());
    }

    [Fact]
    public void CharacterShading_IsWrittenAsWordWritesIt()
    {
        _document.Range(4, 7).Font.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);

        Shading element = Body.Descendants<RunProperties>().Select(p => p.Shading).Single(s => s != null);
        Assert.Equal("00FF00", element.Fill.Value);
        Assert.Equal("auto", element.Color.Value);
        Assert.Equal(ShadingPatternValues.Clear, element.Val.Value);
    }
}
