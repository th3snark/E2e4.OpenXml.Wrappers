using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The borders and shading of styles. The markup is what Word 16 wrote (the style-borders-* and style-shading-* scenarios of
/// MarkupConformanceTest).
/// </summary>
public class StyleBordersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public StyleBordersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void AParagraphStyleHasTheFourOutsideBorders()
    {
        StyleWrapper style = _document.Styles.Add("Boxed", StyleType.Paragraph);

        style.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;

        ParagraphBorders borders = style.Style.StyleParagraphProperties.ParagraphBorders;
        Assert.Equal(4, borders.ChildElements.Count);
        Assert.Equal(BorderLineStyle.SingleLine, style.Borders[BorderPosition.Left].LineStyle);
        Assert.Equal(4u, borders.GetFirstChild<LeftBorder>().Space.Value);
        Assert.Equal(4, style.Borders.Count);
        Assert.Throws<NotSupportedException>(() => style.Borders[BorderPosition.Horizontal]);
    }

    [Fact]
    public void ParagraphFormatOfAStyleHasTheSameBorders()
    {
        StyleWrapper style = _document.Styles.Add("Boxed", StyleType.Paragraph);

        style.ParagraphFormat.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(BorderLineStyle.SingleLine, style.Borders[BorderPosition.Top].LineStyle);
    }

    [Fact]
    public void BordersCanBeClearedFromAStyle()
    {
        StyleWrapper style = _document.Styles.Add("Boxed", StyleType.Paragraph);
        style.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;

        style.Borders.Clear();

        Assert.Null(style.Style.StyleParagraphProperties?.ParagraphBorders);
    }

    [Fact]
    public void AParagraphStylesShadingIsGivenToItsLinkedStyleToo()
    {
        StyleWrapper style = _document.Styles.Add("Shaded", StyleType.Paragraph);

        style.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);

        Assert.Equal(ColorHelper.RGB(255, 0, 0), (int)style.Shading.BackgroundPatternColor.Value.Value);
        Assert.NotNull(style.Style.StyleParagraphProperties.Shading);
        Assert.NotNull(style.LinkStyle.Style.StyleRunProperties.Shading);
    }

    [Fact]
    public void TheBorderOfTheFontOfAParagraphStyleIsGivenToItsLinkedStyleToo_ButItsShadingIsNot()
    {
        StyleWrapper style = _document.Styles.Add("Text", StyleType.Paragraph);

        style.Font.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
        style.Font.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);

        Assert.NotNull(style.Style.StyleRunProperties.Border);
        Assert.NotNull(style.LinkStyle.Style.StyleRunProperties.Border);
        Assert.NotNull(style.Style.StyleRunProperties.Shading);
        Assert.Null(style.LinkStyle.Style.StyleRunProperties.Shading);
        Assert.Equal(1, style.Font.Borders.Count);
    }

    [Fact]
    public void ACharacterStyleHasNoBordersOrShadingOfItsOwn_ButItsFontHas()
    {
        StyleWrapper style = _document.Styles.Add("Char", StyleType.Character);

        Assert.Null(style.Borders);
        Assert.Null(style.Shading);
        style.Font.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
        Assert.Equal(BorderLineStyle.SingleLine, style.Font.Borders[BorderPosition.Top].LineStyle);
    }
}
