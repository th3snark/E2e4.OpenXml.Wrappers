using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The complex script and East Asian members of <see cref="FontWrapper"/>, the fonts by script, Grow and Shrink, and Duplicate. The sizes
/// Grow and Shrink go through are those Word 16 gave, and the markup is what it saved (see the font-* scenarios of MarkupConformanceTest).
/// </summary>
public class FontWrapperScriptsTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public FontWrapperScriptsTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "Sample text here";
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

    [Theory]
    [InlineData(1, 2, 1)]
    [InlineData(3, 4, 2)]
    [InlineData(4.5, 5.5, 3.5)]
    [InlineData(5, 6, 4)]
    [InlineData(8, 9, 7)]
    [InlineData(8.5, 9, 8)]
    [InlineData(9.5, 10, 9)]
    [InlineData(10, 11, 9)]
    [InlineData(10.5, 11, 10)]
    [InlineData(12.5, 14, 12)]
    [InlineData(27, 28, 26)]
    [InlineData(30, 36, 28)]
    [InlineData(36, 48, 28)]
    [InlineData(40, 48, 36)]
    [InlineData(72, 80, 48)]
    [InlineData(75, 80, 72)]
    [InlineData(80, 90, 72)]
    [InlineData(85, 90, 80)]
    [InlineData(1630, 1638, 1620)]
    [InlineData(1638, 1638, 1630)]
    public void GrowAndShrink_GoThroughTheSizesWordGoesThrough(double size, double grown, double shrunk)
    {
        FontWrapper font = _document.Range(0, 6).Font;

        font.Size = size;
        font.Grow();
        Assert.Equal(grown, font.Size);

        font.Size = size;
        font.Shrink();
        Assert.Equal(shrunk, font.Size);
    }

    [Fact]
    public void GrowAndShrink_StepEachRunFromItsOwnSize_AndMoveTheComplexScriptSize()
    {
        _document.Range(0, 3).Font.Size = 10;
        _document.Range(3, 6).Font.Size = 20;
        FontWrapper font = _document.Range(0, 6).Font;
        Assert.Null(font.Size);

        font.Grow();

        Assert.Equal(11, _document.Range(0, 3).Font.Size);
        Assert.Equal(22, _document.Range(3, 6).Font.Size);
        Assert.Equal(11, _document.Range(0, 3).Font.SizeBi);
        Assert.Equal(22, _document.Range(3, 6).Font.SizeBi);
    }

    [Fact]
    public void GrowBackToTheSizeOfTheStyle_LeavesNoDirectSize()
    {
        FontWrapper font = _document.Range(0, 6).Font;
        double normal = font.Size.Value;

        font.Shrink();
        font.Grow();

        Assert.Equal(normal, font.Size);
        Assert.Empty(Body.Descendants<FontSize>());
        Assert.Empty(Body.Descendants<FontSizeComplexScript>());
    }

    [Fact]
    public void ComplexScriptProperties_ReadAndWrite()
    {
        FontWrapper font = _document.Range(0, 6).Font;
        Assert.False(font.BoldBi);
        Assert.False(font.ItalicBi);

        font.BoldBi = true;
        font.ItalicBi = true;
        font.SizeBi = 14;

        Assert.True(font.BoldBi);
        Assert.True(font.ItalicBi);
        Assert.Equal(14, font.SizeBi);
        Assert.False(font.Bold);
        Assert.NotEqual(14, font.Size);
        Assert.NotEmpty(Body.Descendants<BoldComplexScript>());

        font.BoldBi = false;
        font.SizeBi = null;
        Assert.False(font.BoldBi);
        Assert.Empty(Body.Descendants<BoldComplexScript>());
        Assert.Empty(Body.Descendants<FontSizeComplexScript>());
    }

    [Fact]
    public void SizeBi_RejectsASizeWordDoesNotAllow()
    {
        FontWrapper font = _document.Range(0, 6).Font;

        Assert.Throws<ArgumentOutOfRangeException>(() => font.SizeBi = 0);
        Assert.Throws<ArgumentOutOfRangeException>(() => font.SizeBi = 1639);
    }

    [Fact]
    public void NamesByScript_AreSetOnTheirOwn()
    {
        FontWrapper font = _document.Range(0, 6).Font;

        font.NameAscii = "Arial";
        font.NameOther = "Verdana";
        font.NameFarEast = "MS Gothic";
        font.NameBi = "Tahoma";

        Assert.Equal("Arial", font.NameAscii);
        Assert.Equal("Verdana", font.NameOther);
        Assert.Equal("MS Gothic", font.NameFarEast);
        Assert.Equal("Tahoma", font.NameBi);
        Assert.Equal("Arial", font.Name);
        Assert.Equal("Tahoma", font.Direct.NameBi);
    }

    [Fact]
    public void NameAscii_Null_RemovesOnlyThatFont()
    {
        FontWrapper font = _document.Range(0, 6).Font;
        font.NameAscii = "Arial";
        font.NameBi = "Tahoma";

        font.NameAscii = null;

        Assert.Null(font.Direct.NameAscii);
        Assert.Equal("Tahoma", font.NameBi);
    }

    [Fact]
    public void SettingName_ChangesTheComplexScriptFontOnlyIfItWasSetOnItsOwn()
    {
        FontWrapper font = _document.Range(0, 4).Font;
        font.Name = "Arial";
        Assert.Null(font.Direct.NameBi);

        font.NameBi = "Tahoma";
        font.NameFarEast = "MS Gothic";
        font.Name = "Calibri";

        Assert.Equal("Calibri", font.NameBi);
        Assert.Equal("MS Gothic", font.NameFarEast);
    }

    [Fact]
    public void NamesThatComeFromTheTheme_ReadAsTheirPartOfTheTheme()
    {
        FontWrapper font = _document.Range(0, 6).Font;

        Assert.StartsWith("+", font.NameAscii, StringComparison.Ordinal);
        Assert.StartsWith("+", font.NameOther, StringComparison.Ordinal);
        Assert.EndsWith(" CS", font.NameBi, StringComparison.Ordinal);
        Assert.Null(font.Direct.NameAscii);
    }

    [Theory]
    [InlineData(EmphasisMark.OverSolid, "dot")]
    [InlineData(EmphasisMark.OverComma, "comma")]
    [InlineData(EmphasisMark.OverWhiteCircle, "circle")]
    [InlineData(EmphasisMark.UnderSolid, "underDot")]
    public void EmphasisMark_IsWrittenAsWordWritesIt(EmphasisMark mark, string expected)
    {
        FontWrapper font = _document.Range(0, 6).Font;
        Assert.Equal(EmphasisMark.None, font.EmphasisMark);

        font.EmphasisMark = mark;

        Assert.Equal(mark, font.EmphasisMark);
        Assert.Equal(expected, Body.Descendants<Emphasis>().Single().Val.InnerText);

        font.EmphasisMark = EmphasisMark.None;
        Assert.Empty(Body.Descendants<Emphasis>());
    }

    [Fact]
    public void EmphasisMark_RejectsAValueThatIsNotAMark()
    {
        FontWrapper font = _document.Range(0, 6).Font;

        Assert.Throws<ArgumentOutOfRangeException>(() => font.EmphasisMark = (EmphasisMark)9);
    }

    [Fact]
    public void DisableCharacterSpaceGrid_WritesSnapToGridOff()
    {
        FontWrapper font = _document.Range(0, 6).Font;
        Assert.False(font.DisableCharacterSpaceGrid);

        font.DisableCharacterSpaceGrid = true;

        Assert.True(font.DisableCharacterSpaceGrid);
        Assert.False(Body.Descendants<SnapToGrid>().Single().Val.Value);

        font.DisableCharacterSpaceGrid = false;
        Assert.False(font.DisableCharacterSpaceGrid);
        Assert.Empty(Body.Descendants<SnapToGrid>());
    }

    [Fact]
    public void ColorIndexBi_IsTheColorIndex()
    {
        FontWrapper font = _document.Range(0, 6).Font;

        font.ColorIndexBi = ColorIndex.Red;

        Assert.Equal(ColorIndex.Red, font.ColorIndex);
        Assert.Equal(ColorIndex.Red, font.ColorIndexBi);
    }

    [Fact]
    public void Duplicate_KeepsTheFormattingTheTextHasNow()
    {
        FontWrapper font = _document.Range(0, 6).Font;
        font.NameAscii = "Arial";
        font.SizeBi = 14;
        font.BoldBi = true;

        FontWrapper copy = font.Duplicate();
        font.SizeBi = 20;
        font.NameAscii = "Verdana";

        Assert.Equal("Arial", copy.NameAscii);
        Assert.Equal(14, copy.SizeBi);
        Assert.True(copy.BoldBi);
    }

    [Fact]
    public void Duplicate_OfMixedText_IsMixedToo()
    {
        _document.Range(0, 3).Font.Size = 10;
        _document.Range(3, 6).Font.Size = 20;

        FontWrapper copy = _document.Range(0, 6).Font.Duplicate();

        Assert.Null(copy.Size);
    }

    [Fact]
    public void AssigningAFont_GivesTheRangeItsFormatting()
    {
        FontWrapper source = _document.Range(0, 6).Font;
        source.NameAscii = "Arial";
        source.SizeBi = 14;
        source.BoldBi = true;
        source.Italic = true;
        source.EmphasisMark = EmphasisMark.OverComma;
        FontWrapper copy = source.Duplicate();

        RangeWrapper target = _document.Range(7, 11);
        target.Font = copy;

        Assert.Equal("Arial", target.Font.NameAscii);
        Assert.Equal(14, target.Font.SizeBi);
        Assert.True(target.Font.BoldBi);
        Assert.True(target.Font.Italic);
        Assert.Equal(EmphasisMark.OverComma, target.Font.EmphasisMark);
        Assert.False(_document.Range(12, 16).Font.Italic);
    }

    [Fact]
    public void AssigningAFont_LeavesPropertiesWithNoSingleValueAlone()
    {
        _document.Range(0, 3).Font.Size = 10;
        _document.Range(3, 6).Font.Size = 20;
        _document.Range(0, 6).Font.Bold = true;
        FontWrapper mixed = _document.Range(0, 6).Font.Duplicate();

        RangeWrapper target = _document.Range(7, 11);
        target.Font.Size = 16;
        target.Font = mixed;

        Assert.Equal(16, target.Font.Size);
        Assert.True(target.Font.Bold);
    }
}
