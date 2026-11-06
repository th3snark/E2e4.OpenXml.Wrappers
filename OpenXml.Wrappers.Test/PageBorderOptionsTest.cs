using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The options of page borders, <c>TogglePortrait</c> and <c>SectionDirection</c>. Defaults, limits and markup are Word 16's (the
/// page-border-options and page-setup-direction scenarios of MarkupConformanceTest).
/// </summary>
public class PageBorderOptionsTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public PageBorderOptionsTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private BordersWrapper Borders
    {
        get { return _document.Sections[0].Borders; }
    }

    private PageBorders PageBorders
    {
        get { return _document.WordprocessingDocument.MainDocumentPart.Document.Body.Descendants<PageBorders>().Single(); }
    }

    [Fact]
    public void PageBordersHaveWordsDefaults()
    {
        Borders.Enable = true;

        Assert.True(Borders.AlwaysInFront);
        Assert.Equal(BorderDistanceFrom.FromPageEdge, Borders.DistanceFrom);
        Assert.True(Borders.EnableFirstPageInSection);
        Assert.True(Borders.EnableOtherPagesInSection);
        Assert.True(Borders.SurroundHeader);
        Assert.True(Borders.SurroundFooter);
        Assert.Null(PageBorders.ZOrder);
        Assert.Null(PageBorders.Display);
    }

    [Fact]
    public void TheOptionsAreWrittenAsWordWritesThem()
    {
        Borders.Enable = true;
        Borders.AlwaysInFront = false;
        Borders.DistanceFrom = BorderDistanceFrom.FromText;
        Borders.EnableFirstPageInSection = false;

        Assert.Equal(PageBorderZOrderValues.Back, PageBorders.ZOrder.Value);
        Assert.Null(PageBorders.OffsetFrom);
        Assert.Equal(PageBorderDisplayValues.NotFirstPage, PageBorders.Display.Value);
        Assert.False(Borders.AlwaysInFront);
        Assert.Equal(BorderDistanceFrom.FromText, Borders.DistanceFrom);
        Assert.False(Borders.EnableFirstPageInSection);
        Assert.True(Borders.EnableOtherPagesInSection);

        Borders.AlwaysInFront = true;
        Borders.DistanceFrom = BorderDistanceFrom.FromPageEdge;
        Borders.EnableFirstPageInSection = true;

        Assert.Null(PageBorders.ZOrder);
        Assert.Equal(PageBorderOffsetValues.Page, PageBorders.OffsetFrom.Value);
        Assert.Null(PageBorders.Display);
    }

    [Fact]
    public void ThePageBorderIsNeverOnNoPage()
    {
        Borders.Enable = true;

        Borders.EnableFirstPageInSection = false;
        Borders.EnableOtherPagesInSection = false;

        Assert.True(Borders.EnableFirstPageInSection);
        Assert.False(Borders.EnableOtherPagesInSection);
        Assert.Equal(PageBorderDisplayValues.FirstPage, PageBorders.Display.Value);

        Borders.EnableFirstPageInSection = false;

        Assert.False(Borders.EnableFirstPageInSection);
        Assert.True(Borders.EnableOtherPagesInSection);
    }

    [Fact]
    public void TheHeaderAndFooterSettingsAreKeptInTheSettingsOfTheDocument()
    {
        Borders.Enable = true;
        Borders.SurroundHeader = false;
        Borders.SurroundFooter = false;

        string[] names = _document.WordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.Select(e => e.LocalName).ToArray();
        Assert.Contains("bordersDoNotSurroundHeader", names);
        Assert.Contains("bordersDoNotSurroundFooter", names);
        Assert.False(Borders.SurroundHeader);
        Assert.False(Borders.SurroundFooter);

        Borders.SurroundHeader = true;
        Borders.SurroundFooter = true;

        names = _document.WordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.Select(e => e.LocalName).ToArray();
        Assert.DoesNotContain("bordersDoNotSurroundHeader", names);
        Assert.DoesNotContain("bordersDoNotSurroundFooter", names);
        Assert.True(Borders.SurroundHeader);
    }

    [Fact]
    public void TheOptionsExistOnlyForPageBorders()
    {
        _document.Content.InsertAfter("text");
        BordersWrapper paragraphBorders = _document.Paragraphs[0].Borders;

        Assert.Throws<NotSupportedException>(() => paragraphBorders.AlwaysInFront);
        Assert.Throws<NotSupportedException>(() => paragraphBorders.DistanceFrom = BorderDistanceFrom.FromText);
        Assert.Throws<ArgumentOutOfRangeException>(() => Borders.DistanceFrom = (BorderDistanceFrom)5);
    }

    [Fact]
    public void TogglePortraitSwapsTheSizeAndTheOrientation()
    {
        double width = _document.PageSetup.PageWidth;
        double height = _document.PageSetup.PageHeight;

        _document.PageSetup.TogglePortrait();

        Assert.Equal(PageOrientation.Landscape, _document.PageSetup.Orientation);
        Assert.Equal(height, _document.PageSetup.PageWidth);
        Assert.Equal(width, _document.PageSetup.PageHeight);

        _document.PageSetup.TogglePortrait();

        Assert.Equal(PageOrientation.Portrait, _document.PageSetup.Orientation);
        Assert.Equal(width, _document.PageSetup.PageWidth);
    }

    [Fact]
    public void ASectionIsLeftToRightUnlessItIsMadeRightToLeft()
    {
        Assert.Equal(SectionDirection.Ltr, _document.PageSetup.SectionDirection);

        _document.PageSetup.SectionDirection = SectionDirection.Rtl;

        Assert.Equal(SectionDirection.Rtl, _document.PageSetup.SectionDirection);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.Document.Body.Descendants<BiDi>());

        _document.PageSetup.SectionDirection = SectionDirection.Ltr;

        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.Document.Body.Descendants<BiDi>());
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.PageSetup.SectionDirection = (SectionDirection)7);
    }
}
