using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The text columns and page numbers of Word's object model beyond what they already had. Markup is Word 16's (the text-columns-* and
/// page-numbers-* scenarios of MarkupConformanceTest).
/// </summary>
public class TextColumnsAndPageNumbersTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TextColumnsAndPageNumbersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private TextColumnsWrapper Columns
    {
        get { return _document.Sections[0].PageSetup.TextColumns; }
    }

    private PageNumbersWrapper Numbers
    {
        get { return _document.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers; }
    }

    [Fact]
    public void AColumnIsAddedAtTheRightWithTheFirstTakingTheRest()
    {
        TextColumnWrapper added = Columns.Add(72, 10);

        Assert.Equal(2, Columns.Count);
        Assert.False(Columns.EvenlySpaced);
        Assert.Equal(72, added.Width);
        Assert.Equal(10, Columns[0].SpaceAfter);
        Assert.Equal(369.3, Columns[0].Width, 1);
        Assert.Equal(0, Columns[1].SpaceAfter);
    }

    [Fact]
    public void AColumnAddedWithoutAWidthKeepsTheColumnsEven()
    {
        Columns.Add();

        Assert.Equal(2, Columns.Count);
        Assert.True(Columns.EvenlySpaced);
        Assert.Equal(Columns[0].Width, Columns[1].Width);
    }

    [Fact]
    public void AWidthIsNeededOnceTheColumnsDiffer()
    {
        Columns.Add(72);

        Assert.Throws<ArgumentException>(() => Columns.Add());
    }

    [Fact]
    public void TheWidthAndSpacingOfAColumnCanBeChanged()
    {
        Columns.SetCount(3);

        Columns[0].Width = 100;
        Columns[1].SpaceAfter = 10;

        Assert.False(Columns.EvenlySpaced);
        Assert.Equal(100, Columns[0].Width);
        Assert.Equal(10, Columns[1].SpaceAfter);
        Assert.Equal(36, Columns[0].SpaceAfter);
        Assert.Equal(126.4, Columns[2].Width, 1);
    }

    [Fact]
    public void TheLastColumnHasNoSpaceAfterIt()
    {
        Columns.SetCount(2);

        Assert.Equal(0, Columns[1].SpaceAfter);
        Assert.Throws<InvalidOperationException>(() => Columns[1].SpaceAfter = 5);
    }

    [Fact]
    public void ColumnsCanBeEnumeratedAndOutOfRangeIndexesAreRejected()
    {
        Columns.SetCount(3);

        Assert.Equal(3, Columns.ToArray().Length);
        Assert.Throws<ArgumentOutOfRangeException>(() => Columns[3]);
        Assert.Throws<ArgumentOutOfRangeException>(() => Columns[-1]);
        Assert.Throws<ArgumentOutOfRangeException>(() => Columns.Add(0));
    }

    [Fact]
    public void ThereAreAtMost45Columns()
    {
        Columns.SetCount(45);

        Assert.Throws<InvalidOperationException>(() => Columns.Add());
    }

    [Fact]
    public void TheFlowDirectionIsTheDirectionOfTheSection()
    {
        Assert.Equal(FlowDirection.Ltr, Columns.FlowDirection);

        Columns.FlowDirection = FlowDirection.Rtl;

        Assert.Equal(FlowDirection.Rtl, Columns.FlowDirection);
        Assert.Equal(SectionDirection.Rtl, _document.Sections[0].PageSetup.SectionDirection);
    }

    [Fact]
    public void InsideAndOutsidePageNumbersAreAdded()
    {
        Numbers.Add(PageNumberAlignment.Inside);
        Numbers.Add(PageNumberAlignment.Outside);

        Assert.Equal(2, Numbers.Count);
        Assert.Equal(PageNumberAlignment.Outside, Numbers[0].Alignment);
        Assert.Equal(PageNumberAlignment.Inside, Numbers[1].Alignment);
        Assert.Equal(2, Numbers.ToArray().Length);
    }

    [Fact]
    public void ThePageNumbersAreReadByTheirAlignment()
    {
        Numbers.Add(PageNumberAlignment.Center);

        Assert.Equal(PageNumberAlignment.Center, Numbers[0].Alignment);
        Assert.Throws<ArgumentOutOfRangeException>(() => Numbers[1]);
    }

    [Fact]
    public void ChapterNumbersAreIncludedByTheLevelOfAHeading()
    {
        Assert.False(Numbers.IncludeChapterNumber);
        Assert.Equal(0, Numbers.HeadingLevelForChapter);
        Assert.Throws<InvalidOperationException>(() => Numbers.HeadingLevelForChapter = 1);

        Numbers.IncludeChapterNumber = true;
        Numbers.HeadingLevelForChapter = 2;
        Numbers.ChapterPageSeparator = SeparatorType.EnDash;

        Assert.True(Numbers.IncludeChapterNumber);
        Assert.Equal(2, Numbers.HeadingLevelForChapter);
        Assert.Equal(SeparatorType.EnDash, Numbers.ChapterPageSeparator);
        PageNumberType numbering = _document.WordprocessingDocument.MainDocumentPart.Document.Descendants<PageNumberType>().FirstOrDefault();
        Assert.Equal(3, numbering.ChapterStyle.Value);

        Assert.Throws<ArgumentOutOfRangeException>(() => Numbers.HeadingLevelForChapter = 9);
    }

    [Fact]
    public void TurningChapterNumbersOffLeavesNoNumberingBehind()
    {
        Numbers.IncludeChapterNumber = true;
        Numbers.IncludeChapterNumber = false;

        Assert.Null(_document.WordprocessingDocument.MainDocumentPart.Document.Descendants<PageNumberType>().FirstOrDefault());
        Assert.False(Numbers.IncludeChapterNumber);
    }

    [Fact]
    public void ThereIsNoFirstPageNumberWhenTheFirstPageIsDifferent()
    {
        Assert.True(Numbers.ShowFirstPageNumber);

        Numbers.ShowFirstPageNumber = false;

        Assert.False(Numbers.ShowFirstPageNumber);
        Assert.True(_document.Sections[0].PageSetup.DifferentFirstPageHeaderFooter);
    }
}
