using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace OpenXml.Wrappers.Test;

public class WordprocessingDocumentWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;
    private WordprocessingDocumentWrapper _document2;

    public WordprocessingDocumentWrapperTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
        _document2 = WordprocessingDocumentWrapper.Open(TestFile.Document2);
    }

    public void Dispose()
    {
        _document1.Close();
        _document2.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void BookmarksPropertyIsNotNull()
    {
        Assert.NotNull(_document1.Bookmarks);
    }

    [Fact]
    public void TablesPropertyIsNotNull()
    {
        Assert.NotNull(_document1.Tables);
    }

    [Fact]
    public void LastParagraphPropertyIsCorrect()
    {
        Assert.Equal("Last Line\r", _document1.Range.LastParagraph.Range.Text);
        Assert.Equal("Last Line", _document1.Range.LastParagraph.Text);
    }

    [Fact]
    public void FirstParagraphPropertyIsCorrect()
    {
        Assert.Equal("Test Document\r", _document1.Range.FirstParagraph.Range.Text);
        Assert.Equal("Test Document", _document1.Range.FirstParagraph.Text);
    }

    [Fact]
    public void _Test1()
    {
        string reportDocumentPath = "report.docx";
        WordprocessingDocumentWrapper reportWordprocessingDocumentWrapper = WordprocessingDocumentWrapper.Create(reportDocumentPath, true, true);

        //Initialize the main document part
        reportWordprocessingDocumentWrapper.WordprocessingDocument.MainDocumentPart.Document =
           new Document(
              new Body(
                 new Paragraph(),
                 new SectionProperties(
                       new PageSize() { Width = 11906, Height = 16838 },
                       new PageMargin() { Top = 1440, Right = 1440, Left = 1440, Header = 708, Footer = 708, Gutter = 0 },
                       new Columns() { Space = "708" },
                       new DocGrid() { LinePitch = 360 }
                 )
              )
           );

        _document2.CopyAllPartsTo(reportWordprocessingDocumentWrapper);
        _document2.CopyLayoutTo(reportWordprocessingDocumentWrapper);

        reportWordprocessingDocumentWrapper.Close();
    }
}
