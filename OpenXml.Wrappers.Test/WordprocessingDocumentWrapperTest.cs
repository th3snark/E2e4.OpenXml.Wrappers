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

    private static string NewTemporaryPath()
    {
        return System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    }

    [Fact]
    public void CreateInitializesTheMainDocumentAndCommentsRoots()
    {
        string path = NewTemporaryPath();
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, true))
            {
                Assert.NotNull(document.WordprocessingDocument.MainDocumentPart.Document.Body);
                Assert.NotNull(document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments);
                Assert.Equal(1, document.Paragraphs.Count);
                Assert.Equal(0, document.Comments.Count);
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void CreateWithCommentsButNoMainDocumentPartThrows()
    {
        Assert.Throws<ArgumentException>(() => WordprocessingDocumentWrapper.Create(NewTemporaryPath(), false, true));
    }

    [Fact]
    public void CreateWithoutAMainDocumentPartLeavesTheDocumentEmpty()
    {
        string path = NewTemporaryPath();
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, false, false))
            {
                Assert.Null(document.WordprocessingDocument.MainDocumentPart);
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void DisposeAndCloseCanBothBeCalled()
    {
        string path = NewTemporaryPath();
        try
        {
            WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false);
            document.Close();

            Assert.Null(Record.Exception(() => document.Dispose()));
            Assert.Null(Record.Exception(() => document.Close()));
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void ChangesAreSavedWhenTheDocumentIsDisposed()
    {
        string path = NewTemporaryPath();
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
            {
                document.Paragraphs[0].Text = "saved text";
            }

            using (WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(path))
            {
                Assert.Equal("saved text", reopened.Paragraphs[0].Text);
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void SaveWritesChangesWhileTheDocumentRemainsOpen()
    {
        string path = NewTemporaryPath();
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
            {
                document.Paragraphs[0].Text = "first";
                document.Save();
                document.Paragraphs[0].Text = "second";

                Assert.Equal("second", document.Paragraphs[0].Text);
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void SaveAsWritesACopyContainingTheCurrentChanges()
    {
        string path = NewTemporaryPath();
        string copyPath = NewTemporaryPath();
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
            {
                document.Paragraphs[0].Text = "copied text";
                using (WordprocessingDocumentWrapper copy = document.SaveAs(copyPath))
                {
                    Assert.Equal("copied text", copy.Paragraphs[0].Text);
                }

                document.Paragraphs[0].Text = "original only";
            }

            using (WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(copyPath))
            {
                Assert.Equal("copied text", reopened.Paragraphs[0].Text);
            }
        }
        finally
        {
            System.IO.File.Delete(path);
            System.IO.File.Delete(copyPath);
        }
    }

    [Fact]
    public void DocumentsCanBeCreatedAndOpenedOverStreams()
    {
        using (System.IO.MemoryStream stream = new System.IO.MemoryStream())
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
            {
                document.Paragraphs[0].Text = "in a stream";
            }

            using (WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(stream, false))
            {
                Assert.Equal("in a stream", reopened.Paragraphs[0].Text);
            }
        }
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
        Assert.Equal("Last Line\r", _document1.Content.LastParagraph.Range.Text);
        Assert.Equal("Last Line", _document1.Content.LastParagraph.Text);
    }

    [Fact]
    public void FirstParagraphPropertyIsCorrect()
    {
        Assert.Equal("Test Document\r", _document1.Content.FirstParagraph.Range.Text);
        Assert.Equal("Test Document", _document1.Content.FirstParagraph.Text);
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
