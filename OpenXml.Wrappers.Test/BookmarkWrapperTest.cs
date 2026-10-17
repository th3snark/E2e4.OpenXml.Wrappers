using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using System;

namespace OpenXml.Wrappers.Test;

public class BookmarkWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;
    private WordprocessingDocumentWrapper _document2;

    public BookmarkWrapperTest()
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
    public void BookmarkWrappersOfTheSameBookmarkAreEqual()
    {
        Assert.True(_document1.Bookmarks[0] == _document1.Bookmarks[0]);
        Assert.True(_document1.Bookmarks[0] != _document1.Bookmarks[1]);
        Assert.Equal(_document1.Bookmarks[0].GetHashCode(), _document1.Bookmarks[0].GetHashCode());
    }

    [Fact]
    public void NamePropertyIsCorrect()
    {
        Assert.Equal("myBookMark2", _document1.Bookmarks["myBookMark2"].Name);
        Assert.Equal("myBookMark8", _document1.Bookmarks["myBookMark8"].Name);
        Assert.Equal("myBookMark3", _document1.Bookmarks["myBookMark3"].Name);
    }

    [Fact]
    public void IdPropertyIsCorrect()
    {
        Assert.Equal(1, _document1.Bookmarks["myBookMark2"].Id);
        Assert.Equal(6, _document1.Bookmarks["myBookMark8"].Id);
        Assert.Equal(2, _document1.Bookmarks["myBookMark3"].Id);
    }

    [Fact]
    public void DeleteRemovesTheCorrectBookmark()
    {
        Assert.Equal(8, _document1.Bookmarks.Count);
        Assert.NotNull(_document1.Bookmarks["myBookMark3"]);

        _document1.Bookmarks["myBookMark3"].Delete();

        Assert.Equal(7, _document1.Bookmarks.Count);
        BookmarkWrapper bookmark = _document1.Bookmarks["myBookMark3"];
        Assert.Null(bookmark);
    }

    [Fact]
    public void RangeFirstParagraphRangeTextPropertyIsCorrect()
    {
        Assert.Equal("StabilitySummary\r\a", _document2.Bookmarks["confReportType"].Range.FirstParagraph.Range.Text);
        Assert.Equal("StabilitySummary", _document2.Bookmarks["confReportType"].Range.FirstParagraph.Text);
        Assert.Equal("StabilitySummary", _document2.Bookmarks["confReportType"].Range.Text);

        Assert.Equal("SomeBoldText", _document1.Bookmarks["myBookMark8"].Range.Text);
    }

    [Fact]
    public void RangeTextPropertyIsCorrect()
    {
        Assert.Equal("StabilitySummary", _document2.Bookmarks["confReportType"].Range.Text);

        RangeWrapper range = _document1.Bookmarks["myBookMark8"].Range;
        Assert.Equal("SomeBoldText", range.Text);
        range.Text = "Some New Text!";
        //_document1.Bookmarks.Add("myBookMark8", range);
        Assert.Equal("Some New Text!", _document1.Bookmarks["myBookMark8"].Range.Text);
    }

    [Fact]
    public void RangeParagraphsCountPropertyIsCorrect()
    {
        Assert.Equal(35, _document1.Bookmarks["myBookMark1"].Range.Paragraphs.Count);
        Assert.Equal(8, _document1.Bookmarks["myBookMark2"].Range.Paragraphs.Count);
        Assert.Equal(16, _document1.Bookmarks["myBookMark3"].Range.Paragraphs.Count); //Should possible account for vertical column but will do for now.
        Assert.Equal(3, _document1.Bookmarks["myBookMark4"].Range.Paragraphs.Count);
        Assert.Equal(1, _document1.Bookmarks["myBookMark5"].Range.Paragraphs.Count);
        Assert.Equal(1, _document1.Bookmarks["myBookMark6"].Range.Paragraphs.Count);
        Assert.Equal(1, _document1.Bookmarks["myBookMark7"].Range.Paragraphs.Count);
        Assert.Equal(1, _document1.Bookmarks["myBookMark8"].Range.Paragraphs.Count);
    }
}
