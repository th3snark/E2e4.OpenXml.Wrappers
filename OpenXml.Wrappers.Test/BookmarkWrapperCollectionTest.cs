using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using System;

namespace OpenXml.Wrappers.Test;

public class BookmarkWrapperCollectionTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public BookmarkWrapperCollectionTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void IterateCollectionReturnsCorrectItemsInCorrectOrder()
    {
        BookmarkWrapperCollection bookmarks = _document1.Bookmarks;
        bookmarks.DefaultSorting = BookmarkSorting.Location;
        List<string> bookmarkNames = new List<string>();
        List<int> bookmarkIds = new List<int>();
        foreach (BookmarkWrapper bookmark in bookmarks)
        {
            bookmarkNames.Add(bookmark.Name);
            bookmarkIds.Add(bookmark.Id);
        }

        Assert.Equal(
           new string[] { "myBookMark7", "myBookMark2", "myBookMark3", "myBookMark1",
                       "myBookMark5", "myBookMark6", "myBookMark8", "myBookMark4", },
           bookmarkNames.ToArray());

        Assert.Equal(new int[] { 0, 1, 2, 3, 4, 5, 6, 7 }, bookmarkIds.ToArray());
    }

    [Fact]
    public void NameIndexerReturnsCorrectBookmark()
    {
        BookmarkWrapperCollection bookmarks = _document1.Bookmarks;
        BookmarkWrapper bookmark = bookmarks["myBookMark8"];
        Assert.Equal("myBookMark8", bookmark.Name);
        Assert.Equal(6, bookmark.Id);
    }

    [Fact]
    public void NameIndexerReturnsNullForMissingBookmark()
    {
        BookmarkWrapper bookmark = _document1.Bookmarks["missingBookmark"];
        Assert.Null(bookmark);
    }

    [Fact]
    public void IntIndexerReturnsCorrectExceptionForMissingBookmark()
    {
        ArgumentOutOfRangeException ex = Assert.Throws<ArgumentOutOfRangeException>(() => _document1.Bookmarks[23]);
        Assert.Equal("index", ex.ParamName);
        Assert.StartsWith("The requested member of the collection '23' does not exist.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void IntIndexerReturnsCorrectBookmark()
    {
        BookmarkWrapperCollection bookmarks = _document1.Bookmarks;
        BookmarkWrapper bookmark = bookmarks[4];
        Assert.Equal("myBookMark5", bookmark.Name);
        Assert.Equal(4, bookmark.Id);
    }

    [Fact]
    public void DeleteAllRemovesTheCorrectBookmarksFromTheDocument()
    {
        Assert.Equal(8, _document1.Bookmarks.Count);
        _document1.Bookmarks.DeleteAll();
        Assert.Equal(0, _document1.Bookmarks.Count);
    }

    [Fact]
    public void DeleteAllAlsoRemovesTheEndOfABookmarkThatEndsOutsideTheRange()
    {
        Paragraph first = new Paragraph(new BookmarkStart() { Id = "1", Name = "a" }, new Run(new Text("a")));
        Paragraph second = new Paragraph(new Run(new Text("b")), new BookmarkEnd() { Id = "1" });
        Body body = new Body(first, second);

        new RangeWrapper(first).Bookmarks.DeleteAll();

        Assert.Empty(body.Descendants<BookmarkStart>());
        Assert.Empty(body.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void DeleteAllLeavesABookmarkWhoseStartIsOutsideTheRange()
    {
        Paragraph first = new Paragraph(new BookmarkStart() { Id = "1", Name = "a" }, new Run(new Text("a")));
        Paragraph second = new Paragraph(new Run(new Text("b")), new BookmarkEnd() { Id = "1" });
        Body body = new Body(first, second);

        new RangeWrapper(second).Bookmarks.DeleteAll();

        Assert.Single(body.Descendants<BookmarkStart>());
        Assert.Single(body.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void DeleteAllRemovesTheCorrectBookmarksFromATable()
    {
        Assert.Equal(6, _document1.Tables[0].Range.Bookmarks.Count);
        _document1.Tables[0].Range.Bookmarks.DeleteAll();
        Assert.Equal(0, _document1.Tables[0].Range.Bookmarks.Count);
        Assert.Equal(2, _document1.Bookmarks.Count);
    }
}
