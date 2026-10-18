using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class BookmarkWordMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public BookmarkWordMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "alpha beta gamma delta";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void BookmarksAreInAlphabeticalOrderByDefault()
    {
        _document.Bookmarks.Add("zulu", _document.Range(0, 5));
        _document.Bookmarks.Add("Alpha", _document.Range(6, 10));
        _document.Bookmarks.Add("mike", _document.Range(11, 16));

        Assert.Equal(BookmarkSorting.Name, _document.Bookmarks.DefaultSorting);
        Assert.Equal(new[] { "Alpha", "mike", "zulu" }, _document.Bookmarks.Select(b => b.Name).ToArray());
        Assert.Equal("Alpha", _document.Bookmarks[0].Name);

        _document.Bookmarks.DefaultSorting = BookmarkSorting.Location;
        Assert.Equal(new[] { "zulu", "Alpha", "mike" }, _document.Bookmarks.Select(b => b.Name).ToArray());
        Assert.Equal(BookmarkSorting.Location, _document.Bookmarks.DefaultSorting);
    }

    [Fact]
    public void HiddenBookmarksAreLeftOutUnlessAskedFor()
    {
        _document.Bookmarks.Add("_Toc1", _document.Range(0, 5));
        _document.Bookmarks.Add("shown", _document.Range(6, 10));

        Assert.Equal(1, _document.Bookmarks.Count);
        Assert.False(_document.Bookmarks.Exists("_Toc1"));
        Assert.Null(_document.Bookmarks["_Toc1"]);

        _document.Bookmarks.ShowHidden = true;
        Assert.Equal(2, _document.Bookmarks.Count);
        Assert.True(_document.Bookmarks.Exists("_Toc1"));
        Assert.Equal("_Toc1", _document.Bookmarks[0].Name);
        Assert.True(_document.Content.Bookmarks.ShowHidden);
    }

    [Fact]
    public void ABookmarkKnowsWhereItIsAndCanBeCopied()
    {
        BookmarkWrapper bookmark = _document.Bookmarks.Add("first", _document.Range(6, 10));
        BookmarkWrapper point = _document.Bookmarks.Add("point", _document.Range(3, 3));

        Assert.Equal(6, bookmark.Start);
        Assert.Equal(10, bookmark.End);
        Assert.False(bookmark.Empty);
        Assert.True(point.Empty);

        BookmarkWrapper copy = bookmark.Copy("second");
        Assert.Equal("beta", copy.Range.Text);
        Assert.Equal(6, copy.Start);
        Assert.Equal(2 + 1, _document.Bookmarks.Count);
        Assert.Throws<ArgumentException>(() => bookmark.Copy("1bad"));
        _document.Save();
        Assert.Empty(new OpenXmlValidator().Validate(_document.WordprocessingDocument).Select(e => e.Description));
    }
}
