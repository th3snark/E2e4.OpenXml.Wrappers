using System;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// StoryType, StoryRanges and NextStoryRange. The values are Word 16's, measured in a document with a header, a footer and two comments.
/// </summary>
public class StoryRangesTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public StoryRangesTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ANewDocumentHasOnlyItsMainText()
    {
        _document.Content.InsertAfter("Body text");

        Assert.Equal(1, _document.StoryRanges.Count);
        Assert.Equal(StoryType.MainText, _document.StoryRanges[StoryType.MainText].StoryType);
        Assert.Equal(StoryType.MainText, _document.Content.StoryType);
        Assert.Null(_document.Content.NextStoryRange);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.StoryRanges[StoryType.PrimaryHeader]);
    }

    [Fact]
    public void HeadersFootersAndCommentsAreStoriesOfTheirOwn()
    {
        _document.Content.InsertAfter("Body text");
        _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Header one";
        _document.Sections[0].Footers[HeaderFooterIndex.Primary].Range.Text = "Footer one";
        CommentWrapper first = _document.Comments.Add(_document.Range(0, 4), "A note");
        _document.Comments.Add(_document.Range(5, 9), "Second note");

        StoryType[] types = _document.StoryRanges.Select(r => r.StoryType).ToArray();

        Assert.Equal(new[] { StoryType.MainText, StoryType.Comments, StoryType.PrimaryHeader, StoryType.PrimaryFooter }, types);
        Assert.Equal(types.Length, _document.StoryRanges.Count);
        Assert.Equal(StoryType.PrimaryHeader, _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.StoryType);
        Assert.Equal(StoryType.PrimaryFooter, _document.Sections[0].Footers[HeaderFooterIndex.Primary].Range.StoryType);
        Assert.Equal(StoryType.Comments, first.Range.StoryType);
        Assert.Equal(StoryType.MainText, first.Scope.StoryType);
        Assert.Contains("A note", _document.StoryRanges[StoryType.Comments].Text, StringComparison.Ordinal);
    }

    [Fact]
    public void FirstPageAndEvenPageHeadersHaveTheirOwnTypes()
    {
        _document.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true;
        _document.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter = true;
        _document.Sections[0].Headers[HeaderFooterIndex.FirstPage].Range.Text = "First header";
        _document.Sections[0].Headers[HeaderFooterIndex.EvenPages].Range.Text = "Even header";
        _document.Sections[0].Footers[HeaderFooterIndex.FirstPage].Range.Text = "First footer";

        Assert.Equal(StoryType.FirstPageHeader, _document.StoryRanges[StoryType.FirstPageHeader].StoryType);
        Assert.Equal(StoryType.EvenPagesHeader, _document.StoryRanges[StoryType.EvenPagesHeader].StoryType);
        Assert.Equal(StoryType.FirstPageFooter, _document.StoryRanges[StoryType.FirstPageFooter].StoryType);
        Assert.Contains("Even header", _document.StoryRanges[StoryType.EvenPagesHeader].Text, StringComparison.Ordinal);
    }

    [Fact]
    public void NextStoryRangeFollowsTheHeadersThatSectionsDoNotShare()
    {
        _document.Content.InsertAfter("Body text");
        _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Header one";
        _document.Sections.Add();
        _document.Sections.Add();
        _document.Sections[1].Headers[HeaderFooterIndex.Primary].LinkToPrevious = false;
        _document.Sections[1].Headers[HeaderFooterIndex.Primary].Range.Text = "Header two";

        RangeWrapper header = _document.StoryRanges[StoryType.PrimaryHeader];
        RangeWrapper next = header.NextStoryRange;

        Assert.NotNull(next);
        Assert.Contains("Header two", next.Text, StringComparison.Ordinal);
        Assert.Null(next.NextStoryRange);
        Assert.Equal(StoryType.PrimaryHeader, next.StoryType);
    }

    [Fact]
    public void ABookmarkIsInTheStoryOfItsRange()
    {
        _document.Content.InsertAfter("Body text");
        _document.Bookmarks.Add("Mark", _document.Range(0, 4));

        Assert.Equal(StoryType.MainText, _document.Bookmarks["Mark"].StoryType);
    }
}
