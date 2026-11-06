using System;
using System.Collections.Generic;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The styles a table of contents makes entries from besides the headings. The markup is Word 16's (the toc-added-styles* scenarios of
/// MarkupConformanceTest); the values are those Word printed for the same calls.
/// </summary>
public class TableOfContentsHeadingStylesTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableOfContentsHeadingStylesTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("First\rSome text\rTitled\rMore text\rSub\rEnd");
        _document.Paragraphs[0].Style = "Heading 1";
        _document.Paragraphs[2].Style = "Title";
        _document.Paragraphs[4].Style = "Subtitle";
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private string[] Entries(TableOfContentsWrapper toc)
    {
        return toc.Range.Paragraphs.Select(p => p.Range.Text.Split('\t')[0].TrimEnd('\r')).Where(t => t.Length > 0).ToArray();
    }

    [Fact]
    public void AddedStylesAreEntriesAtTheirLevels()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0), true, 1, 3, false, null, true, true, "Title,1,Subtitle,2");

        Assert.Equal(2, toc.HeadingStyles.Count);
        Assert.Equal("Title", toc.HeadingStyles[0].Style);
        Assert.Equal(1, toc.HeadingStyles[0].Level);
        Assert.Equal("Subtitle", toc.HeadingStyles[1].Style);
        Assert.Equal(2, toc.HeadingStyles[1].Level);
        Assert.Contains("Titled", Entries(toc));
        Assert.Contains("Sub", Entries(toc));
    }

    [Fact]
    public void HeadingStylesCanBeAddedChangedAndDeleted()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));
        Assert.Equal(0, toc.HeadingStyles.Count);

        toc.HeadingStyles.Add("Title", 1);
        HeadingStyleWrapper subtitle = toc.HeadingStyles.Add("Subtitle", 2);
        Assert.Equal(2, toc.HeadingStyles.Count);
        Assert.Equal(new[] { "Title", "Subtitle" }, toc.HeadingStyles.Select(h => h.Style).ToArray());
        Assert.Contains("Titled", Entries(toc));

        subtitle.Level = 3;
        Assert.Equal(3, toc.HeadingStyles[1].Level);

        HeadingStyleWrapper title = toc.HeadingStyles[0];
        title.Delete();
        Assert.Equal(1, toc.HeadingStyles.Count);
        Assert.DoesNotContain("Titled", Entries(toc));
        Assert.Throws<InvalidOperationException>(() => title.Level);
    }

    [Fact]
    public void TheRangeOfHeadingLevelsIsKeptWhenStylesAreAdded()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0), true, 2, 4);

        toc.HeadingStyles.Add("Title", 1);

        Assert.Equal(2, toc.UpperHeadingLevel);
        Assert.Equal(4, toc.LowerHeadingLevel);
        Assert.Contains("Titled", Entries(toc));
    }

    [Fact]
    public void WithOutlineLevelsAddedStylesMakeNoEntries()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0), true, 1, 3, false, null, true, true, "Title,2", false, false, true);

        Assert.DoesNotContain("Titled", Entries(toc));
    }

    [Fact]
    public void TheArgumentsAreChecked()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));

        Assert.Throws<ArgumentException>(() => _document.TablesOfContents.Add(_document.Range(0, 0), true, 1, 3, false, null, true, true, "Title"));
        Assert.Throws<ArgumentException>(() => _document.TablesOfContents.Add(_document.Range(0, 0), true, 1, 3, false, null, true, true, "Title,10"));
        Assert.Throws<ArgumentNullException>(() => toc.HeadingStyles.Add(null, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => toc.HeadingStyles.Add("Title", 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => toc.HeadingStyles[0]);
        IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }
}
