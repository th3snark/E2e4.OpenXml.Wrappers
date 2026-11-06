using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Tables of contents. The markup is Word 16's (the toc-* scenarios of MarkupConformanceTest).
/// </summary>
public class TableOfContentsWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableOfContentsWrapperTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("First\rSome text\rSecond\rSub heading\rEnd");
        _document.Paragraphs[0].Style = "Heading 1";
        _document.Paragraphs[2].Style = "Heading 1";
        _document.Paragraphs[3].Style = "Heading 2";
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    private string[] EntryTexts()
    {
        return _document.Paragraphs.Select(p => p.Paragraph).Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC", StringComparison.Ordinal) == true)
            .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
    }

    [Fact]
    public void ADefaultTableListsEveryHeadingWithItsPage()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));

        Assert.Equal(1, _document.TablesOfContents.Count);
        Assert.Equal(toc, _document.TablesOfContents[0]);
        Assert.Equal(new[] { "First1", "Second1", "Sub heading1" }, EntryTexts());
        Assert.Equal(1, toc.UpperHeadingLevel);
        Assert.Equal(9, toc.LowerHeadingLevel);
        Assert.True(toc.UseHeadingStyles);
        Assert.True(toc.IncludePageNumbers);
        Assert.True(toc.RightAlignPageNumbers);
        Assert.False(toc.UseHyperlinks);
        Assert.False(toc.UseFields);
        Assert.Equal(string.Empty, toc.TableID);
        Assert.Equal(TabLeader.Dots, toc.TabLeader);
        Assert.Equal(1, _document.Content.TablesOfContents.Count);
        AssertValid();
    }

    [Fact]
    public void TheSettingsAreKeptInTheFieldCode()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0), true, 2, 3, true, "A", false, false, null, true, true, true);

        Assert.Equal(2, toc.UpperHeadingLevel);
        Assert.Equal(3, toc.LowerHeadingLevel);
        Assert.True(toc.UseFields);
        Assert.Equal("A", toc.TableID);
        Assert.False(toc.RightAlignPageNumbers);
        Assert.False(toc.IncludePageNumbers);
        Assert.True(toc.UseHyperlinks);
        Assert.True(toc.HidePageNumbersInWeb);

        toc.UpperHeadingLevel = 1;
        toc.IncludePageNumbers = true;
        toc.TableID = null;
        Assert.Equal(1, toc.UpperHeadingLevel);
        Assert.True(toc.IncludePageNumbers);
        Assert.Equal(string.Empty, toc.TableID);
        Assert.Throws<ArgumentOutOfRangeException>(() => toc.LowerHeadingLevel = 10);
        AssertValid();
    }

    [Fact]
    public void UpdateRebuildsTheEntriesAndKeepsThePagesOfTheOnesItHad()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));
        Field(toc, "Second").Descendants<Text>().First(t => t.Text == "1").Text = "4";
        _document.Paragraphs[_document.Paragraphs.Count - 1].Style = "Heading 1";

        toc.Update();

        Assert.Equal(new[] { "First1", "Second4", "Sub heading1", "End1" }, EntryTexts());
        Assert.Equal(1, _document.TablesOfContents.Count);
        AssertValid();
    }

    private Paragraph Field(TableOfContentsWrapper toc, string heading)
    {
        return _document.Paragraphs.Select(p => p.Paragraph).First(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value?.StartsWith("TOC", StringComparison.Ordinal) == true
            && p.Descendants<Text>().First().Text == heading);
    }

    [Fact]
    public void DeleteTakesTheEntriesAwayAndLeavesTheHeadings()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));
        toc.Delete();

        Assert.Equal(0, _document.TablesOfContents.Count);
        Assert.Empty(EntryTexts());
        Assert.Equal("First\rSome text\rSecond\rSub heading\rEnd", _document.Content.Text.TrimEnd('\r'));
        Assert.Equal(3, _document.WordprocessingDocument.MainDocumentPart.Document.Descendants<BookmarkStart>().Count());
        AssertValid();
    }

    [Fact]
    public void TheLeaderOfTheEntriesCanBeChanged()
    {
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));

        toc.TabLeader = TabLeader.Lines;
        Assert.Equal(TabLeader.Lines, toc.TabLeader);
        toc.Update();
        Assert.Equal(TabLeader.Lines, toc.TabLeader);
        Assert.Throws<ArgumentOutOfRangeException>(() => toc.TabLeader = (TabLeader)99);
    }

    [Fact]
    public void OutlineLevelsAndTheStylesAreBothUsed()
    {
        _document.Paragraphs[1].Format.OutlineLevel = ParagraphOutlineLevel.Level2;
        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0), true, 1, 3, false, null, true, true, null, false, false, true);

        Assert.Equal(new[] { "First1", "Some text1", "Second1", "Sub heading1" }, EntryTexts());
        Assert.True(toc.UpperHeadingLevel == 1);
    }

    [Fact]
    public void WithoutHeadingsTheTableSaysSo()
    {
        _document.Paragraphs[0].Style = "Normal";
        _document.Paragraphs[2].Style = "Normal";
        _document.Paragraphs[3].Style = "Normal";

        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0));

        Assert.Equal(1, _document.TablesOfContents.Count);
        Assert.StartsWith("No table of contents entries found.", _document.Paragraphs[0].Paragraph.InnerText, StringComparison.Ordinal);
        toc.Delete();
        Assert.Equal(0, _document.TablesOfContents.Count);
        AssertValid();
    }

    [Fact]
    public void TheArgumentsAreChecked()
    {
        Assert.Throws<NotSupportedException>(() => _document.TablesOfContents.Add(_document.Range(2, 2)));
        Assert.Throws<ArgumentException>(() => _document.TablesOfContents.Add(_document.Range(0, 0), true, 1, 3, false, null, true, true, "Quote"));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.TablesOfContents.Add(_document.Range(0, 0), true, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.TablesOfContents.Add(_document.Range(0, 0), true, 3, 2));
        Assert.Throws<ArgumentNullException>(() => _document.TablesOfContents.Add(null));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.TablesOfContents[0]);
    }
}
