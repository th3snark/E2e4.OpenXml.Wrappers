using System;
using System.Collections.Generic;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Captions, tables of figures and TC entries. The markup is Word 16's (the caption-*, tof-* and toc-mark-entry scenarios of
/// MarkupConformanceTest); the values are those Word printed for the same calls.
/// </summary>
public class TableOfFiguresTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableOfFiguresTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Intro\rPicture one\rMiddle\rPicture two\rEnd");
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
        IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    private static string[] Entries(RangeWrapper range)
    {
        //The last paragraph of the range is the one the table was added in front of, which holds the end of the field.
        string[] all = range.Paragraphs.Select(p => p.Range.Text.Split('\t')[0].TrimEnd('\r')).ToArray();
        return all.Take(all.Length - 1).Where(t => t.Length > 0).ToArray();
    }

    [Fact]
    public void CaptionsAreNumberedInTheOrderOfTheText()
    {
        _document.Paragraphs[3].Range.InsertCaption("Figure", ": Second");
        _document.Paragraphs[1].Range.InsertCaption("Figure", ": First");

        string text = _document.Content.Text;
        Assert.Contains("Figure 1: First\r", text, StringComparison.Ordinal);
        Assert.Contains("Figure 2: Second\r", text, StringComparison.Ordinal);
        Assert.Equal("Caption", _document.Paragraphs[2].Style);
        Assert.Throws<ArgumentNullException>(() => _document.Paragraphs[1].Range.InsertCaption(null));
        Assert.Throws<NotSupportedException>(() => _document.Paragraphs[1].Range.InsertCaption("Figure", "x", "AutoText"));
        AssertValid();
    }

    [Fact]
    public void ATableOfFiguresListsTheCaptionsOfALabel()
    {
        _document.Paragraphs[3].Range.InsertCaption("Figure", ": Second");
        _document.Paragraphs[1].Range.InsertCaption("Figure", ": First");
        _document.Paragraphs[0].Range.InsertCaption("Table", ": Other");

        TableOfFiguresWrapper table = _document.TablesOfFigures.Add(_document.Range(0, 0), "Figure");

        Assert.Equal(1, _document.TablesOfFigures.Count);
        Assert.Equal(0, _document.TablesOfContents.Count);
        Assert.Equal(table, _document.TablesOfFigures[0]);
        Assert.Equal("Figure", table.Caption);
        Assert.True(table.IncludeLabel);
        Assert.True(table.IncludePageNumbers);
        Assert.True(table.RightAlignPageNumbers);
        Assert.False(table.UseHyperlinks);
        Assert.Equal(TabLeader.Dots, table.TabLeader);
        Assert.Equal(string.Empty, table.TableID);
        Assert.Equal(new[] { "Figure 1: First", "Figure 2: Second" }, Entries(table.Range));
        AssertValid();
    }

    [Fact]
    public void TheLabelAndNumberCanBeLeftOutOfTheEntries()
    {
        _document.Paragraphs[1].Range.InsertCaption("Figure", ": First");
        TableOfFiguresWrapper table = _document.TablesOfFigures.Add(_document.Range(0, 0), "Figure", false);

        Assert.False(table.IncludeLabel);
        Assert.Equal(new[] { "First" }, Entries(table.Range));

        table.IncludeLabel = true;
        Assert.Equal(new[] { "Figure 1: First" }, Entries(table.Range));
        table.Delete();
        Assert.Equal(0, _document.TablesOfFigures.Count);
        AssertValid();
    }

    [Fact]
    public void ATableOfFiguresWithoutCaptionsSaysSo()
    {
        TableOfFiguresWrapper table = _document.TablesOfFigures.Add(_document.Range(0, 0), "Figure");

        Assert.Contains("No table of figures entries found.", _document.Content.Text, StringComparison.Ordinal);
        Assert.Equal("Figure", table.Caption);
        Assert.Throws<ArgumentNullException>(() => _document.TablesOfFigures.Add(_document.Range(0, 0), null));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.TablesOfFigures[1]);
    }

    [Fact]
    public void MarkedEntriesAreListedByATableWithTheSameIdentifier()
    {
        _document.TablesOfContents.MarkEntry(_document.Range(0, 5), "Marked one", null, "A", 2);
        _document.TablesOfContents.MarkEntry(_document.Range(_document.Paragraphs[1].Range.Start, _document.Paragraphs[1].Range.Start + 7), "Marked two", null, "B", 1);

        TableOfContentsWrapper toc = _document.TablesOfContents.Add(_document.Range(0, 0), false, 1, 3, true, "A");

        Assert.Equal(new[] { "Marked one" }, Entries(toc.Range));
        Assert.True(toc.UseFields);
        Assert.Equal("A", toc.TableID);

        toc.Update();
        Assert.Equal(new[] { "Marked one" }, Entries(toc.Range));
        AssertValid();
    }

    [Fact]
    public void MarkingAnEntryNeedsAnEntryAndALevel()
    {
        Assert.Throws<ArgumentException>(() => _document.TablesOfContents.MarkEntry(_document.Range(0, 5)));
        Assert.Throws<NotSupportedException>(() => _document.TablesOfContents.MarkEntry(_document.Range(0, 5), null, "AutoText"));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.TablesOfContents.MarkEntry(_document.Range(0, 5), "x", null, null, 10));
    }
}
