using System;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Clearing formatting, bold and italic runs and inserting rows through the selection. The markup is Word 16's (the selection-* scenarios of
/// MarkupConformanceTest); these tests cover the values and the arguments.
/// </summary>
public class SelectionFormattingTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public SelectionFormattingTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ClearingCharacterFormattingKeepsWhatWasNotAsked()
    {
        _document.Range(0, 5).Style = "Strong";
        _document.Range(0, 5).Font.Size = 20;

        _document.Range(0, 5).Select();
        _document.Selection.ClearCharacterDirectFormatting();
        Assert.Equal(12, _document.Range(0, 5).Font.Size);
        Assert.True(_document.Range(0, 5).Font.Bold);

        _document.Range(0, 5).Font.Size = 20;
        _document.Selection.ClearCharacterStyle();
        Assert.Equal(20, _document.Range(0, 5).Font.Size);
        Assert.False(_document.Range(0, 5).Font.Bold);

        _document.Range(0, 5).Style = "Strong";
        _document.Selection.ClearCharacterAllFormatting();
        Assert.Equal(12, _document.Range(0, 5).Font.Size);
        Assert.False(_document.Range(0, 5).Font.Bold);
    }

    [Fact]
    public void ClearingParagraphFormattingKeepsWhatWasNotAsked()
    {
        _document.Paragraphs[0].Style = "Heading 1";
        _document.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        _document.Paragraphs[0].Range.Select();

        _document.Selection.ClearParagraphDirectFormatting();
        Assert.Equal("Heading 1", _document.Paragraphs[0].Style);
        Assert.Equal(ParagraphAlignment.Left, _document.Paragraphs[0].Range.ParagraphFormat.Alignment);

        _document.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        _document.Selection.ClearParagraphStyle();
        Assert.Equal("Normal", _document.Paragraphs[0].Style);
        Assert.Equal(ParagraphAlignment.Center, _document.Paragraphs[0].Range.ParagraphFormat.Alignment);

        _document.Selection.ClearParagraphAllFormatting();
        Assert.Equal(ParagraphAlignment.Left, _document.Paragraphs[0].Range.ParagraphFormat.Alignment);
    }

    [Fact]
    public void ARunTogglesBoldAndItalic()
    {
        _document.Range(0, 5).Select();
        _document.Selection.BoldRun();
        Assert.True(_document.Range(0, 5).Font.Bold);
        Assert.True(_document.Range(0, 5).Font.BoldBi);

        _document.Selection.BoldRun();
        Assert.False(_document.Range(0, 5).Font.Bold);

        _document.Range(0, 11).Select();
        _document.Selection.ItalicRun();
        Assert.True(_document.Range(0, 11).Font.Italic);
        _document.Selection.ItalicRun();
        Assert.False(_document.Range(0, 11).Font.Italic);
    }

    [Fact]
    public void RowsAreInsertedAboveAndBelowTheSelection()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);
        table.Cell(0, 0).Select();

        _document.Selection.InsertRowsBelow(2);
        Assert.Equal(4, table.Rows.Count);

        table.Cell(0, 0).Select();
        _document.Selection.InsertRowsAbove();
        Assert.Equal(5, table.Rows.Count);

        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Selection.InsertRowsAbove(0));
        _document.Range(_document.Content.End - 1, _document.Content.End - 1).Select();
        Assert.Throws<InvalidOperationException>(() => _document.Selection.InsertRowsBelow());
    }
}
