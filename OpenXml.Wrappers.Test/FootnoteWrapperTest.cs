using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Footnotes and endnotes. The markup is Word 16's (the footnote-* and endnote-* scenarios of MarkupConformanceTest); the values of the members
/// are those Word gave for the same calls.
/// </summary>
public class FootnoteWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public FootnoteWrapperTest()
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

    private void AssertValid()
    {
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    [Fact]
    public void AFootnoteHasItsTextAndMarkWhereWordPutsThem()
    {
        FootnoteWrapper note = _document.Footnotes.Add(_document.Range(5, 5), null, "The note text");

        Assert.Equal(1, _document.Footnotes.Count);
        Assert.Equal(0, note.Index);
        Assert.Equal("The note text", note.Range.Text);
        Assert.Equal(2, note.Range.Start);
        Assert.Equal(15, note.Range.End);
        Assert.Equal(StoryType.Footnotes, note.Range.StoryType);
        Assert.Equal(5, note.Reference.Start);
        Assert.Equal(6, note.Reference.End);
        Assert.Equal(StoryType.MainText, note.Reference.StoryType);
        Assert.Equal("Hello world again", _document.Content.Text.Replace("\r", string.Empty));
        AssertValid();
    }

    [Fact]
    public void TheMarkGoesAtTheEndOfTheRangeAndTheRangeKeepsItsText()
    {
        FootnoteWrapper note = _document.Footnotes.Add(_document.Range(5, 10));

        Assert.Equal(10, note.Reference.Start);
        Assert.Equal(string.Empty, note.Range.Text);
        AssertValid();
    }

    [Fact]
    public void NotesAreNumberedInTheOrderOfTheText()
    {
        FootnoteWrapper later = _document.Footnotes.Add(_document.Range(11, 11), null, "second in the file");
        FootnoteWrapper earlier = _document.Footnotes.Add(_document.Range(5, 5), null, "first in the text");

        Assert.Equal(new[] { "first in the text", "second in the file" }, _document.Footnotes.Select(n => n.Range.Text).ToArray());
        Assert.Equal(0, earlier.Index);
        Assert.Equal(1, later.Index);
        Assert.Equal(earlier, _document.Footnotes[0]);
        AssertValid();
    }

    [Fact]
    public void DeletingANoteTakesItsMarkAndNumbersTheRestAgain()
    {
        _document.Footnotes.Add(_document.Range(5, 5), null, "one");
        FootnoteWrapper second = _document.Footnotes.Add(_document.Range(11, 11), null, "two");

        _document.Footnotes[0].Delete();

        Assert.Equal(1, _document.Footnotes.Count);
        Assert.Equal("two", _document.Footnotes[0].Range.Text);
        Assert.Equal(0, second.Index);
        Assert.Equal("Hello world again", _document.Content.Text.Replace("\r", string.Empty));
        AssertValid();
    }

    [Fact]
    public void AMarkOfYourOwnIsKeptAsText()
    {
        FootnoteWrapper note = _document.Footnotes.Add(_document.Range(5, 5), "ref*", "custom mark");

        Assert.Equal("custom mark", note.Range.Text);
        Assert.Equal("Helloref* world again", _document.Content.Text.Replace("\r", string.Empty));
        AssertValid();
    }

    [Fact]
    public void ARangeHasOnlyTheNotesWhoseMarksAreInIt()
    {
        _document.Footnotes.Add(_document.Range(5, 5), null, "one");
        _document.Footnotes.Add(_document.Range(12, 12), null, "two");

        RangeWrapper start = _document.Range(0, 4);
        RangeWrapper everything = _document.Content;

        Assert.Equal(0, start.Footnotes.Count);
        Assert.Equal(2, everything.Footnotes.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Footnotes[2]);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Footnotes[-1]);
    }

    [Fact]
    public void EndnotesAreTheSameAsFootnotesWithTheirOwnParts()
    {
        FootnoteWrapper footnote = _document.Footnotes.Add(_document.Range(5, 5), null, "The footnote");
        EndnoteWrapper endnote = _document.Endnotes.Add(_document.Range(11, 11), null, "The end note");

        Assert.Equal(1, _document.Footnotes.Count);
        Assert.Equal(1, _document.Endnotes.Count);
        Assert.Equal("The end note", endnote.Range.Text);
        Assert.Equal(StoryType.Endnotes, endnote.Range.StoryType);
        Assert.Equal(footnote, _document.Footnotes[0]);

        endnote.Delete();

        Assert.Equal(0, _document.Endnotes.Count);
        Assert.Equal(1, _document.Footnotes.Count);
        AssertValid();
    }

    [Fact]
    public void ANewDocumentHasNoNotesParts()
    {
        Assert.Null(_document.WordprocessingDocument.MainDocumentPart.FootnotesPart);
        Assert.Equal(0, _document.Footnotes.Count);

        _document.Footnotes.Add(_document.Range(5, 5));

        Assert.NotNull(_document.WordprocessingDocument.MainDocumentPart.FootnotesPart);
        Assert.NotNull(_document.WordprocessingDocument.MainDocumentPart.EndnotesPart);
    }

    [Fact]
    public void ANoteCannotBeAddedInsideANote()
    {
        FootnoteWrapper note = _document.Footnotes.Add(_document.Range(5, 5), null, "text");

        Assert.Throws<NotSupportedException>(() => _document.Footnotes.Add(note.Range));
        Assert.Throws<ArgumentNullException>(() => _document.Footnotes.Add(null));
    }
}
