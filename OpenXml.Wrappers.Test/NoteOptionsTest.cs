using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The options of footnotes and endnotes, converting and swapping them. The markup is Word 16's (the note-option-*, note-convert-* and
/// note-swap scenarios of MarkupConformanceTest); the defaults are what Word answered in a new document.
/// </summary>
public class NoteOptionsTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public NoteOptionsTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
        _document.Footnotes.Add(_document.Range(5, 5), null, "one");
        _document.Endnotes.Add(_document.Range(11, 11), null, "two");
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
    public void ANewDocumentHasWordsDefaults()
    {
        Assert.Equal(FootnoteLocation.BottomOfPage, _document.Footnotes.Location);
        Assert.Equal(NumberingRule.Continuous, _document.Footnotes.NumberingRule);
        Assert.Equal(NoteNumberStyle.Arabic, _document.Footnotes.NumberStyle);
        Assert.Equal(1, _document.Footnotes.StartingNumber);
        Assert.Equal(EndnoteLocation.EndOfDocument, _document.Endnotes.Location);
        Assert.Equal(NumberingRule.Continuous, _document.Endnotes.NumberingRule);
        Assert.Equal(NoteNumberStyle.LowercaseRoman, _document.Endnotes.NumberStyle);
        Assert.Equal(1, _document.Endnotes.StartingNumber);
    }

    [Fact]
    public void TheOptionsReadBackAndGoToTheSectionsToo()
    {
        _document.Footnotes.NumberStyle = NoteNumberStyle.UppercaseLetter;
        _document.Footnotes.Location = FootnoteLocation.BeneathText;
        _document.Endnotes.Location = EndnoteLocation.EndOfSection;

        Assert.Equal(NoteNumberStyle.UppercaseLetter, _document.Footnotes.NumberStyle);
        Assert.Equal(FootnoteLocation.BeneathText, _document.Footnotes.Location);
        Assert.Equal(EndnoteLocation.EndOfSection, _document.Endnotes.Location);
        Assert.Equal(NoteNumberStyle.UppercaseLetter, _document.Content.FootnoteOptions.NumberStyle);
        AssertValid();
    }

    [Fact]
    public void ARestartingNumberOtherThanOneNeedsContinuousNumbering()
    {
        _document.Footnotes.StartingNumber = 3;

        Assert.Throws<ArgumentException>(() => _document.Footnotes.NumberingRule = NumberingRule.RestartPage);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Footnotes.StartingNumber = 0);

        _document.Footnotes.StartingNumber = 1;
        _document.Footnotes.NumberingRule = NumberingRule.RestartSection;

        Assert.Throws<ArgumentException>(() => _document.Footnotes.StartingNumber = 5);
        Assert.Equal(NumberingRule.RestartSection, _document.Footnotes.NumberingRule);
    }

    [Fact]
    public void ASectionCanHaveOptionsOfItsOwn()
    {
        _document.Content.InsertAfter("\rSecond");
        _document.Sections.Add();

        _document.Sections[1].Range.FootnoteOptions.NumberStyle = NoteNumberStyle.LowercaseLetter;

        Assert.Equal(NoteNumberStyle.LowercaseLetter, _document.Sections[1].Range.FootnoteOptions.NumberStyle);
        Assert.Equal(NoteNumberStyle.Arabic, _document.Sections[0].Range.FootnoteOptions.NumberStyle);
        Assert.Equal(NoteNumberStyle.Arabic, _document.Footnotes.NumberStyle);
        AssertValid();
    }

    [Fact]
    public void ConvertingMakesTheNotesTheOtherKind()
    {
        _document.Footnotes.Convert();

        Assert.Equal(0, _document.Footnotes.Count);
        Assert.Equal(2, _document.Endnotes.Count);
        Assert.Equal(new[] { "one", "two" }, _document.Endnotes.Select(n => n.Range.Text).ToArray());
        AssertValid();

        _document.Endnotes.Convert();

        Assert.Equal(2, _document.Footnotes.Count);
        Assert.Equal(0, _document.Endnotes.Count);
        AssertValid();
    }

    [Fact]
    public void SwappingExchangesTheKinds()
    {
        _document.Footnotes.SwapWithEndnotes();

        Assert.Equal("two", _document.Footnotes[0].Range.Text);
        Assert.Equal("one", _document.Endnotes[0].Range.Text);
        AssertValid();
    }

    [Fact]
    public void ConvertingNoNotesDoesNothing()
    {
        _document.Footnotes[0].Delete();

        _document.Footnotes.Convert();

        Assert.Equal(1, _document.Endnotes.Count);
    }

    [Fact]
    public void TheSeparatorsAreStoriesOfTheirOwn()
    {
        Assert.Equal(StoryType.FootnoteSeparator, _document.Footnotes.Separator.StoryType);
        Assert.Equal(StoryType.FootnoteContinuationSeparator, _document.Footnotes.ContinuationSeparator.StoryType);
        Assert.Equal(StoryType.EndnoteSeparator, _document.Endnotes.Separator.StoryType);
        Assert.Equal(StoryType.EndnoteContinuationSeparator, _document.Endnotes.ContinuationSeparator.StoryType);
    }
}
