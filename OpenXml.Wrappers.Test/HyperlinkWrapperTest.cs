using System;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Hyperlinks, Hyperlink and Range.Hyperlinks. The markup is Word 16's (the hyperlink-* scenarios of MarkupConformanceTest); the values of the
/// members are those Word gave for the same calls.
/// </summary>
public class HyperlinkWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public HyperlinkWrapperTest()
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
    public void AHyperlinkReadsBackWhatWasGiven()
    {
        _document.Bookmarks.Add("Mark", _document.Range(6, 11));

        HyperlinkWrapper link = _document.Hyperlinks.Add(_document.Range(0, 5), "https://example.org/page", "Mark", "A tip", "Shown", "_blank");

        Assert.Equal("https://example.org/page", link.Address);
        Assert.Equal("Mark", link.SubAddress);
        Assert.Equal("A tip", link.ScreenTip);
        Assert.Equal("Shown", link.TextToDisplay);
        Assert.Equal("_blank", link.Target);
        Assert.Equal("https://example.org/page - Mark", link.Name);
        Assert.Equal(HyperlinkType.Range, link.Type);
        Assert.Equal("Shown", link.Range.Text);
    }

    [Fact]
    public void ThePlaceAloneIsTheName()
    {
        HyperlinkWrapper link = _document.Hyperlinks.Add(_document.Range(0, 5), string.Empty, "Mark");

        Assert.Equal(string.Empty, link.Address);
        Assert.Equal("Mark", link.Name);
        Assert.Equal("Hello", link.TextToDisplay);
    }

    [Fact]
    public void AMailSubjectIsPartOfTheAddress()
    {
        HyperlinkWrapper link = _document.Hyperlinks.Add(_document.Range(0, 5), "mailto:a@b.com");
        Assert.Equal(string.Empty, link.EmailSubject);

        link.EmailSubject = "Hi there";

        Assert.Equal("Hi there", link.EmailSubject);
        Assert.Equal("mailto:a@b.com?subject=Hi there", link.Address);
    }

    [Fact]
    public void ARangeWithoutTextGetsTheAddressForItsText()
    {
        _document.Hyperlinks.Add(_document.Range(5, 5), "https://example.com/");

        Assert.Equal("Hellohttps://example.com/ world again", _document.Content.Text.TrimEnd('\r'));
    }

    [Fact]
    public void AddingNeedsAnAddressOrAPlace()
    {
        Assert.Throws<ArgumentException>(() => _document.Hyperlinks.Add(_document.Range(0, 5), string.Empty));
        Assert.Throws<ArgumentNullException>(() => _document.Hyperlinks.Add(null, "https://example.com/"));
    }

    [Fact]
    public void TheCollectionCountsAndIndexesFromZero()
    {
        HyperlinkWrapper later = _document.Hyperlinks.Add(_document.Range(12, 17), "https://later.example/");
        HyperlinkWrapper first = _document.Hyperlinks.Add(_document.Range(0, 5), "https://first.example/");

        Assert.Equal(2, _document.Hyperlinks.Count);
        Assert.Equal(first, _document.Hyperlinks[0]);
        Assert.Equal(later, _document.Hyperlinks[1]);
        Assert.Equal(new[] { "https://first.example/", "https://later.example/" }, _document.Hyperlinks.Select(l => l.Address).ToArray());
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Hyperlinks[2]);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Hyperlinks[-1]);
    }

    [Fact]
    public void ARangeHasOnlyTheHyperlinksInIt()
    {
        _document.Hyperlinks.Add(_document.Range(12, 17), "https://later.example/");
        _document.Hyperlinks.Add(_document.Range(0, 5), "https://first.example/");

        RangeWrapper firstWord = _document.Hyperlinks[0].Range;

        Assert.Equal(1, firstWord.Hyperlinks.Count);
        Assert.Equal("https://first.example/", firstWord.Hyperlinks[0].Address);
        Assert.Equal(2, _document.Content.Hyperlinks.Count);
    }

    [Fact]
    public void ADeletedHyperlinkLeavesItsTextAndDropsItsAddress()
    {
        HyperlinkWrapper link = _document.Hyperlinks.Add(_document.Range(0, 5), "https://example.com/");

        link.Delete();

        Assert.Equal(0, _document.Hyperlinks.Count);
        Assert.Equal("Hello world again", _document.Content.Text.TrimEnd('\r'));
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.HyperlinkRelationships);
    }

    [Fact]
    public void TheTextCanBeChanged()
    {
        HyperlinkWrapper link = _document.Hyperlinks.Add(_document.Range(0, 5), "https://example.com/");

        link.TextToDisplay = "Longer replacement";

        Assert.Equal("Longer replacement", link.TextToDisplay);
        Assert.Equal("Longer replacement world again", _document.Content.Text.TrimEnd('\r'));
        Assert.Equal(1, _document.Hyperlinks.Count);
    }

    [Fact]
    public void AHyperlinkFieldIsAHyperlinkYouCanReadAndDelete()
    {
        _document.Fields.Add(_document.Range(5, 5), FieldType.Empty, "HYPERLINK \"https://field.example/\" \\l \"Place\"", false);
        Assert.Equal(1, _document.Hyperlinks.Count);

        HyperlinkWrapper link = _document.Hyperlinks[0];

        Assert.Equal("https://field.example/", link.Address);
        Assert.Equal("Place", link.SubAddress);
        Assert.Throws<NotSupportedException>(() => link.Address = "https://changed.example/");

        link.Delete();

        Assert.Equal(0, _document.Hyperlinks.Count);
    }

    [Fact]
    public void AHyperlinkCannotSpanParagraphs()
    {
        _document.Content.InsertAfter("\rSecond paragraph");

        Assert.Throws<NotSupportedException>(() => _document.Hyperlinks.Add(_document.Range(10, 25), "https://example.com/"));
    }
}
