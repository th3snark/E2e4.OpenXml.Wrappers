using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class HeaderFooterTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly Body _body;

    public HeaderFooterTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private MainDocumentPart MainPart
    {
        get { return _document.WordprocessingDocument.MainDocumentPart; }
    }

    //Puts a section break after the first paragraph so the document has two sections.
    private void AddSecondSection()
    {
        _body.InsertAt(new Paragraph(new ParagraphProperties(new SectionProperties()), new Run(new Text("one"))), 0);
        _body.InsertAt(new Paragraph(new Run(new Text("two"))), 1);
    }

    [Fact]
    public void ANewDocumentHasNoHeadersOrFooters()
    {
        SectionWrapper section = _document.Sections[0];

        Assert.Equal(3, section.Headers.Count);
        Assert.Equal(3, section.Footers.Count);
        Assert.False(section.Headers[HeaderFooterIndex.Primary].Exists);
        Assert.False(section.Footers[HeaderFooterIndex.FirstPage].Exists);
        Assert.False(section.Headers[HeaderFooterIndex.Primary].LinkToPrevious);
        Assert.Empty(MainPart.HeaderParts);
    }

    [Fact]
    public void ReadingTheRangeCreatesAnEmptyHeader()
    {
        HeaderFooterWrapper header = _document.Sections[0].Headers[HeaderFooterIndex.Primary];

        RangeWrapper range = header.Range;

        Assert.True(header.Exists);
        Assert.True(header.IsHeader);
        Assert.Equal(string.Empty, range.Text.TrimEnd('\r'));
        Assert.Single(MainPart.HeaderParts);
        HeaderReference reference = _body.Elements<SectionProperties>().Single().Elements<HeaderReference>().Single();
        Assert.Equal(HeaderFooterValues.Default, reference.Type.Value);
        Assert.Equal(MainPart.HeaderParts.Single(), MainPart.GetPartById(reference.Id.Value));
        Assert.IsType<HeaderReference>(_body.Elements<SectionProperties>().Single().FirstChild);
    }

    [Fact]
    public void SettingTheTextReplacesTheContentAndSplitsLines()
    {
        RangeWrapper range = _document.Sections[0].Footers[HeaderFooterIndex.Primary].Range;

        range.Text = "Page one\rSecond line";

        Assert.Equal("Page one\rSecond line\r", range.Text);
        Assert.Equal(2, range.Paragraphs.Count);
        Assert.Single(MainPart.FooterParts);

        range.InsertAfter("!");
        Assert.Equal("Page one\rSecond line!\r", range.Text);
        range.InsertBefore("> ");
        Assert.Equal("> Page one\rSecond line!\r", range.Text);
    }

    [Fact]
    public void TheContentSurvivesSavingAndReopening()
    {
        _document.Sections[0].Headers[HeaderFooterIndex.FirstPage].Range.Text = "Cover";
        _document.Save();
        string copy = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            _document.SaveAs(copy).Close();
            using (WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(copy, false))
            {
                HeaderFooterWrapper header = reopened.Sections[0].Headers[HeaderFooterIndex.FirstPage];
                Assert.True(header.Exists);
                Assert.Equal("Cover\r", header.Range.Text);
                Assert.False(reopened.Sections[0].Headers[HeaderFooterIndex.Primary].Exists);
            }
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void FindAndReplaceWorkInAHeader()
    {
        RangeWrapper range = _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range;
        range.Text = "Draft report";

        Assert.True(range.Find.Execute("Draft", matchCase: true, replaceWith: "Final", replace: ReplaceMode.All));

        Assert.Equal("Final report\r", range.Text);
    }

    [Fact]
    public void ASectionWithoutAHeaderIsLinkedToThePreviousOne()
    {
        AddSecondSection();
        HeaderFooterWrapper first = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        HeaderFooterWrapper second = _document.Sections[1].Headers[HeaderFooterIndex.Primary];
        first.Range.Text = "Shared";

        Assert.True(second.LinkToPrevious);
        Assert.False(second.Exists);
        Assert.Equal("Shared\r", second.Range.Text);
        Assert.False(second.Exists);
        Assert.Single(MainPart.HeaderParts);
    }

    [Fact]
    public void UnlinkingCopiesTheContentAndLinkingRemovesIt()
    {
        AddSecondSection();
        _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Shared";
        HeaderFooterWrapper second = _document.Sections[1].Headers[HeaderFooterIndex.Primary];

        second.LinkToPrevious = false;

        Assert.True(second.Exists);
        Assert.False(second.LinkToPrevious);
        Assert.Equal(2, MainPart.HeaderParts.Count());
        second.Range.Text = "Own";
        Assert.Equal("Shared\r", _document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text);

        second.LinkToPrevious = true;

        Assert.False(second.Exists);
        Assert.Single(MainPart.HeaderParts);
        Assert.Equal("Shared\r", second.Range.Text);
        Assert.Throws<InvalidOperationException>(() => _document.Sections[0].Headers[HeaderFooterIndex.Primary].LinkToPrevious = true);
    }

    [Fact]
    public void HeadersAndFootersOfTheSameKindAreIndependentAndEqual()
    {
        SectionWrapper section = _document.Sections[0];
        section.Headers[HeaderFooterIndex.Primary].Range.Text = "H";

        Assert.False(section.Footers[HeaderFooterIndex.Primary].Exists);
        Assert.False(section.Headers[HeaderFooterIndex.EvenPages].Exists);
        Assert.Equal(section.Headers[HeaderFooterIndex.Primary], _document.Sections[0].Headers[HeaderFooterIndex.Primary]);
        Assert.NotEqual(section.Headers[HeaderFooterIndex.Primary], section.Footers[HeaderFooterIndex.Primary]);
        Assert.Equal(section, section.Headers[HeaderFooterIndex.Primary].Parent);
        Assert.Equal(3, section.Headers.ToList().Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => section.Headers[(HeaderFooterIndex)9]);
    }

    [Fact]
    public void AFieldCanBeAddedToAFooter()
    {
        RangeWrapper range = _document.Sections[0].Footers[HeaderFooterIndex.Primary].Range;

        FieldWrapper field = range.Fields.Add(range, FieldType.Page);

        Assert.Equal(field, range.Fields.Single());
        Assert.Equal(FieldType.Page, field.Type);
    }
}
