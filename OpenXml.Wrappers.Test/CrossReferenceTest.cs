using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Cross references and the hyperlinks of shapes. The expected values are those Word 16 gave for the same calls (the cross-reference-heading,
/// cross-reference-bookmark, cross-reference-caption and shape-hyperlink scenarios of tools/WordReference/scenarios.ps1).
/// </summary>
public class CrossReferenceTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public CrossReferenceTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private RangeWrapper EndOfLastParagraph()
    {
        RangeWrapper end = _document.Paragraphs[_document.Paragraphs.Count - 1].Range;
        end.SetRange(end.End - 1, end.End - 1);
        return end;
    }

    [Fact]
    public void AReferenceToAHeadingShowsItsText()
    {
        _document.Content.InsertAfter("First\rSome text\rSecond\rSee: ");
        _document.Paragraphs[0].Style = "Heading 1";
        _document.Paragraphs[2].Style = "Heading 1";
        EndOfLastParagraph().InsertCrossReference(CrossReferenceType.Heading, CrossReferenceKind.ContentText, 1);
        EndOfLastParagraph().InsertCrossReference(CrossReferenceType.Heading, CrossReferenceKind.PageNumber, 0);

        Assert.Equal("See: Second1", _document.Paragraphs[3].Range.Text.TrimEnd('\r'));
        Assert.Equal(2, _document.Fields.Count);
        Assert.Matches(@"^REF _Ref\d+ \\h$", _document.Fields[0].Code.Text);
        Assert.Matches(@"^PAGEREF _Ref\d+ \\h$", _document.Fields[1].Code.Text);
        Assert.Empty(DocumentIntegrity.CheckAll(_document.WordprocessingDocument));
    }

    [Fact]
    public void AReferenceToACaptionOrABookmark()
    {
        _document.Content.InsertAfter("Body\rAlpha beta gamma\rSee: ");
        _document.Paragraphs[0].Range.InsertCaption("Figure", ": first", string.Empty, CaptionPosition.Above, false);
        _document.Bookmarks.Add("Mark", _document.Range(_document.Paragraphs[2].Range.Start + 6, _document.Paragraphs[2].Range.Start + 10));
        Assert.Equal(new[] { "Figure 1: first" }, _document.GetCrossReferenceItems("Figure").ToArray());
        Assert.Equal(new[] { "Mark" }, _document.GetCrossReferenceItems(CrossReferenceType.Bookmark).ToArray());
        Assert.Empty(_document.GetCrossReferenceItems("Table"));
        Assert.Empty(_document.GetCrossReferenceItems(CrossReferenceType.Heading));

        EndOfLastParagraph().InsertCrossReference("Figure", CrossReferenceKind.EntireCaption, 0);
        EndOfLastParagraph().InsertCrossReference("Figure", CrossReferenceKind.OnlyLabelAndNumber, 0);
        EndOfLastParagraph().InsertCrossReference("Mark", CrossReferenceKind.ContentText, false);
        Assert.Equal("See: Figure 1: firstFigure 1beta", _document.Paragraphs[3].Range.Text.TrimEnd('\r'));
        Assert.Empty(DocumentIntegrity.CheckAll(_document.WordprocessingDocument));
    }

    [Fact]
    public void UnsupportedAndWrongReferencesFail()
    {
        _document.Content.InsertAfter("Text");
        RangeWrapper end = EndOfLastParagraph();
        Assert.Throws<NotSupportedException>(() => end.InsertCrossReference(CrossReferenceType.Footnote, CrossReferenceKind.ContentText, 0));
        Assert.Throws<NotSupportedException>(() => end.InsertCrossReference(CrossReferenceType.Heading, CrossReferenceKind.EntireCaption, 0));
        Assert.Throws<NotSupportedException>(() => _document.GetCrossReferenceItems(CrossReferenceType.Endnote));
        Assert.Throws<ArgumentOutOfRangeException>(() => end.InsertCrossReference(CrossReferenceType.Heading, CrossReferenceKind.ContentText, 0));
        Assert.Throws<ArgumentException>(() => end.InsertCrossReference("Nope", CrossReferenceKind.ContentText));
        Assert.Throws<ArgumentOutOfRangeException>(() => end.InsertCrossReference("Figure", CrossReferenceKind.EntireCaption, 0));
        Assert.Throws<ArgumentNullException>(() => end.InsertCrossReference((string)null, CrossReferenceKind.ContentText));
    }

    [Fact]
    public void AShapeCanBeAHyperlink()
    {
        _document.Content.InsertAfter("Hello world again");
        ShapeWrapper shape = _document.Shapes.AddShape(AutoShapeType.Rectangle, 20, 30, 80, 40, _document.Range(5, 5));
        Assert.Null(shape.Hyperlink);

        HyperlinkWrapper link = _document.Hyperlinks.Add(shape, "https://example.com/a", "Mark", "Tip");
        Assert.Equal(HyperlinkType.Shape, link.Type);
        Assert.Equal("https://example.com/a", link.Address);
        Assert.Equal("Mark", link.SubAddress);
        Assert.Equal("Tip", link.ScreenTip);
        Assert.Equal(link, shape.Hyperlink);
        Assert.Throws<NotSupportedException>(() => link.TextToDisplay = "x");
        Assert.Throws<NotSupportedException>(() => link.Range);
        Assert.Throws<ArgumentException>(() => _document.Hyperlinks.Add(shape, string.Empty));

        link.Address = "https://example.com/b";
        Assert.Equal("https://example.com/b", shape.Hyperlink.Address);
        Assert.Equal("Mark", shape.Hyperlink.SubAddress);
        Assert.Empty(DocumentIntegrity.CheckAll(_document.WordprocessingDocument));

        link.Delete();
        Assert.Null(shape.Hyperlink);
    }
}
