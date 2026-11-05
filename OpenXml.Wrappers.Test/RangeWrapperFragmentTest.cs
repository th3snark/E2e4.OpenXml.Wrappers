using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Copying content between ranges, files and XML. Every text and count is the one Word gives for the same content (found by running the same
/// calls in Word 16).
/// </summary>
public class RangeWrapperFragmentTest : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

    public RangeWrapperFragmentTest()
    {
        Directory.CreateDirectory(_folder);
    }

    public void Dispose()
    {
        Directory.Delete(_folder, true);
        GC.SuppressFinalize(this);
    }

    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Run Bold(string text)
    {
        return new Run(new RunProperties(new Bold()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static RangeWrapper At(OpenXmlElement story, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(story);
        range.SetRange(start, end);
        return range;
    }

    private static string Show(RangeWrapper range)
    {
        return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0}-{1}", range.Start, range.End);
    }

    private static string[] Paragraphs(OpenXmlElement story)
    {
        return story.Descendants<Paragraph>().Select(p => p.InnerText).ToArray();
    }

    private static string Runs(Paragraph paragraph)
    {
        return string.Join("|", paragraph.Elements<Run>().Select(r => (r.RunProperties?.Bold != null ? "*" : string.Empty) + r.InnerText));
    }

    [Fact]
    public void BoldTextCopiedToTheEndKeepsItsFormatting()
    {
        Body body = new Body(new Paragraph(Plain("abc "), Bold("def"), Plain(" ghi")));
        RangeWrapper target = At(body, 11, 11);

        target.FormattedText = At(body, 4, 7).FormattedText;

        Assert.Equal("abc |*def| ghi|*def", Runs(body.Elements<Paragraph>().Single()));
        Assert.Equal("11-14", Show(target));
    }

    [Fact]
    public void CopyingOverARangeReplacesIt()
    {
        Body body = new Body(new Paragraph(Plain("abc "), Bold("def"), Plain(" ghi")));
        RangeWrapper target = At(body, 0, 3);

        target.FormattedText = At(body, 4, 7).FormattedText;

        Assert.Equal("*def| |*def| ghi", Runs(body.Elements<Paragraph>().Single()));
        Assert.Equal("0-3", Show(target));
    }

    [Fact]
    public void CopyingAcrossAParagraphMarkBringsTheMarkAndLeavesTheEndHanging()
    {
        Body body = new Body(
            new Paragraph(Plain("aaa")),
            new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), Plain("bbb")),
            new Paragraph(Plain("ccc")));
        RangeWrapper target = At(body, 11, 11);

        target.FormattedText = At(body, 2, 6).FormattedText;

        Assert.Equal(new[] { "aaa", "bbb", "ccca", "bb" }, Paragraphs(body));
        Assert.Equal(1, body.Elements<Paragraph>().Count(p => p.ParagraphProperties?.Justification != null));
        Assert.Equal("11-15", Show(target));
    }

    [Fact]
    public void ACopiedParagraphWithItsMarkBringsItsProperties()
    {
        Body body = new Body(
            new Paragraph(Plain("aaa")),
            new Paragraph(new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" }), Plain("bbb")),
            new Paragraph(Plain("ccc")));
        RangeWrapper target = At(body, 0, 0);

        target.FormattedText = At(body, 4, 8).FormattedText;

        Assert.Equal(new[] { "bbb", "aaa", "bbb", "ccc" }, Paragraphs(body));
        Paragraph[] paragraphs = body.Elements<Paragraph>().ToArray();
        Assert.Equal("Heading1", paragraphs[0].ParagraphProperties.ParagraphStyleId.Val.Value);
        Assert.Null(paragraphs[1].ParagraphProperties);
        Assert.Equal("0-4", Show(target));
    }

    [Fact]
    public void CopiedTextInTheMiddleOfAParagraphSplitsItWhenItBringsAMark()
    {
        Body body = new Body(new Paragraph(Plain("xxyy")), new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Right }), Plain("bbb")));
        RangeWrapper target = At(body, 2, 2);

        target.FormattedText = At(body, 5, 9).FormattedText;

        Assert.Equal(new[] { "xxbbb", "yy", "bbb" }, Paragraphs(body));
        Assert.NotNull(body.Elements<Paragraph>().First().ParagraphProperties.Justification);
        Assert.Null(body.Elements<Paragraph>().ElementAt(1).ParagraphProperties);
    }

    [Fact]
    public void ACollapsedRangeCopiesNothing()
    {
        Body body = new Body(new Paragraph(Plain("abc def")));
        RangeWrapper target = At(body, 7, 7);

        target.FormattedText = At(body, 2, 2).FormattedText;

        Assert.Equal("abc def", body.InnerText);
        Assert.Equal("7-7", Show(target));
    }

    [Fact]
    public void AFieldIsCopiedWithItsResult()
    {
        Body body = new Body(new Paragraph(
            Plain("abc"),
            new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" PAGE ")),
            new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("1")),
            new Run(new FieldChar() { FieldCharType = FieldCharValues.End }),
            Plain(" def")));
        RangeWrapper target = At(body, 17, 17);

        target.FormattedText = At(body, 0, 17).FormattedText;

        Assert.Equal(2, body.Descendants<FieldCode>().Count());
        Assert.Equal("abc1 defabc1 def", body.Elements<Paragraph>().Single().InnerText.Replace(" PAGE ", string.Empty));
    }

    [Fact]
    public void BookmarksStayWithTheOriginalAndTheCopyHasNoneOfItsOwn()
    {
        Body body = new Body(new Paragraph(Plain("abc "), new BookmarkStart() { Id = "0", Name = "bk" }, Plain("def"), new BookmarkEnd() { Id = "0" }, Plain(" ghi")));
        RangeWrapper target = At(body, 11, 11);

        target.FormattedText = At(body, 2, 9).FormattedText;

        Assert.Single(body.Descendants<BookmarkStart>());
        Assert.Equal("abc def ghic def g", body.Elements<Paragraph>().Single().InnerText);
    }

    [Fact]
    public void ATableIsCopiedWhole()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph(Plain("x"))), new TableCell(new Paragraph())));
        Body body = new Body(new Paragraph(Plain("aaa")), table, new Paragraph(Plain("bbb")));
        int end = new RangeWrapper(body).End - 1;
        RangeWrapper target = At(body, end, end);

        target.FormattedText = new RangeWrapper(table).FormattedText;

        Assert.Equal(2, body.Descendants<Table>().Count());
        Assert.Equal(new[] { "aaa", "x", string.Empty, "bbb" }, body.Descendants<Paragraph>().Take(4).Select(p => p.InnerText).ToArray());
    }

    [Fact]
    public void StylesAreCopiedToAnotherDocument()
    {
        using (WordprocessingDocumentWrapper source = WordprocessingDocumentWrapper.Create(Path.Combine(_folder, "source.docx"), true, false))
        using (WordprocessingDocumentWrapper target = WordprocessingDocumentWrapper.Create(Path.Combine(_folder, "target.docx"), true, false))
        {
            Styles styles = source.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
            styles.AppendChild(new Style(new StyleName() { Val = "Parent" }) { Type = StyleValues.Paragraph, StyleId = "Parent", CustomStyle = true });
            styles.AppendChild(new Style(new StyleName() { Val = "Child" }, new BasedOn() { Val = "Parent" }) { Type = StyleValues.Paragraph, StyleId = "Child", CustomStyle = true });
            source.Content.InsertAfter("Head\rbody text");
            source.Paragraphs[0].Style = "Child";
            source.Range(5, 9).Font.Bold = true;
            target.Content.InsertAfter("target");
            Assert.Null(StyleOf(target, "Child"));

            RangeWrapper into = target.Range(0, 0);
            into.FormattedText = source.Range(0, 14).FormattedText;

            Assert.Equal("Head\rbody texttarget\r", target.Content.Text);
            Assert.Equal("Child", target.Paragraphs[0].Paragraph.ParagraphProperties.ParagraphStyleId.Val.Value);
            Assert.NotNull(StyleOf(target, "Child"));
            Assert.NotNull(StyleOf(target, "Parent"));
            Assert.Empty(DocumentIntegrity.SchemaErrors(target.WordprocessingDocument));
            Assert.Empty(DocumentIntegrity.Check(target.WordprocessingDocument, DocumentIntegrity.Styles));
        }
    }

    private static Style StyleOf(WordprocessingDocumentWrapper document, string id)
    {
        return document.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == id);
    }

    [Fact]
    public void ALinkIsCopiedToAnotherDocumentWithARelationshipOfItsOwn()
    {
        using (WordprocessingDocumentWrapper source = WordprocessingDocumentWrapper.Create(Path.Combine(_folder, "source.docx"), true, false))
        using (WordprocessingDocumentWrapper target = WordprocessingDocumentWrapper.Create(Path.Combine(_folder, "target.docx"), true, false))
        {
            MainDocumentPart main = source.WordprocessingDocument.MainDocumentPart;
            string id = main.AddHyperlinkRelationship(new Uri("https://example.com/"), true).Id;
            main.Document.Body.RemoveAllChildren<Paragraph>();
            main.Document.Body.InsertAt(new Paragraph(new Hyperlink(Plain("link")) { Id = id }, Plain(" text")), 0);
            target.Content.InsertAfter("target");
            target.WordprocessingDocument.MainDocumentPart.AddHyperlinkRelationship(new Uri("https://other.example/"), true);

            target.Range(0, 0).FormattedText = source.Content.FormattedText;

            Hyperlink copied = target.WordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Hyperlink>().Single();
            HyperlinkRelationship relationship = target.WordprocessingDocument.MainDocumentPart.HyperlinkRelationships.Single(r => r.Id == copied.Id.Value);
            Assert.Equal("https://example.com/", relationship.Uri.ToString());
            Assert.Empty(DocumentIntegrity.SchemaErrors(target.WordprocessingDocument));
            Assert.Empty(DocumentIntegrity.Check(target.WordprocessingDocument, DocumentIntegrity.Relationships));
        }
    }

    private string SourceFile(string text, string name = "source.docx")
    {
        string path = Path.Combine(_folder, name);
        using (WordprocessingDocumentWrapper source = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            source.Content.InsertAfter(text);
            if (text.Contains('\r', StringComparison.Ordinal))
            {
                source.Paragraphs[1].Format.Alignment = ParagraphAlignment.Center;
                source.Range(0, 2).Font.Bold = true;
            }
        }
        return path;
    }

    private static string TextOf(OpenXmlElement story, int paragraphs)
    {
        Assert.Equal(paragraphs, story.Descendants<Paragraph>().Count());
        return string.Join("|", story.Descendants<Paragraph>().Select(p => p.InnerText)) + "|";
    }

    [Theory]
    [InlineData(3, 3, true, "abcX1|Y2 two|Z3| def ghi|", 4, "3-3")]
    [InlineData(0, 3, true, "X1|Y2 two|Z3| def ghi|", 4, "0-0")]
    [InlineData(11, 11, true, "abc def ghiX1|Y2 two|Z3||", 4, "11-11")]
    [InlineData(3, 3, false, "abcX1|Y2 two|Z3 def ghi|", 3, "3-3")]
    [InlineData(0, 3, false, "X1|Y2 two|Z3 def ghi|", 3, "0-0")]
    public void AFileIsPutInAsWordDoes(int start, int end, bool withMark, string expected, int paragraphs, string range)
    {
        string file = SourceFile("X1\rY2 two\rZ3");
        Body body = new Body(new Paragraph(Plain("abc def ghi")));
        RangeWrapper target = At(body, start, end);

        if (withMark)
        {
            target.InsertFile(file);
        }
        else
        {
            target.ImportFragment(file);
        }

        Assert.Equal(expected, TextOf(body, paragraphs));
        Assert.Equal(range, Show(target));
    }

    [Theory]
    [InlineData(3, 3, true, "abcsingle| def ghi|", 2)]
    [InlineData(11, 11, true, "abc def ghisingle||", 2)]
    [InlineData(0, 0, true, "single|abc def ghi|", 2)]
    [InlineData(3, 3, false, "abcsingle def ghi|", 1)]
    [InlineData(0, 3, false, "single def ghi|", 1)]
    public void AFileOfOneParagraphIsPutInAsWordDoes(int start, int end, bool withMark, string expected, int paragraphs)
    {
        string file = SourceFile("single", "single.docx");
        Body body = new Body(new Paragraph(Plain("abc def ghi")));
        RangeWrapper target = At(body, start, end);

        if (withMark)
        {
            target.InsertFile(file);
        }
        else
        {
            target.ImportFragment(file);
        }

        Assert.Equal(expected, TextOf(body, paragraphs));
    }

    [Fact]
    public void AFileKeepsTheFormattingOfItsParagraphs()
    {
        string file = SourceFile("X1\rY2 two\rZ3");
        Body body = new Body(new Paragraph(Plain("abc def ghi")));

        At(body, 3, 3).InsertFile(file);

        Assert.Single(body.Descendants<Justification>());
        Assert.NotEmpty(body.Descendants<Bold>());
    }

    [Fact]
    public void ARangeIsSavedAsAFile()
    {
        Body body = new Body(new Paragraph(Plain("one")), new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), Plain("two three")), new Paragraph(Plain("four")));
        string file = Path.Combine(_folder, "export.docx");

        At(body, 2, 8).ExportFragment(file);

        using (WordprocessingDocument saved = WordprocessingDocument.Open(file, false))
        {
            Assert.Equal(new[] { "e", "two " }, Paragraphs(saved.MainDocumentPart.Document.Body));
            Assert.Empty(DocumentIntegrity.SchemaErrors(saved));
        }

        string whole = Path.Combine(_folder, "one.docx");
        At(body, 0, 3).ExportFragment(whole);
        using (WordprocessingDocument saved = WordprocessingDocument.Open(whole, false))
        {
            Assert.Equal(new[] { "one" }, Paragraphs(saved.MainDocumentPart.Document.Body));
        }
        Assert.Throws<NotSupportedException>(() => At(body, 0, 3).ExportFragment(whole, SaveFormat.XmlTemplate));
    }

    [Fact]
    public void WordOpenXMLIsAFlatPackageOfTheRange()
    {
        Body body = new Body(new Paragraph(Plain("one")), new Paragraph(Plain("two three")), new Paragraph(Plain("four")));

        string xml = At(body, 0, 8).WordOpenXML;

        Assert.Contains("<pkg:package", xml, StringComparison.Ordinal);
        Assert.Contains("pkg:name=\"/word/document.xml\"", xml, StringComparison.Ordinal);
        Assert.Contains("pkg:name=\"/word/styles.xml\"", xml, StringComparison.Ordinal);
        Assert.Contains(">two </w:t>", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("three", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void XMLFromARangeCanBeInsertedAgain()
    {
        Body body = new Body(new Paragraph(Plain("one")), new Paragraph(Plain("two three")));
        string xml = At(body, 2, 6).WordOpenXML;
        Body target = new Body(new Paragraph(Plain("target")));

        RangeWrapper range = At(target, 3, 3);
        range.InsertXML(xml);

        Assert.Equal("tare|tw|get|", TextOf(target, 3));
        Assert.Equal("3-3", Show(range));
    }

    [Fact]
    public void XMLThatIsNotAPackageIsRefused()
    {
        RangeWrapper range = At(new Body(new Paragraph(Plain("target"))), 3, 3);

        Assert.Throws<ArgumentException>(() => range.InsertXML("<not-a-package/>"));
        Assert.Throws<ArgumentException>(() => range.InsertXML("this is not xml"));
        Assert.Throws<ArgumentNullException>(() => range.InsertXML(null));
    }
}
