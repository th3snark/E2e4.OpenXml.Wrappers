using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Ranges of characters in a whole document: bookmarks, comments, fields, styles, paragraphs and headers, and that the result is valid.
/// </summary>
public class RangeWrapperSpanDocumentTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");

    public void Dispose()
    {
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private WordprocessingDocumentWrapper CreateDocument(params string[] paragraphs)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(_path, true, true);
        document.Content.Text = string.Join("\r", paragraphs);
        return document;
    }

    private static void AssertValid(WordprocessingDocumentWrapper document)
    {
        Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document.WordprocessingDocument));
    }
    [Fact]
    public void ADocumentRangeIsASpanOfTheMainText()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Hello world", "Second"))
        {
            Assert.Equal("Hello world\rSecond\r", document.Content.Text);
            Assert.Equal("world", document.Range(6, 11).Text);
            Assert.Equal(19, document.Content.StoryLength);
            Assert.Throws<ArgumentOutOfRangeException>(() => document.Range(0, 100));
        }
    }

    [Fact]
    public void ASpanCanBeBookmarked()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Hello world", "Second"))
        {
            BookmarkWrapper bookmark = document.Bookmarks.Add("who", document.Range(6, 11));

            Assert.Equal("world", bookmark.Range.Text);
            Assert.Equal("world", document.Bookmarks["who"].Range.Text);
            Assert.Equal(6, bookmark.Range.Start);
            Assert.Equal(11, bookmark.Range.End);
            AssertValid(document);
        }
    }

    [Fact]
    public void ABookmarkCanBeAddedAtAnInsertionPoint()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Hello world"))
        {
            document.Bookmarks.Add("here", document.Range(5, 5));

            BookmarkWrapper bookmark = document.Bookmarks["here"];
            Assert.Equal(string.Empty, bookmark.Range.Text);
            bookmark.Range.InsertAfter(",");
            Assert.Equal("Hello, world\r", document.Content.Text);
            AssertValid(document);
        }
    }

    [Fact]
    public void ASpanCanHaveAComment()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Hello world"))
        {
            CommentWrapper comment = document.Comments.Add(document.Range(0, 5), "Greeting");

            Assert.Equal("Greeting", comment.Text.TrimEnd('\r'));
            Assert.Equal("Hello", comment.Scope.Text);
            AssertValid(document);
        }
    }

    [Fact]
    public void AFieldCanBeAddedInTheMiddleOfText()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Page of"))
        {
            RangeWrapper range = document.Range(4, 4);

            document.Fields.Add(range, FieldType.Page).Result = "1";

            Assert.Single(document.Fields);
            Assert.Equal("Page1 of\r", document.Content.Text);
            AssertValid(document);
        }
    }

    [Fact]
    public void AStyleCanBeAppliedToPartOfAParagraphAndToParagraphs()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Hello world", "Second"))
        {
            StyleDefinitionsPart part = document.WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            part.Styles = new Styles(
                new Style(new StyleName() { Val = "Strong" }) { Type = StyleValues.Character, StyleId = "Strong" },
                new Style(new StyleName() { Val = "heading 1" }) { Type = StyleValues.Paragraph, StyleId = "Heading1" });

            document.Range(6, 11).Style = "Strong";
            document.Range(14, 14).Style = "Heading 1";

            Assert.Equal("Strong", document.Range(6, 11).Style);
            Assert.Equal("Heading 1", document.Paragraphs[1].Range.Style);
            AssertValid(document);
        }
    }

    [Fact]
    public void InsertParagraphAfterASpanSplitsTheParagraph()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Hello world"))
        {
            RangeWrapper range = document.Range(5, 5);

            ParagraphWrapper second = range.InsertParagraphAfter();

            Assert.Equal(2, document.Paragraphs.Count);
            Assert.Equal(" world", second.Text);
            Assert.Equal("Hello\r world\r", document.Content.Text);
            Assert.Equal(1, range.End - range.Start);
            AssertValid(document);
        }
    }

    [Fact]
    public void ARangeInAHeaderCountsFromTheStartOfTheHeader()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("Body"))
        {
            RangeWrapper header = document.Sections[0].Headers[HeaderFooterIndex.Primary].Range;
            header.Text = "Report 2026";

            RangeWrapper year = header.Duplicate();
            year.SetRange(7, 11);
            year.Font.Bold = true;

            Assert.Equal("2026", year.Text);
            Assert.True(year.Font.Bold);
            Assert.Equal("Report ", header.Duplicate().Text.Substring(0, 7));
            AssertValid(document);
        }
    }

    [Fact]
    public void FormattingAndReplacingWordsKeepsTheDocumentValid()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("The quick brown fox jumps. Over the lazy dog.", "Second paragraph here."))
        {
            RangeWrapper content = document.Content;
            content.Words[1].Font.Bold = true;
            content.Words[2].Text = "red ";
            content.Sentences[1].Font.Italic = true;
            Assert.True(content.Find.Execute("lazy"));
            content.Font.Color = "C00000";

            Assert.Equal("The quick red fox jumps. Over the lazy dog.\rSecond paragraph here.\r", document.Content.Text);
            AssertValid(document);
        }
    }

    [Fact]
    public void DeletingASpanAcrossParagraphsKeepsTheDocumentValid()
    {
        using (WordprocessingDocumentWrapper document = CreateDocument("First one", "Second two", "Third three"))
        {
            document.Bookmarks.Add("inside", document.Range(12, 16));

            document.Range(6, 28).Delete();

            Assert.Equal("First hree\r", document.Content.Text);
            Assert.Equal(0, document.Bookmarks.Count);
            AssertValid(document);
        }
    }
}
