using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Character ranges over the real fixture documents: positions agree with the text, and splitting runs everywhere changes nothing visible.
/// </summary>
public class RangeWrapperSpanFixtureTest : IDisposable
{
    private readonly string _copy = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");

    public void Dispose()
    {
        File.Delete(_copy);
        GC.SuppressFinalize(this);
    }

    [Theory]
    [InlineData(TestFile.Document1)]
    [InlineData(TestFile.Document3)]
    public void ASpanOverAParagraphHasTheTextOfTheParagraph(string fixture)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(fixture))
        {
            foreach (ParagraphWrapper paragraph in document.Paragraphs.Where(p => !(p.Paragraph.Parent is DocumentFormat.OpenXml.Wordprocessing.TableCell)))
            {
                RangeWrapper range = paragraph.Range;
                Assert.Equal(range.Text, document.Range(range.Start, range.End).Text);
            }
        }
    }

    [Theory]
    [InlineData(TestFile.Document1)]
    [InlineData(TestFile.Document3)]
    public void TheCharactersOfTheDocumentAreTheCharactersOfItsText(string fixture)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(fixture))
        {
            RangeWrapper content = document.Content;
            string text = content.Text;

            //Each position is one character, except that the end of a cell and of a row are written as two characters, and that a comment
            //reference takes a position but shows no text.
            int doubled = text.Split('\a').Length - 1;
            int references = document.WordprocessingDocument.MainDocumentPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.CommentReference>().Count();
            Assert.Equal(text.Length - doubled + references, content.StoryLength);
            Assert.Equal(content.StoryLength, content.Characters.Count);
            Assert.Equal(text, string.Concat(content.Characters.Select(c => c.Text)));
            Assert.Equal(text, string.Concat(content.Words.Select(w => w.Text)));
            Assert.Equal(text, string.Concat(content.Sentences.Select(s => s.Text)));
        }
    }

    [Theory]
    [InlineData(TestFile.Document1)]
    [InlineData(TestFile.Document3)]
    public void FormattingEveryOtherCharacterLeavesTheTextAndAValidDocument(string fixture)
    {
        File.Copy(fixture, _copy);
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(_copy, true))
        {
            string before = document.Content.Text;
            int length = document.Content.StoryLength;

            for (int i = 0; i < length; i += 2)
            {
                document.Range(i, i + 1).Font.Bold = true;
            }

            Assert.Equal(before, document.Content.Text);
            Assert.Equal(length, document.Content.StoryLength);
            Assert.True(document.Range(0, 1).Font.Bold == true || document.Range(0, 1).Text.All(c => c == '\r'));
            Assert.DoesNotContain(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document.WordprocessingDocument)
, e => !e.Description.Contains("vbaProject", StringComparison.Ordinal));
        }
    }

    [Theory]
    [InlineData(TestFile.Document1)]
    [InlineData(TestFile.Document3)]
    public void ReplacingEveryLetterEWithTwoKeepsTheDocumentValid(string fixture)
    {
        File.Copy(fixture, _copy);
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(_copy, true))
        {
            string before = document.Content.Text;
            int count = before.Count(c => c == 'e' || c == 'E');

            document.Content.Find.Execute("e", replaceWith: "ee", replace: ReplaceMode.All);

            Assert.Equal(before.Length + count, document.Content.Text.Length);
            Assert.Equal(before.Replace("e", "ee", StringComparison.OrdinalIgnoreCase).Length, document.Content.Text.Length);
            Assert.DoesNotContain(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document.WordprocessingDocument)
, e => !e.Description.Contains("vbaProject", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void DeletingEveryThirdWordKeepsTheDocumentValid()
    {
        File.Copy(TestFile.Document3, _copy);
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(_copy, true))
        {
            int words = document.Content.Words.Count;
            for (int i = words - 1; i >= 0; i -= 3)
            {
                RangeWrapper word = document.Content.Words[i];
                if (!word.Text.Contains('\r', StringComparison.Ordinal) && !word.Text.Contains('\a', StringComparison.Ordinal))
                {
                    word.Delete();
                }
            }

            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document.WordprocessingDocument));
        }
    }
}
