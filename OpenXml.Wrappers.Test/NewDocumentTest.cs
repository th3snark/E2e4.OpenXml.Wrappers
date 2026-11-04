using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class NewDocumentTest
{
    private static T InCulture<T>(string name, Func<T> action)
    {
        CultureInfo culture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = name.Length == 0 ? CultureInfo.InvariantCulture : CultureInfo.GetCultureInfo(name);
        try
        {
            return action();
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
        }
    }

    private static MemoryStream CreateDocument(Action<WordprocessingDocumentWrapper> act = null)
    {
        MemoryStream stream = new MemoryStream();
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
        {
            act?.Invoke(document);
        }
        stream.Position = 0;
        return stream;
    }

    private static string[] EntryNames(Stream stream)
    {
        stream.Position = 0;
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
        {
            return archive.Entries.Select(e => e.FullName).ToArray();
        }
    }

    private static string ReadEntry(Stream stream, string name)
    {
        stream.Position = 0;
        using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
        using (StreamReader reader = new StreamReader(archive.GetEntry(name).Open()))
        {
            return reader.ReadToEnd();
        }
    }

    [Theory]
    [InlineData("en-GB", 11906u, 16838u, 708u, "en-GB")]
    [InlineData("en-US", 12240u, 15840u, 720u, "en-US")]
    [InlineData("de-DE", 11906u, 16838u, 708u, "de-DE")]
    [InlineData("en-CA", 12240u, 15840u, 708u, "en-CA")] //Letter paper, metric measurements
    [InlineData("", 12240u, 15840u, 720u, "en-US")] //The invariant culture has no region.
    [InlineData("en", 12240u, 15840u, 720u, "en-US")] //A neutral culture stands for its main region.
    public void ThePageAndLanguageComeFromTheCurrentCulture(string culture, uint width, uint height, uint distance, string language)
    {
        using (MemoryStream stream = InCulture(culture, () => CreateDocument()))
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, false))
        {
            MainDocumentPart main = document.MainDocumentPart;
            SectionProperties section = main.Document.Body.Elements<SectionProperties>().Single();
            PageSize size = section.GetFirstChild<PageSize>();
            PageMargin margin = section.GetFirstChild<PageMargin>();

            Assert.Equal(width, size.Width.Value);
            Assert.Equal(height, size.Height.Value);
            Assert.Equal(1440, margin.Top.Value);
            Assert.Equal(1440u, margin.Left.Value);
            Assert.Equal(distance, margin.Header.Value);
            Assert.Equal(distance, margin.Footer.Value);
            Assert.Equal(distance.ToString(CultureInfo.InvariantCulture), section.GetFirstChild<Columns>().Space.Value);
            Assert.Equal(language, main.DocumentSettingsPart.Settings.GetFirstChild<ThemeFontLanguages>().Val.Value);
            Assert.Equal(language, main.StyleDefinitionsPart.Styles.DocDefaults.RunPropertiesDefault.RunPropertiesBaseStyle.GetFirstChild<Languages>().Val.Value);
        }
    }

    [Fact]
    public void TheSeparatorsComeFromTheCurrentCulture()
    {
        CultureInfo german = CultureInfo.GetCultureInfo("de-DE");
        using (MemoryStream stream = InCulture("de-DE", () => CreateDocument()))
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, false))
        {
            Settings settings = document.MainDocumentPart.DocumentSettingsPart.Settings;
            Assert.Equal(german.NumberFormat.NumberDecimalSeparator, settings.GetFirstChild<DecimalSymbol>().Val.Value);
            Assert.Equal(german.TextInfo.ListSeparator, settings.GetFirstChild<ListSeparator>().Val.Value);
        }
    }

    [Fact]
    public void ANewDocumentHasThePartsWordWritesAndItsTextUsesTheTheme()
    {
        using (MemoryStream stream = CreateDocument(document =>
        {
            document.Content.Text = "Text";
            Assert.Equal("Aptos", document.Content.Font.Name);
            Assert.Equal(12, document.Content.Font.Size);
            Assert.Equal(8, document.Content.ParagraphFormat.SpaceAfter);
            Assert.Equal("Aptos Display", document.Styles["Heading 1"].Font.Name);
        }))
        {
            string[] names = EntryNames(stream);
            foreach (string name in new[] { "word/styles.xml", "word/settings.xml", "word/fontTable.xml", "word/webSettings.xml", "word/theme/theme1.xml", "docProps/core.xml", "docProps/app.xml" })
            {
                Assert.Contains(name, names);
            }

            stream.Position = 0;
            using (WordprocessingDocument document = WordprocessingDocument.Open(stream, false))
            {
                Assert.Empty(DocumentIntegrity.SchemaErrors(document));
            }
        }
    }

    [Fact]
    public void ANewDocumentDoesNotClaimAnAuthorOrWord()
    {
        using (MemoryStream stream = CreateDocument())
        {
            string core = ReadEntry(stream, "docProps/core.xml");
            string application = ReadEntry(stream, "docProps/app.xml");

            Assert.Contains(">1<", core);
            Assert.DoesNotContain("creator", core);
            Assert.DoesNotContain("Application", application);
            Assert.DoesNotContain("Words", application);
        }
    }

    [Fact]
    public void EachDocumentGetsItsOwnDocumentIds()
    {
        using (MemoryStream first = CreateDocument())
        using (MemoryStream second = CreateDocument())
        {
            Assert.NotEqual(ReadEntry(first, "word/settings.xml"), ReadEntry(second, "word/settings.xml"));
        }
    }

    [Fact]
    public void InitializingTheMainDocumentKeepsThePartsThatAreThere()
    {
        using (MemoryStream stream = CreateDocument(document =>
        {
            document.Styles.Add("Mine", StyleType.Paragraph);
            document.Content.Text = "Text";

            document.InitializeMainDocumentPart();

            Assert.True(document.Styles.Exists("Mine"));
            Assert.Equal("\r", document.Content.Text);
        }))
        {
            Assert.Contains("word/styles.xml", EntryNames(stream));
        }
    }

    [Fact]
    public void CorePropertiesAreKeptWhereWordKeepsThem()
    {
        using (MemoryStream stream = CreateDocument(document => document.BuiltInDocumentProperties.Title = "A title"))
        {
            string[] names = EntryNames(stream);
            Assert.Contains("docProps/core.xml", names);
            Assert.DoesNotContain(names, n => n.EndsWith(".psmdcp", StringComparison.Ordinal));
            Assert.Contains("A title", ReadEntry(stream, "docProps/core.xml"));
        }
    }

    [Fact]
    public void CorePropertiesOfADocumentThatHasNoneAreKeptWhereWordKeepsThem()
    {
        using (MemoryStream stream = new MemoryStream())
        {
            using (WordprocessingDocument plain = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                plain.AddMainDocumentPart().Document = new Document(new Body(new Paragraph()));
            }

            stream.Position = 0;
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(stream, true))
            {
                document.BuiltInDocumentProperties.Title = "A title";
            }

            string[] names = EntryNames(stream);
            Assert.Contains("docProps/core.xml", names);
            Assert.DoesNotContain(names, n => n.EndsWith(".psmdcp", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void IndentationKeepsWordsAttributeNamesWhenADocumentIsLoadedAndSaved()
    {
        using (MemoryStream stream = CreateDocument(document =>
        {
            document.Content.Text = "Text";
            document.Paragraphs[0].LeftIndent = 18;
            document.Paragraphs[0].RightIndent = 9;
        }))
        {
            stream.Position = 0;
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(stream, true))
            {
                Assert.Equal(18, document.Paragraphs[0].LeftIndent);
                document.Paragraphs[0].Range.InsertAfter(" more");
            }

            string xml = ReadEntry(stream, "word/document.xml");
            Assert.Contains("w:left=\"360\"", xml);
            Assert.Contains("w:right=\"180\"", xml);
            Assert.DoesNotContain("w:start=", xml);
            Assert.DoesNotContain("w:end=", xml);

            stream.Position = 0;
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(stream, false))
            {
                Assert.Equal(18, document.Paragraphs[0].LeftIndent);
                Assert.Equal(9, document.Paragraphs[0].RightIndent);
            }
        }
    }

    [Theory]
    [InlineData("en-GB", 4513, 9026)]
    [InlineData("en-US", 4680, 9360)]
    public void TheTabStopsOfHeadersAndFootersFollowTheTextWidth(string culture, int centre, int right)
    {
        InCulture(culture, () =>
        {
            using (MemoryStream stream = CreateDocument(document =>
            {
                foreach (BuiltInStyle id in new[] { BuiltInStyle.Header, BuiltInStyle.Footer })
                {
                    TabStop[] tabs = document.Styles[id].Style.StyleParagraphProperties.Tabs.Elements<TabStop>().ToArray();
                    Assert.Equal(centre, tabs[0].Position.Value);
                    Assert.Equal(right, tabs[1].Position.Value);
                }
            }))
            {
                return stream.Length;
            }
        });
    }

    [Fact]
    public void AddingACommentAddsTheStylesItUses()
    {
        using (MemoryStream stream = CreateDocument(document =>
        {
            document.Content.Text = "Commented";
            document.Comments.Add(document.Range(0, 4), "A comment");
        }))
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, false))
        {
            Assert.Empty(DocumentIntegrity.Check(document, DocumentIntegrity.Styles));
            string[] ids = document.MainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>().Select(s => s.StyleId.Value).ToArray();
            Assert.Equal(new[] { "CommentReference", "CommentText", "CommentTextChar", "CommentSubject", "CommentSubjectChar" }, ids.Skip(ids.Length - 5));
        }
    }
}
