using System;
using System.Globalization;
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
/// The case of text, symbols, alignment tabs, the date, and the information about where a range is. The numbers are those Word gives for the
/// same content (found by running the same calls in Word 16), and the markup is the markup it saved.
/// </summary>
public class RangeWrapperContentTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public RangeWrapperContentTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.WordprocessingDocument.MainDocumentPart.Document.Body.RemoveAllChildren<Paragraph>();
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static RangeWrapper At(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    private Body AddText(string text)
    {
        Body body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        body.InsertAt(new Paragraph(Plain(text)), 0);
        return body;
    }

    [Theory]
    [InlineData("hello wORLD. second sentence here", CharacterCase.LowerCase, "hello world. second sentence here", CharacterCase.LowerCase)]
    [InlineData("hello wORLD. second sentence here", CharacterCase.UpperCase, "HELLO WORLD. SECOND SENTENCE HERE", CharacterCase.UpperCase)]
    [InlineData("hello wORLD. second sentence here", CharacterCase.TitleWord, "Hello World. Second Sentence Here", CharacterCase.TitleWord)]
    [InlineData("hello wORLD. second sentence here", CharacterCase.TitleSentence, "Hello world. Second sentence here", CharacterCase.TitleSentence)]
    [InlineData("hello wORLD. second sentence here", CharacterCase.ToggleCase, "HELLO World. SECOND SENTENCE HERE", CharacterCase.TitleWord)]
    [InlineData("ONE two THREE", CharacterCase.LowerCase, "one two three", CharacterCase.LowerCase)]
    [InlineData("ONE two THREE", CharacterCase.UpperCase, "ONE TWO THREE", CharacterCase.UpperCase)]
    [InlineData("ONE two THREE", CharacterCase.TitleWord, "ONE Two THREE", CharacterCase.TitleWord)]
    [InlineData("ONE two THREE", CharacterCase.TitleSentence, "ONE two THREE", CharacterCase.ToggleCase)]
    [InlineData("ONE two THREE", CharacterCase.ToggleCase, "one TWO three", CharacterCase.ToggleCase)]
    [InlineData("x", CharacterCase.TitleWord, "X", CharacterCase.UpperCase)]
    [InlineData("x", CharacterCase.TitleSentence, "X", CharacterCase.UpperCase)]
    [InlineData("HELLO WORLD", CharacterCase.TitleWord, "Hello World", CharacterCase.TitleWord)]
    [InlineData("HELLO WORLD", CharacterCase.TitleSentence, "Hello world", CharacterCase.TitleSentence)]
    [InlineData("Mixed Case Words here. ANOTHER one", CharacterCase.TitleWord, "Mixed Case Words Here. ANOTHER One", CharacterCase.TitleWord)]
    [InlineData("Mixed Case Words here. ANOTHER one", CharacterCase.TitleSentence, "Mixed Case Words here. ANOTHER one", CharacterCase.ToggleCase)]
    [InlineData("iPhone and McDonald i think", CharacterCase.TitleWord, "Iphone And Mcdonald I Think", CharacterCase.TitleWord)]
    [InlineData("iPhone and McDonald i think", CharacterCase.TitleSentence, "Iphone and mcdonald i think", CharacterCase.TitleSentence)]
    [InlineData("it's a well-known fact. dr. smith", CharacterCase.TitleWord, "It's A Well-Known Fact. Dr. Smith", CharacterCase.TitleWord)]
    [InlineData("it's a well-known fact. dr. smith", CharacterCase.TitleSentence, "It's a well-known fact. Dr. Smith", CharacterCase.TitleSentence)]
    [InlineData("A b C d", CharacterCase.TitleWord, "A B C D", CharacterCase.UpperCase)]
    [InlineData("A b C d", CharacterCase.TitleSentence, "A b C d", CharacterCase.ToggleCase)]
    public void CaseChangesTextAsWordDoes(string text, CharacterCase set, string expected, CharacterCase read)
    {
        Body body = AddText(text);
        RangeWrapper range = At(body, 0, text.Length);

        range.Case = set;

        Assert.Equal(expected, At(body, 0, text.Length).Text);
        Assert.Equal(read, range.Case);
    }

    [Fact]
    public void TheCaseOfTextIsReadFromItsLetters()
    {
        Body body = AddText("abc DEF");

        Assert.Equal(CharacterCase.LowerCase, At(body, 0, 3).Case);
        Assert.Equal(CharacterCase.UpperCase, At(body, 4, 7).Case);
        Assert.Equal(CharacterCase.ToggleCase, At(body, 0, 7).Case);
        Assert.Equal(CharacterCase.LowerCase, At(body, 2, 2).Case);
    }

    [Fact]
    public void ChangingTheCaseKeepsTheFormattingAndTheRuns()
    {
        Body body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        body.InsertAt(new Paragraph(new Run(new RunProperties(new Bold()), new Text("abc")), Plain(" def")), 0);

        At(body, 0, 7).Case = CharacterCase.UpperCase;

        Paragraph paragraph = body.Elements<Paragraph>().First();
        Assert.Equal("ABC DEF", paragraph.InnerText);
        Assert.NotNull(paragraph.Elements<Run>().First().RunProperties.Bold);
        Assert.Equal(2, paragraph.Elements<Run>().Count());
    }

    [Fact]
    public void ChangingTheCaseOfPartOfARunChangesOnlyThatPart()
    {
        Body body = AddText("abcdef");

        At(body, 2, 4).Case = CharacterCase.UpperCase;

        Assert.Equal("abCDef", body.Elements<Paragraph>().First().InnerText);
    }

    [Fact]
    public void AnAnsiSymbolIsASymElement()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertSymbol(171, "Symbol");

        Paragraph paragraph = body.Elements<Paragraph>().First();
        Assert.Equal(3, paragraph.Elements<Run>().Count());
        SymbolChar symbol = paragraph.Elements<Run>().ElementAt(1).GetFirstChild<SymbolChar>();
        Assert.Equal("Symbol", symbol.Font.Value);
        Assert.Equal("F0AB", symbol.Char.Value);
        Assert.Null(paragraph.Elements<Run>().ElementAt(1).RunProperties);
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void AUnicodeSymbolIsTheCharacterInTheFont()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertSymbol(0x20AC, "Calibri", true);

        Run run = body.Elements<Paragraph>().First().Elements<Run>().ElementAt(1);
        Assert.Equal("€", run.InnerText);
        Assert.Equal("Calibri", run.RunProperties.RunFonts.Ascii.Value);
        Assert.Equal("Calibri", run.RunProperties.RunFonts.HighAnsi.Value);
        Assert.Equal("Calibri", run.RunProperties.RunFonts.ComplexScript.Value);
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void TheComplexScriptFontIsLeftOutWhenTheTextHasItAnyway()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertSymbol(0xA9, "Arial", true);

        RunFonts fonts = body.Elements<Paragraph>().First().Elements<Run>().ElementAt(1).RunProperties.RunFonts;
        Assert.Equal("Arial", fonts.Ascii.Value);
        Assert.Equal("Arial", fonts.HighAnsi.Value);
        Assert.Null(fonts.ComplexScript);
    }

    [Fact]
    public void ASymbolTakesTheFormattingOfTheTextAroundIt()
    {
        Body body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        body.InsertAt(new Paragraph(new Run(new RunProperties(new Bold()), new Text("ab"))), 0);

        At(body, 1, 1).InsertSymbol(0xA9, "Arial", true);

        Run run = body.Elements<Paragraph>().First().Elements<Run>().ElementAt(1);
        Assert.NotNull(run.RunProperties.Bold);
        Assert.NotNull(run.RunProperties.RunFonts);
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void ASymbolReplacesTheRange()
    {
        Body body = AddText("ab");
        RangeWrapper range = At(body, 0, 1);

        range.InsertSymbol(0xA9, "Times New Roman", true);

        Assert.Equal("©b", body.Elements<Paragraph>().First().InnerText);
    }

    [Fact]
    public void ASymbolWithoutAFontIsTheCharacterInTheFontOfTheText()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertSymbol(0x20AC);

        Assert.Equal("a€b", body.Elements<Paragraph>().First().InnerText);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(0x110000)]
    [InlineData(0xD800)]
    public void ASymbolMustBeACharacter(int number)
    {
        Body body = AddText("ab");

        Assert.Throws<ArgumentOutOfRangeException>(() => At(body, 1, 1).InsertSymbol(number, "Arial", true));
    }

    [Fact]
    public void AnAnsiSymbolHasANumberFromZeroTo255()
    {
        Body body = AddText("ab");

        Assert.Throws<ArgumentOutOfRangeException>(() => At(body, 1, 1).InsertSymbol(300, "Symbol"));
    }

    [Theory]
    [InlineData(AlignmentTabAlignment.Left, AlignmentTabRelative.Margin, "margin", "left")]
    [InlineData(AlignmentTabAlignment.Center, AlignmentTabRelative.Indent, "indent", "center")]
    [InlineData(AlignmentTabAlignment.Right, AlignmentTabRelative.Margin, "margin", "right")]
    [InlineData(AlignmentTabAlignment.Right, AlignmentTabRelative.Indent, "indent", "right")]
    [InlineData(AlignmentTabAlignment.Left, AlignmentTabRelative.Indent, "indent", "left")]
    [InlineData(AlignmentTabAlignment.Center, AlignmentTabRelative.Margin, "margin", "center")]
    public void AnAlignmentTabIsAPositionalTab(AlignmentTabAlignment alignment, AlignmentTabRelative relative, string relativeTo, string aligned)
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertAlignmentTab(alignment, relative);

        Paragraph paragraph = body.Elements<Paragraph>().First();
        Assert.Equal(3, paragraph.Elements<Run>().Count());
        PositionalTab tab = paragraph.Elements<Run>().ElementAt(1).GetFirstChild<PositionalTab>();
        Assert.Equal(relativeTo, tab.RelativeTo.InnerText);
        Assert.Equal(aligned, tab.Alignment.InnerText);
        Assert.Equal("none", tab.Leader.InnerText);
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void AnAlignmentTabIsRelativeToTheMarginsAndTakesOnePosition()
    {
        Body body = AddText("ab");
        int before = new RangeWrapper(body).End;

        At(body, 1, 1).InsertAlignmentTab(AlignmentTabAlignment.Right);

        Assert.Equal("margin", body.Descendants<PositionalTab>().Single().RelativeTo.InnerText);
        Assert.Equal(before + 1, new RangeWrapper(body).End);
        Assert.Equal("ab", At(body, 0, 3).Text);
    }

    [Fact]
    public void TheDateIsAFieldWithAResultThatIsNotProofed()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertDateTime("yyyy-MM-dd");

        Paragraph paragraph = body.Elements<Paragraph>().First();
        Run[] runs = paragraph.Elements<Run>().ToArray();
        Assert.Equal(7, runs.Length);
        Assert.Equal(FieldCharValues.Begin, runs[1].GetFirstChild<FieldChar>().FieldCharType.Value);
        Assert.Equal(" DATE \\@ \"yyyy-MM-dd\" ", runs[2].GetFirstChild<FieldCode>().Text);
        Assert.Equal(FieldCharValues.Separate, runs[3].GetFirstChild<FieldChar>().FieldCharType.Value);
        Assert.NotNull(runs[4].RunProperties.NoProof);
        Assert.Contains(runs[4].InnerText, new[] { DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.CurrentCulture), DateTime.Now.AddMinutes(-1).ToString("yyyy-MM-dd", CultureInfo.CurrentCulture) });
        Assert.Equal(FieldCharValues.End, runs[5].GetFirstChild<FieldChar>().FieldCharType.Value);
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    [Fact]
    public void TheDateWithoutAFormatIsTheShortDateOfTheCulture()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertDateTime();

        Assert.Equal(" DATE \\@ \"dd/MM/yyyy\" ", body.Descendants<FieldCode>().Single().Text);
    }

    [Fact]
    public void TheDateCanBeText()
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertDateTime("yyyy", false);

        Assert.Equal("a" + DateTime.Now.ToString("yyyy", CultureInfo.CurrentCulture) + "b", body.Elements<Paragraph>().First().InnerText);
        Assert.Empty(body.Descendants<FieldCode>());
    }

    [Theory]
    [InlineData("d MMMM yyyy", "d MMMM yyyy")]
    [InlineData("HH:mm", "HH:mm")]
    [InlineData("h:mm am/pm", "h:mm tt")]
    [InlineData("d", "d")]
    public void TheDatePictureIsWordsLetters(string picture, string net)
    {
        Body body = AddText("ab");

        At(body, 1, 1).InsertDateTime(picture, false);

        string expected = DateTime.Now.ToString(net.Length == 1 ? "%" + net : net, CultureInfo.CurrentCulture);
        string actual = body.Elements<Paragraph>().First().InnerText;
        Assert.StartsWith("a", actual, StringComparison.Ordinal);
        Assert.Equal(expected.Length, actual.Length - 2);
    }

    [Fact]
    public void TheDateFieldTakesTheFormattingOfTheTextAroundIt()
    {
        Body body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        body.InsertAt(new Paragraph(new Run(new RunProperties(new Bold()), new Text("ab"))), 0);

        At(body, 1, 1).InsertDateTime("yyyy");

        Run[] runs = body.Elements<Paragraph>().First().Elements<Run>().ToArray();
        Assert.All(runs, r => Assert.NotNull(r.RunProperties?.Bold));
        Assert.NotNull(runs[4].RunProperties.NoProof);
        Assert.Empty(DocumentIntegrity.SchemaErrors(_document.WordprocessingDocument));
    }

    private static Body TableDocument()
    {
        //"Before", a table of three rows and two columns with a in the first cell and b in the cell of the second row and column, and "After".
        Table table = new Table();
        for (int row = 1; row <= 3; row++)
        {
            table.AppendChild(new TableRow(
                new TableCell(new Paragraph(row == 1 ? Plain("a") : new Run())),
                new TableCell(new Paragraph(row == 2 ? Plain("b") : new Run()))));
        }
        return new Body(new Paragraph(Plain("Before")), table, new Paragraph(Plain("After")));
    }

    [Theory]
    [InlineData(2, 3, false, -1, -1, -1, -1, -1, -1)]
    [InlineData(7, 9, true, 1, 1, 3, 1, 1, 2)]
    [InlineData(12, 14, true, 2, 2, 3, 2, 2, 2)]
    [InlineData(11, 15, true, 2, 2, 3, 1, 2, 2)]
    [InlineData(7, 18, true, 1, 3, 3, 1, 2, 2)]
    [InlineData(9, 17, true, 1, 3, 3, 2, 2, 2)]
    [InlineData(0, 24, false, -1, -1, -1, -1, -1, -1)]
    [InlineData(21, 23, false, -1, -1, -1, -1, -1, -1)]
    [InlineData(11, 11, true, 2, 2, 3, 1, 1, 2)]
    [InlineData(10, 11, true, 1, 1, 3, 3, 3, 2)]
    public void InformationAboutTablesIsWhatWordSays(int start, int end, bool inTable, int startRow, int endRow, int rows, int startColumn, int endColumn, int columns)
    {
        RangeWrapper range = At(TableDocument(), start, end);

        Assert.Equal(inTable, range.Information(InformationType.WithInTable));
        Assert.Equal(startRow, range.Information(InformationType.StartOfRangeRowNumber));
        Assert.Equal(endRow, range.Information(InformationType.EndOfRangeRowNumber));
        Assert.Equal(rows, range.Information(InformationType.MaximumNumberOfRows));
        Assert.Equal(startColumn, range.Information(InformationType.StartOfRangeColumnNumber));
        Assert.Equal(endColumn, range.Information(InformationType.EndOfRangeColumnNumber));
        Assert.Equal(columns, range.Information(InformationType.MaximumNumberOfColumns));
    }

    [Fact]
    public void InformationAboutWhereTheStoryIsIsFalseInTheMainText()
    {
        RangeWrapper range = At(TableDocument(), 2, 3);

        Assert.Equal(false, range.Information(InformationType.InHeaderFooter));
        Assert.Equal(false, range.Information(InformationType.InCommentPane));
        Assert.Equal(false, range.Information(InformationType.InFootnote));
        Assert.Equal(false, range.Information(InformationType.InEndnote));
        Assert.Equal(-1, range.Information(InformationType.HeaderFooterType));
        Assert.Throws<ArgumentOutOfRangeException>(() => range.Information((InformationType)1));
    }

    [Fact]
    public void InformationTellsAHeaderFromTheMainText()
    {
        HeaderPart part = _document.WordprocessingDocument.MainDocumentPart.AddNewPart<HeaderPart>();
        part.Header = new Header(new Paragraph(Plain("Head")));
        string id = _document.WordprocessingDocument.MainDocumentPart.GetIdOfPart(part);
        Body body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        body.AppendChild(new SectionProperties(new HeaderReference() { Type = HeaderFooterValues.First, Id = id }));

        RangeWrapper range = new RangeWrapper(part.Header);

        Assert.Equal(true, range.Information(InformationType.InHeaderFooter));
        Assert.Equal(2, range.Information(InformationType.HeaderFooterType));
    }

    [Theory]
    [InlineData(0, 0, 0, 0)]
    [InlineData(2, 2, 1, 1)]
    [InlineData(3, 3, 1, 1)]
    [InlineData(5, 5, 1, 1)]
    [InlineData(6, 6, 0, 1)]
    [InlineData(2, 5, 1, 1)]
    [InlineData(9, 9, 2, 2)]
    [InlineData(10, 10, 3, 3)]
    [InlineData(11, 11, 3, 3)]
    [InlineData(12, 12, 2, 3)]
    [InlineData(15, 15, 0, 3)]
    [InlineData(0, 20, 0, 0)]
    [InlineData(10, 11, 3, 3)]
    [InlineData(4, 9, 1, 1)]
    public void BookmarksAreNumberedInTheOrderTheyStart(int start, int end, int id, int previous)
    {
        Body body = AddText("0123456789ABCDEFGHIJ");
        RangeWrapper whole = new RangeWrapper(body);
        _document.Content.Bookmarks.Add("one", At(body, 2, 5));
        _document.Content.Bookmarks.Add("two", At(body, 8, 12));
        _document.Content.Bookmarks.Add("three", At(body, 10, 11));
        Assert.NotNull(whole);

        RangeWrapper range = At(body, start, end);

        Assert.Equal(id, range.BookmarkID);
        Assert.Equal(previous, range.PreviousBookmarkID);
    }

    private static Body TwoSections()
    {
        return new Body(
            new Paragraph(new ParagraphProperties(new SectionProperties()), Plain("0123456789")),
            new Paragraph(Plain("ABCDEFGHIJ")));
    }

    [Theory]
    [InlineData(0, 5, 1, 0)]
    [InlineData(0, 22, 2, 0)]
    [InlineData(5, 15, 2, 0)]
    [InlineData(11, 15, 1, 1)]
    [InlineData(11, 11, 1, 1)]
    [InlineData(10, 10, 1, 0)]
    [InlineData(0, 11, 1, 0)]
    public void ARangeIsInTheSectionsItTouches(int start, int end, int count, int firstIndex)
    {
        SectionWrapperCollection sections = At(TwoSections(), start, end).Sections;

        Assert.Equal(count, sections.Count);
        Assert.Equal(firstIndex, sections.First.Index);
    }

    [Fact]
    public void SectionsAreOnlyCountedInTheMainText()
    {
        Header header = new Header(new Paragraph(Plain("Head")));

        Assert.Throws<NotSupportedException>(() => new RangeWrapper(header).Sections);
    }
}
