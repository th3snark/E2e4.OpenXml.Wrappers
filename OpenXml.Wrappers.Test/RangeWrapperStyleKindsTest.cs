using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>CharacterStyle</c>, <c>ParagraphStyle</c> and <c>TableStyle</c> of a range, which are read only in Word, and the default table style of a
/// document. The results are what Word 16 gave.
/// </summary>
public class RangeWrapperStyleKindsTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public RangeWrapperStyleKindsTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private RangeWrapper Content
    {
        get { return _document.Content; }
    }

    [Fact]
    public void TextWithoutACharacterStyleHasTheDefaultParagraphFont()
    {
        Content.InsertAfter("Plain text");

        Assert.Equal("Default Paragraph Font", Content.CharacterStyle);
        Assert.Equal("Default Paragraph Font", _document.Paragraphs[0].Range.CharacterStyle);
    }

    [Fact]
    public void ACharacterStyleIsFoundForTextThatHasItAndNullWhereTheyDiffer()
    {
        Content.InsertAfter("Plain text here");
        _document.Range(0, 5).Style = "Strong";

        Assert.Equal("Strong", _document.Range(0, 5).CharacterStyle);
        Assert.Equal("Strong", _document.Range(0, 3).CharacterStyle);
        Assert.Equal("Strong", _document.Range(2, 2).CharacterStyle);
        Assert.Equal("Default Paragraph Font", _document.Range(8, 8).CharacterStyle);
        Assert.Equal("Strong", _document.Range(0, 0).CharacterStyle);
        Assert.Equal("Strong", _document.Range(5, 5).CharacterStyle);
        Assert.Equal("Default Paragraph Font", _document.Range(6, 6).CharacterStyle);
        Assert.Null(_document.Range(3, 7).CharacterStyle);
        Assert.Null(_document.Paragraphs[0].Range.CharacterStyle);
        Assert.Equal("Strong", _document.Range(0, 5).CharacterStyleObject.NameLocal);
    }

    [Fact]
    public void SpacesAtTheEndOfARangeAreNotLookedAtForTheCharacterStyle()
    {
        Content.InsertAfter("AAAABBBBCCCC DDDD");
        _document.Range(8, 12).Style = "Emphasis";

        Assert.Equal("Emphasis", _document.Range(11, 13).CharacterStyle);
        Assert.Null(_document.Range(11, 14).CharacterStyle);
        Assert.Equal("Default Paragraph Font", _document.Range(12, 13).CharacterStyle);
    }

    [Fact]
    public void ACollapsedRangeAtTheStartOfAParagraphTakesTheCharacterStyleFromTheTextAfterIt()
    {
        Content.InsertAfter("One");
        Content.InsertParagraphAfter();
        Content.InsertAfter("Two");
        _document.Range(4, 7).Style = "Emphasis";

        Assert.Equal("Emphasis", _document.Range(4, 4).CharacterStyle);
        Assert.Equal("Default Paragraph Font", _document.Range(3, 3).CharacterStyle);
    }

    [Fact]
    public void AParagraphHasItsStyleAndTheCharacterStyleIsNotItsStyle()
    {
        Content.InsertAfter("Heading");
        _document.Paragraphs[0].Range.Style = "Heading 1";

        Assert.Equal("Heading 1", _document.Paragraphs[0].Range.ParagraphStyle, ignoreCase: true);
        Assert.Equal("Default Paragraph Font", _document.Paragraphs[0].Range.CharacterStyle);
        Assert.Equal("Normal", new RangeWrapper(new Paragraph(new Run(new Text("x")))).ParagraphStyle);
    }

    [Fact]
    public void ParagraphStyleIsTheParagraphStyleEvenForACharacterStyledRange()
    {
        Content.InsertAfter("Plain text here");
        _document.Range(0, 5).Style = "Strong";

        Assert.Equal("Normal", _document.Range(0, 3).ParagraphStyle);
        Assert.Equal("Strong", _document.Range(0, 3).Style);
        Assert.Equal("Normal", _document.Range(0, 3).ParagraphStyleObject.NameLocal);
    }

    [Fact]
    public void ParagraphsWithDifferentStylesHaveNoParagraphStyle()
    {
        Content.InsertAfter("One");
        Content.InsertParagraphAfter();
        Content.InsertAfter("Two");
        _document.Paragraphs[1].Range.Style = "Heading 1";

        Assert.Null(Content.ParagraphStyle);
        Assert.Null(Content.ParagraphStyleObject);
        Assert.Equal("Normal", _document.Paragraphs[0].Range.ParagraphStyle);
    }

    [Fact]
    public void ARangeInsideATableHasTheStyleOfTheTableAndAnyOtherTheDefault()
    {
        Content.InsertAfter("Before");
        Content.InsertParagraphAfter();
        TableWrapper table = _document.Tables.Add(_document.Paragraphs[_document.Paragraphs.Count - 1].Range, 2, 2);
        table.Style = "Table Grid";

        Assert.Equal("Table Grid", table.Cell(0, 0).Range.TableStyle);
        Assert.Equal("Table Grid", table.Rows[0].Range.TableStyle);
        Assert.Equal("Table Grid", table.Range.TableStyle);
        Assert.Equal("Table Grid", table.Cell(1, 1).Range.TableStyleObject.NameLocal);
        Assert.Equal("Table Normal", _document.Paragraphs[0].Range.TableStyle);
        Assert.Equal("Table Normal", Content.TableStyle);
    }

    [Fact]
    public void ATableWithoutAStyleHasTheDefaultTableStyle()
    {
        Content.InsertAfter("Before");
        Content.InsertParagraphAfter();
        TableWrapper table = _document.Tables.Add(_document.Paragraphs[_document.Paragraphs.Count - 1].Range, 1, 1);
        table.Style = null;

        Assert.Equal("Table Normal", table.Cell(0, 0).Range.TableStyle);
    }

    [Fact]
    public void ADocumentHasNoDefaultTableStyleUntilOneIsSet()
    {
        Assert.Null(_document.DefaultTableStyle);
        Assert.Null(_document.DefaultTableStyleObject);

        _document.SetDefaultTableStyle("Table Grid", false);

        Assert.Equal("Table Grid", _document.DefaultTableStyle);
        Assert.Equal("Table Grid", _document.DefaultTableStyleObject.NameLocal);

        _document.SetDefaultTableStyle(_document.Styles["Table Normal"], true);

        Assert.Equal("Table Normal", _document.DefaultTableStyle);
    }

    [Fact]
    public void TheDefaultTableStyleIsWrittenBetweenTheTabStopAndTheCharacterSpacing()
    {
        _document.SetDefaultTableStyle("Table Grid", false);

        string[] names = _document.WordprocessingDocument.MainDocumentPart.DocumentSettingsPart.Settings.ChildElements.Select(e => e.LocalName).ToArray();
        int style = Array.IndexOf(names, "defaultTableStyle");
        Assert.True(style > Array.IndexOf(names, "defaultTabStop"));
        Assert.True(style < Array.IndexOf(names, "characterSpacingControl"));
        Assert.Equal(1, names.Count(n => n == "defaultTableStyle"));
    }

    [Fact]
    public void AStyleThatIsNotATableStyleCannotBeTheDefault()
    {
        Assert.Throws<ArgumentException>(() => _document.SetDefaultTableStyle("Heading 1", false));
        Assert.Throws<ArgumentException>(() => _document.SetDefaultTableStyle("No Such Style", false));
        Assert.Throws<ArgumentNullException>(() => _document.SetDefaultTableStyle((string)null, false));
        Assert.Null(_document.DefaultTableStyle);
    }
}
