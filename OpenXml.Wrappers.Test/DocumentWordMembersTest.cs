using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Members of the document, sections, headers and footers, page setup and properties that follow Word's object model.
/// </summary>
public class DocumentWordMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public DocumentWordMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "one two three";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        string[] errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010).Validate(_document.WordprocessingDocument).Select(e => e.Description + " at " + e.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, String.Join("; ", errors));
    }

    #region Sections

    [Fact]
    public void ASectionHasItsRange()
    {
        _document.Paragraphs.Add().Text = "second";
        _document.Paragraphs.Add().Text = "third";
        _document.Range(0, 0).InsertBreak(BreakType.SectionBreakNextPage);
        _document.Paragraphs[3].Range.InsertBreak(BreakType.SectionBreakContinuous);

        Assert.Equal(3, _document.Sections.Count);
        Assert.Equal("\r", _document.Sections[0].Range.Text);
        Assert.StartsWith("one two three\rsecond\r", _document.Sections[1].Range.Text, StringComparison.Ordinal);
        Assert.Equal("third\r", _document.Sections[2].Range.Text);
        Assert.Equal(_document.Sections[0], _document.Sections.First);
        Assert.Equal(_document.Sections[2], _document.Sections.Last);
    }

    [Fact]
    public void SectionsAddAddsABreakAndReturnsTheNewSection()
    {
        SectionWrapper added = _document.Sections.Add(null, SectionStart.EvenPage);

        Assert.Equal(2, _document.Sections.Count);
        Assert.Equal(1, added.Index);
        Assert.Equal(SectionStart.EvenPage, added.PageSetup.SectionStart);

        SectionWrapper middle = _document.Sections.Add(_document.Range(3, 3));
        Assert.Equal(3, _document.Sections.Count);
        Assert.Equal(1, middle.Index);
        Assert.Equal(SectionStart.NewPage, middle.PageSetup.SectionStart);
        Assert.Equal(" two three", middle.Range.Text.TrimEnd('\r').Substring(0, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Sections.Add(null, SectionStart.NewColumn));
        AssertValid();
    }

    [Fact]
    public void PageBordersAreTheFourEdgesWithADistanceFromTheText()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        BordersWrapper borders = setup.Borders;
        Assert.Equal(4, borders.Count);

        borders.OutsideLineStyle = BorderLineStyle.DoubleLine;
        borders.OutsideLineWidth = 1.5;
        borders.OutsideColor = "0000FF";

        Assert.Equal(BorderLineStyle.DoubleLine, _document.Sections[0].Borders[BorderPosition.Left].LineStyle);
        Assert.Equal(1.5, borders[BorderPosition.Bottom].LineWidth);
        Assert.Equal(24u, _document.Sections[0].SectionProperties.GetFirstChild<PageBorders>().TopBorder.Space.Value);
        Assert.Throws<NotSupportedException>(() => borders[BorderPosition.Horizontal]);
        AssertValid();

        borders.Clear();
        Assert.Null(_document.Sections[0].SectionProperties.GetFirstChild<PageBorders>());
    }

    #endregion

    #region Headers and footers

    [Fact]
    public void ExistsCanBeSetToCreateAndRemoveAHeader()
    {
        HeaderFooterWrapper header = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        Assert.False(header.Exists);

        header.Exists = true;
        Assert.True(header.Exists);
        Assert.Equal("\r", header.Range.Paragraphs[0].Text);
        header.Range.Text = "Top";

        header.Exists = false;
        Assert.False(header.Exists);
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.HeaderParts);
        AssertValid();
    }

    [Fact]
    public void PageNumbersAreAddedAsAPageField()
    {
        HeaderFooterWrapper footer = _document.Sections[0].Footers[HeaderFooterIndex.Primary];

        FieldWrapper field = footer.PageNumbers.Add(PageNumberAlignment.Center, true);

        Assert.Equal(FieldType.Page, field.Type);
        Assert.Equal(1, footer.PageNumbers.Count);
        Assert.Equal(ParagraphAlignment.Center, footer.Range.Paragraphs[0].Alignment);
        Assert.Throws<NotSupportedException>(() => footer.PageNumbers.Add(PageNumberAlignment.Inside));

        _document.Sections[0].Footers[HeaderFooterIndex.FirstPage].PageNumbers.Add(PageNumberAlignment.Right, false);
        Assert.True(_document.Sections[0].PageSetup.DifferentFirstPageHeaderFooter);
        AssertValid();
    }

    [Fact]
    public void PageNumberingOptionsBelongToTheSection()
    {
        PageNumbersWrapper numbers = _document.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers;
        Assert.False(numbers.RestartNumberingAtSection);
        Assert.Equal(1, numbers.StartingNumber);
        Assert.Equal(PageNumberStyle.Arabic, numbers.NumberStyle);

        numbers.RestartNumberingAtSection = true;
        numbers.StartingNumber = 5;
        numbers.NumberStyle = PageNumberStyle.LowercaseRoman;

        Assert.True(numbers.RestartNumberingAtSection);
        Assert.Equal(5, numbers.StartingNumber);
        Assert.Equal(PageNumberStyle.LowercaseRoman, numbers.NumberStyle);
        AssertValid();

        numbers.NumberStyle = PageNumberStyle.Arabic;
        numbers.RestartNumberingAtSection = false;
        Assert.Null(_document.Sections[0].SectionProperties.GetFirstChild<PageNumberType>());
    }

    #endregion

    #region Page setup

    [Fact]
    public void PaperSizesSetTheSizeAndAreReadBack()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;

        setup.PaperSize = PaperSize.A4;
        Assert.Equal(595.3, setup.PageWidth);
        Assert.Equal(841.9, setup.PageHeight);
        Assert.Equal(PaperSize.A4, setup.PaperSize);

        setup.Orientation = PageOrientation.Landscape;
        Assert.Equal(PaperSize.A4, setup.PaperSize);
        setup.PaperSize = PaperSize.Legal;
        Assert.Equal(1008, setup.PageWidth);
        Assert.Equal(612, setup.PageHeight);

        setup.PageWidth = 500;
        Assert.Equal(PaperSize.Custom, setup.PaperSize);
        Assert.Throws<ArgumentOutOfRangeException>(() => setup.PaperSize = PaperSize.Custom);
        AssertValid();
    }

    [Fact]
    public void VerticalAlignmentAndEndnotesAreSectionProperties()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        Assert.Equal(PageVerticalAlignment.Top, setup.VerticalAlignment);

        setup.VerticalAlignment = PageVerticalAlignment.Center;
        setup.SuppressEndnotes = true;

        Assert.Equal(PageVerticalAlignment.Center, setup.VerticalAlignment);
        Assert.True(setup.SuppressEndnotes);
        AssertValid();

        setup.VerticalAlignment = PageVerticalAlignment.Top;
        setup.SuppressEndnotes = false;
        Assert.Null(_document.Sections[0].SectionProperties.GetFirstChild<VerticalTextAlignmentOnPage>());
    }

    [Fact]
    public void MirrorMarginsAndGutterPositionAreSettingsOfTheDocument()
    {
        PageSetupWrapper setup = _document.Sections[0].PageSetup;
        Assert.False(setup.MirrorMargins);
        Assert.Equal(GutterPosition.Left, setup.GutterPos);

        setup.MirrorMargins = true;
        setup.GutterPos = GutterPosition.Top;

        Assert.True(setup.MirrorMargins);
        Assert.Equal(GutterPosition.Top, setup.GutterPos);
        setup.GutterPos = GutterPosition.Right;
        Assert.Equal(GutterPosition.Right, setup.GutterPos);
        AssertValid();

        setup.MirrorMargins = false;
        setup.GutterPos = GutterPosition.Left;
        Assert.False(setup.MirrorMargins);
        Assert.Equal(GutterPosition.Left, setup.GutterPos);
    }

    [Fact]
    public void LineNumberingHasItsOptions()
    {
        LineNumberingWrapper numbering = _document.Sections[0].PageSetup.LineNumbering;
        Assert.False(numbering.Active);

        numbering.Active = true;
        numbering.CountBy = 5;
        numbering.StartingNumber = 10;
        numbering.DistanceFromText = 18;
        numbering.RestartMode = LineNumberingRestart.Section;

        Assert.True(numbering.Active);
        Assert.Equal(5, numbering.CountBy);
        Assert.Equal(10, numbering.StartingNumber);
        Assert.Equal(18, numbering.DistanceFromText);
        Assert.Equal(LineNumberingRestart.Section, numbering.RestartMode);
        Assert.Throws<ArgumentOutOfRangeException>(() => numbering.CountBy = 0);
        AssertValid();

        numbering.Active = false;
        Assert.False(numbering.Active);
    }

    [Fact]
    public void TextColumnsAreEvenlySpacedOrListed()
    {
        TextColumnsWrapper columns = _document.Sections[0].PageSetup.TextColumns;
        Assert.Equal(1, columns.Count);

        columns.SetCount(3);
        columns.Spacing = 18;
        columns.LineBetween = true;

        Assert.Equal(3, columns.Count);
        Assert.True(columns.EvenlySpaced);
        Assert.Equal(18, columns.Spacing);
        Assert.True(columns.LineBetween);
        PageSetupWrapper page = _document.Sections[0].PageSetup;
        Assert.Equal((page.PageWidth - page.LeftMargin - page.RightMargin - 2 * 18) / 3.0, columns.Width, 1);
        AssertValid();

        columns.EvenlySpaced = false;
        Assert.False(columns.EvenlySpaced);
        Assert.Equal(3, columns.Count);
        AssertValid();

        columns.SetCount(1);
        Assert.Equal(1, columns.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => columns.SetCount(0));
    }

    [Fact]
    public void TheDocumentPageSetupAppliesToEverySection()
    {
        _document.Sections.Add(null, SectionStart.NewPage);

        _document.PageSetup.Orientation = PageOrientation.Landscape;
        _document.PageSetup.LeftMargin = 36;
        _document.PageSetup.PaperSize = PaperSize.A4;

        Assert.All(_document.Sections, s =>
        {
            Assert.Equal(PageOrientation.Landscape, s.PageSetup.Orientation);
            Assert.Equal(36, s.PageSetup.LeftMargin);
            Assert.Equal(PaperSize.A4, s.PageSetup.PaperSize);
        });
        Assert.Equal(PageOrientation.Landscape, _document.PageSetup.Orientation);
        AssertValid();
    }

    #endregion

    #region Document

    [Fact]
    public void TheDocumentKnowsItsFile()
    {
        Assert.Equal(Path.GetFileName(_path), _document.Name);
        Assert.Equal(Path.GetFullPath(_path), _document.FullName);
        Assert.Equal(Path.GetDirectoryName(Path.GetFullPath(_path)), _document.Path);
        Assert.False(_document.ReadOnly);

        using (MemoryStream stream = new MemoryStream())
        using (WordprocessingDocumentWrapper inMemory = WordprocessingDocumentWrapper.Create(stream, true, false))
        {
            Assert.Equal("Document1", inMemory.Name);
            Assert.Equal("", inMemory.Path);
        }
    }

    [Fact]
    public void AReadOnlyDocumentSaysSo()
    {
        string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            WordprocessingDocumentWrapper.Create(path, true, false).Close();
            using (WordprocessingDocumentWrapper readOnly = WordprocessingDocumentWrapper.Open(path))
            {
                Assert.True(readOnly.ReadOnly);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ClosingCanSaveOrDiscardTheChanges()
    {
        string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            WordprocessingDocumentWrapper first = WordprocessingDocumentWrapper.Create(path, true, false);
            first.Content.Text = "saved";
            first.Close(SaveOptions.SaveChanges);

            WordprocessingDocumentWrapper second = WordprocessingDocumentWrapper.Open(path, true);
            Assert.StartsWith("saved", second.Content.Text, StringComparison.Ordinal);
            second.Content.Text = "discarded";
            second.Close(SaveOptions.DoNotSaveChanges);

            using (WordprocessingDocumentWrapper third = WordprocessingDocumentWrapper.Open(path))
            {
                Assert.StartsWith("saved", third.Content.Text, StringComparison.Ordinal);
            }

            WordprocessingDocumentWrapper fourth = WordprocessingDocumentWrapper.Open(path, true);
            fourth.Content.Text = "default";
            fourth.Close();
            using (WordprocessingDocumentWrapper fifth = WordprocessingDocumentWrapper.Open(path))
            {
                Assert.StartsWith("default", fifth.Content.Text, StringComparison.Ordinal);
            }
            Assert.Throws<NotSupportedException>(() => _document.Close(SaveOptions.PromptToSaveChanges));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(SaveFormat.XmlDocument, WordprocessingDocumentType.Document)]
    [InlineData(SaveFormat.XmlDocumentMacroEnabled, WordprocessingDocumentType.MacroEnabledDocument)]
    [InlineData(SaveFormat.XmlTemplate, WordprocessingDocumentType.Template)]
    [InlineData(SaveFormat.XmlTemplateMacroEnabled, WordprocessingDocumentType.MacroEnabledTemplate)]
    public void ADocumentCanBeSavedInAnotherOpenXmlFormat(SaveFormat format, WordprocessingDocumentType expected)
    {
        string copyPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (WordprocessingDocumentWrapper copy = _document.SaveAs(copyPath, format))
            {
                Assert.Equal(expected, copy.WordprocessingDocument.DocumentType);
                Assert.StartsWith("one two three", copy.Content.Text, StringComparison.Ordinal);
            }
        }
        finally
        {
            File.Delete(copyPath);
        }
    }

    [Fact]
    public void TheDocumentHasWordsCharactersAndSentences()
    {
        //As in Word the paragraph mark counts as a word.
        Assert.Equal(4, _document.Words.Count);
        Assert.Equal(1, _document.Sentences.Count);
        Assert.Equal(_document.Content.Characters.Count, _document.Characters.Count);
    }

    [Fact]
    public void StatisticsAreComputedFromTheText()
    {
        _document.Content.Text = "one two  three\rfour";

        Assert.Equal(4, _document.ComputeStatistics(Statistic.Words));
        Assert.Equal(15, _document.ComputeStatistics(Statistic.Characters));
        Assert.Equal(18, _document.ComputeStatistics(Statistic.CharactersWithSpaces));
        Assert.Equal(2, _document.ComputeStatistics(Statistic.Paragraphs));
        Assert.Throws<NotSupportedException>(() => _document.ComputeStatistics(Statistic.Pages));
        Assert.Throws<NotSupportedException>(() => _document.ComputeStatistics(Statistic.Lines));
    }

    [Fact]
    public void TrackRevisionsAndCompatibilityModeAreSettings()
    {
        Assert.False(_document.TrackRevisions);
        _document.TrackRevisions = true;
        Assert.True(_document.TrackRevisions);
        _document.TrackRevisions = false;
        Assert.False(_document.TrackRevisions);

        Assert.Equal(15, _document.CompatibilityMode);
        _document.CompatibilityMode = 14;
        Assert.Equal(14, _document.CompatibilityMode);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.CompatibilityMode = 13);
        AssertValid();
    }

    [Fact]
    public void ADocumentCanBeProtectedWithAPassword()
    {
        Assert.Equal(ProtectionType.NoProtection, _document.ProtectionType);

        _document.Protect(ProtectionType.AllowOnlyComments, "secret");
        Assert.Equal(ProtectionType.AllowOnlyComments, _document.ProtectionType);
        Assert.Throws<InvalidOperationException>(() => _document.Protect(ProtectionType.AllowOnlyReading));
        AssertValid();

        Assert.Throws<UnauthorizedAccessException>(() => _document.Unprotect("wrong"));
        Assert.Equal(ProtectionType.AllowOnlyComments, _document.ProtectionType);
        _document.Unprotect("secret");
        Assert.Equal(ProtectionType.NoProtection, _document.ProtectionType);

        _document.Protect(ProtectionType.AllowOnlyReading);
        Assert.Equal(ProtectionType.AllowOnlyReading, _document.ProtectionType);
        _document.Unprotect();
        Assert.Equal(ProtectionType.NoProtection, _document.ProtectionType);
    }

    [Fact]
    public void VariablesAreStoredInTheDocumentAndShownByAField()
    {
        Assert.Equal(0, _document.Variables.Count);

        VariableWrapper variable = _document.Variables.Add("Client", "Acme");
        _document.Variables.Add("Other", "x");
        _document.Variables.Add("client", "Acme Ltd");

        Assert.Equal(2, _document.Variables.Count);
        Assert.Equal("Acme Ltd", _document.Variables["CLIENT"].Value);
        Assert.Equal("Acme Ltd", variable.Value);
        Assert.Null(_document.Variables["missing"]);
        Assert.Equal("Other", _document.Variables[1].Name);

        FieldWrapper field = _document.Fields.Add(_document.Paragraphs.Last.Range, FieldType.DocVariable, "Client", false);
        Assert.Equal("Acme Ltd", field.Result.Text);
        AssertValid();

        variable.Delete();
        Assert.Equal(1, _document.Variables.Count);
        Assert.False(_document.Variables.Exists("Client"));
    }

    #endregion

    #region Properties

    [Fact]
    public void MoreBuiltInPropertiesAreAvailable()
    {
        DocumentPropertiesWrapper properties = _document.BuiltInDocumentProperties;

        properties.Revision = 3;
        properties.Language = "en-GB";
        properties.ContentStatus = "Draft";
        properties.LastPrintDate = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc);
        properties.Template = "Normal.dotm";
        properties.TotalEditingTime = 42;
        properties.HyperlinkBase = "http://example.com/";
        properties.ApplicationName = "Test";
        properties.Pages = 2;
        properties.Words = 3;
        properties.Characters = 13;
        properties.CharactersWithSpaces = 15;
        properties.Lines = 1;
        properties.Paragraphs = 1;
        AssertValid();

        Assert.Equal(3, properties.Revision);
        Assert.Equal("en-GB", properties.Language);
        Assert.Equal("Draft", properties.ContentStatus);
        Assert.Equal(new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Utc), properties.LastPrintDate);
        Assert.Equal("Normal.dotm", properties.Template);
        Assert.Equal(42, properties.TotalEditingTime);
        Assert.Equal("http://example.com/", properties.HyperlinkBase);
        Assert.Equal("Test", properties.ApplicationName);
        Assert.Equal(2, properties.Pages);
        Assert.Equal(15, properties.CharactersWithSpaces);

        properties.Template = null;
        Assert.Null(properties.Template);
    }

    [Fact]
    public void ABuiltInPropertyCanBeFoundByItsWordName()
    {
        DocumentPropertiesWrapper properties = _document.BuiltInDocumentProperties;

        properties["Title"] = "Named";
        properties["Number of words"] = 99;
        properties["Last Author"] = "Pat";

        Assert.Equal("Named", properties.Title);
        Assert.Equal(99, properties.Words);
        Assert.Equal("Pat", properties["last author"]);
        Assert.Equal(99, properties["NUMBER OF WORDS"]);
        Assert.Throws<ArgumentException>(() => properties["Nonsense"]);
    }

    #endregion
}
