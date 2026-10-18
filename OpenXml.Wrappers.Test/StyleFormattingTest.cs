using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class StyleFormattingTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public StyleFormattingTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        StyleDefinitionsPart part = _document.WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        part.Styles = new Styles(new Style(new StyleName() { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true });
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void AStyleStartsWithNoDirectFormatting()
    {
        StyleWrapper style = _document.Styles.Add("Note", StyleType.Paragraph);

        Assert.False(style.Font.Bold);
        Assert.Null(style.Font.Direct.Size);
        Assert.Null(style.Font.Direct.Name);
        Assert.Equal(ParagraphAlignment.Left, style.ParagraphFormat.Alignment);
        Assert.Equal(0, style.ParagraphFormat.SpaceAfter);
        Assert.Null(style.Style.StyleRunProperties);
        Assert.Null(style.Style.StyleParagraphProperties);
    }

    [Fact]
    public void FontFormattingIsStoredInTheStyleRunProperties()
    {
        StyleWrapper style = _document.Styles.Add("Note", StyleType.Paragraph);

        style.Font.Bold = true;
        style.Font.Size = 10.5;
        style.Font.Name = "Arial";
        style.Font.Color = "C00000";

        Assert.True(style.Font.Bold);
        Assert.Equal(10.5, style.Font.Size);
        Assert.Equal("Arial", style.Font.Name);
        Assert.Equal("C00000", style.Font.Color);
        Assert.Equal("21", style.Style.StyleRunProperties.GetFirstChild<FontSize>().Val.Value);

        style.Font.Bold = null;
        style.Font.Size = null;
        style.Font.Name = null;
        style.Font.Color = null;
        Assert.Null(style.Style.StyleRunProperties);
    }

    [Fact]
    public void ParagraphFormattingIsStoredInTheStyleParagraphProperties()
    {
        StyleWrapper style = _document.Styles.Add("Note", StyleType.Paragraph);

        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        style.ParagraphFormat.SpaceAfter = 6;
        style.ParagraphFormat.LeftIndent = 36;
        style.ParagraphFormat.KeepWithNext = true;

        Assert.Equal(ParagraphAlignment.Center, style.ParagraphFormat.Alignment);
        Assert.Equal(6, style.ParagraphFormat.SpaceAfter);
        Assert.Equal(36, style.ParagraphFormat.LeftIndent);
        Assert.True(style.ParagraphFormat.KeepWithNext);
        StyleParagraphProperties properties = style.Style.StyleParagraphProperties;
        Assert.Equal("120", properties.SpacingBetweenLines.After.Value);
        Assert.Equal("720", properties.Indentation.Left.Value);

        style.ParagraphFormat.Alignment = null;
        style.ParagraphFormat.SpaceAfter = null;
        style.ParagraphFormat.LeftIndent = null;
        style.ParagraphFormat.KeepWithNext = null;
        Assert.Null(style.Style.StyleParagraphProperties);
    }

    [Fact]
    public void TheChildrenOfAStyleStayInSchemaOrder()
    {
        StyleWrapper style = _document.Styles.Add("Note", StyleType.Paragraph);

        //Set the run properties first, so the paragraph properties have to be placed before them.
        style.Font.Italic = true;
        style.ParagraphFormat.Alignment = ParagraphAlignment.Right;

        string[] names = style.Style.ChildElements.Select(e => e.LocalName).ToArray();
        Assert.Equal(new[] { "name", "basedOn", "pPr", "rPr" }, names);
    }

    [Fact]
    public void FormattingOfACharacterStyleAndOtherStylesIsIndependent()
    {
        StyleWrapper emphasis = _document.Styles.Add("Emphasis2", StyleType.Character);
        StyleWrapper note = _document.Styles.Add("Note", StyleType.Paragraph);

        emphasis.Font.Italic = true;

        Assert.True(emphasis.Font.Italic);
        Assert.False(note.Font.Italic);
    }

    [Fact]
    public void FormattingAppliesToTheStyleOnlyNotItsDerivedStyles()
    {
        StyleWrapper normal = _document.Styles["Normal"];
        StyleWrapper note = _document.Styles.Add("Note", StyleType.Paragraph);

        normal.Font.Size = 11;

        Assert.Equal(11, normal.Font.Size);
        Assert.Null(note.Font.Direct.Size);
        Assert.Equal(11, note.Font.Size);
    }

    [Fact]
    public void FormattingSurvivesSavingAndReopening()
    {
        StyleWrapper style = _document.Styles.Add("Note", StyleType.Paragraph);
        style.Font.Bold = true;
        style.ParagraphFormat.SpaceBefore = 12;
        string copy = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            _document.SaveAs(copy).Close();
            using (WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(copy, false))
            {
                StyleWrapper reopenedStyle = reopened.Styles["Note"];
                Assert.True(reopenedStyle.Font.Bold);
                Assert.Equal(12, reopenedStyle.ParagraphFormat.SpaceBefore);
            }
        }
        finally
        {
            File.Delete(copy);
        }
    }
}
