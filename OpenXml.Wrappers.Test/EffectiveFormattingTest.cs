using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class EffectiveFormattingTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public EffectiveFormattingTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        StyleDefinitionsPart part = _document.WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        part.Styles = new Styles(
            new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri" }, new FontSize() { Val = "22" })),
                new ParagraphPropertiesDefault(new ParagraphPropertiesBaseStyle(new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }))),
            new Style(new StyleName() { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
            new Style(
                new StyleName() { Val = "heading 1" },
                new BasedOn() { Val = "Normal" },
                new StyleParagraphProperties(new KeepNext(), new SpacingBetweenLines() { Before = "240" }, new Indentation() { Left = "720" }),
                new StyleRunProperties(new Bold(), new FontSize() { Val = "32" }, new Color() { Val = "2F5496" }))
            { Type = StyleValues.Paragraph, StyleId = "Heading1" },
            new Style(new StyleName() { Val = "Heading 2 Like" }, new BasedOn() { Val = "Heading1" }, new StyleRunProperties(new FontSize() { Val = "26" }))
            { Type = StyleValues.Paragraph, StyleId = "Heading2Like" },
            new Style(new StyleName() { Val = "Strong" }, new StyleRunProperties(new Bold())) { Type = StyleValues.Character, StyleId = "Strong" },
            new Style(new StyleName() { Val = "Table Grid" }, new StyleRunProperties(new Italic())) { Type = StyleValues.Table, StyleId = "TableGrid" });
        _document.Content.Text = "x";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private ParagraphWrapper Paragraph(string styleId, string runStyleId = null)
    {
        Paragraph paragraph = _document.Paragraphs[0].Paragraph;
        paragraph.ParagraphProperties = styleId == null ? null : new ParagraphProperties(new ParagraphStyleId() { Val = styleId });
        Run run = paragraph.Elements<Run>().First();
        run.RunProperties = runStyleId == null ? null : new RunProperties(new RunStyle() { Val = runStyleId });
        return _document.Paragraphs[0];
    }

    [Fact]
    public void DocumentDefaultsApplyToUnformattedText()
    {
        FontWrapper font = Paragraph(null).Range.Font;

        Assert.Equal("Calibri", font.Name);
        Assert.Equal(11, font.Size);
        Assert.False(font.Bold);
        Assert.Equal("auto", font.Color);
        Assert.Null(font.Direct.Name);
        Assert.Null(font.Direct.Size);
    }

    [Fact]
    public void ParagraphStyleFormattingIsInheritedAlongTheBasedOnChain()
    {
        FontWrapper font = Paragraph("Heading2Like").Range.Font;

        Assert.True(font.Bold);
        Assert.Equal(13, font.Size);
        Assert.Equal("2F5496", font.Color);
        Assert.Equal("Calibri", font.Name);
        Assert.False(font.Direct.Bold);
    }

    [Fact]
    public void DirectFormattingOverridesTheStyleAndSettingNullRestoresIt()
    {
        ParagraphWrapper paragraph = Paragraph("Heading1");

        paragraph.Range.Font.Size = 20;
        paragraph.Range.Font.Bold = false;
        Assert.Equal(20, paragraph.Range.Font.Size);
        Assert.False(paragraph.Range.Font.Bold);

        paragraph.Range.Font.Size = null;
        paragraph.Range.Font.Bold = null;
        Assert.Equal(16, paragraph.Range.Font.Size);
        Assert.True(paragraph.Range.Font.Bold);
    }

    [Fact]
    public void OnOffPropertiesOfDifferentStylesToggle()
    {
        Assert.True(Paragraph(null, "Strong").Range.Font.Bold);
        //Strong is bold and so is Heading 1, so the text is not bold.
        Assert.False(Paragraph("Heading1", "Strong").Range.Font.Bold);
        //A direct setting wins over the toggling.
        Paragraph("Heading1", "Strong").Range.Font.Bold = true;
        Assert.True(_document.Paragraphs[0].Range.Font.Bold);
    }

    [Fact]
    public void TableStyleFormattingAppliesToTextInTheTable()
    {
        Table table = new Table(
            new TableProperties(new TableStyle() { Val = "TableGrid" }),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("cell"))))));
        _document.WordprocessingDocument.MainDocumentPart.Document.Body.PrependChild(table);

        Assert.True(new RangeWrapper(table.Descendants<Paragraph>().First()).Font.Italic);
        Assert.False(_document.Paragraphs.Last().Range.Font.Italic);
    }

    [Fact]
    public void ParagraphFormatIsInheritedAttributeByAttribute()
    {
        ParagraphFormatWrapper format = Paragraph("Heading1").Range.ParagraphFormat;

        Assert.Equal(12, format.SpaceBefore);
        Assert.Equal(8, format.SpaceAfter);
        Assert.Equal(36, format.LeftIndent);
        Assert.True(format.KeepWithNext);
        Assert.Equal(0, format.Direct.SpaceBefore);
        Assert.Equal(0, format.Direct.LeftIndent);
        Assert.False(format.Direct.KeepWithNext);
    }

    [Fact]
    public void AStyleReportsWhatItInheritsAndDirectReadsOnlyItsOwn()
    {
        StyleWrapper style = _document.Styles["Heading2Like"];

        Assert.True(style.Font.Bold);
        Assert.Equal(13, style.Font.Size);
        Assert.Equal(12, style.ParagraphFormat.SpaceBefore);
        Assert.False(style.Font.Direct.Bold);
        Assert.Equal(13, style.Font.Direct.Size);
        Assert.Equal(0, style.ParagraphFormat.Direct.SpaceBefore);
    }

    [Fact]
    public void RunsThatDifferInTheirEffectiveValueReadAsNull()
    {
        Paragraph paragraph = _document.Paragraphs[0].Paragraph;
        paragraph.AppendChild(new Run(new RunProperties(new RunStyle() { Val = "Strong" }), new Text("y")));

        Assert.Null(_document.Paragraphs[0].Range.Font.Bold);
    }
}
