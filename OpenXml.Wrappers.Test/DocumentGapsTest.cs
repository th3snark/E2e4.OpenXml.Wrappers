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
/// Odd and even page headers, deleting header ranges, the text of the last paragraph of a cell, and borders that come from a table style.
/// </summary>
public class DocumentGapsTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");

    public void Dispose()
    {
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void OddAndEvenPagesCanBeSwitchedOnAndOffAndTheFileStaysValid()
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            PageSetupWrapper setup = document.Sections[0].PageSetup;
            Assert.False(setup.OddAndEvenPagesHeaderFooter);

            setup.OddAndEvenPagesHeaderFooter = true;
            Assert.True(setup.OddAndEvenPagesHeaderFooter);
            document.Sections[0].Headers[HeaderFooterIndex.EvenPages].Range.Text = "even";
        }

        using (WordprocessingDocument file = WordprocessingDocument.Open(_path, false))
        {
            Assert.NotNull(file.MainDocumentPart.DocumentSettingsPart.Settings.GetFirstChild<EvenAndOddHeaders>());
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(file));
        }

        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(_path, true))
        {
            document.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter = false;
            Assert.False(document.Sections[0].PageSetup.OddAndEvenPagesHeaderFooter);
        }
    }

    [Fact]
    public void DeletingAHeaderRangeLeavesOneEmptyParagraph()
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            RangeWrapper header = document.Sections[0].Headers[HeaderFooterIndex.Primary].Range;
            header.Text = "one\rtwo";

            header.Delete();

            Assert.Single(header.Paragraphs);
            Assert.Equal(string.Empty, header.Paragraphs[0].Text);
        }
    }

    [Fact]
    public void ParagraphTextOfTheLastParagraphOfACellHasNoCellMarker()
    {
        TableCell cell = new TableCell(new Paragraph(new Run(new Text("only"))));
        Table table = new Table(new TableRow(cell));

        Assert.Equal("only", new ParagraphWrapper(cell.Elements<Paragraph>().Single()).Text);
    }

    [Fact]
    public void BordersFromTheTableStyleAreRead()
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            StylesPart(document).Styles = new Styles(
                new Style(
                    new StyleName() { Val = "Base Grid" },
                    new StyleTableProperties(new TableBorders(
                        new TopBorder() { Val = BorderValues.Single, Size = 12, Color = "FF0000" },
                        new InsideHorizontalBorder() { Val = BorderValues.Dashed, Size = 4 })))
                { Type = StyleValues.Table, StyleId = "BaseGrid" },
                new Style(
                    new StyleName() { Val = "My Grid" },
                    new BasedOn() { Val = "BaseGrid" })
                { Type = StyleValues.Table, StyleId = "MyGrid" });

            TableWrapper table = document.Tables.Add(document.Paragraphs[0].Range, 2, 2);
            table.Borders.Clear();
            table.Table.GetFirstChild<TableProperties>().TableStyle = new TableStyle() { Val = "MyGrid" };

            //The table reads the style it is based on; a top cell on the edge takes the outside border, a bottom one the inside.
            Assert.Equal(BorderLineStyle.SingleLine, table.Borders[BorderPosition.Top].LineStyle);
            Assert.Equal(1.5, table.Borders[BorderPosition.Top].LineWidth);
            Assert.Equal("FF0000", table.Cell(0, 1).Borders[BorderPosition.Top].Color);
            Assert.Equal(BorderLineStyle.DashLargeGap, table.Cell(0, 0).Borders[BorderPosition.Bottom].LineStyle);
            Assert.Equal(BorderLineStyle.None, table.Cell(1, 0).Borders[BorderPosition.Bottom].LineStyle);

            //A direct setting wins over the style.
            table.Cell(0, 0).Borders[BorderPosition.Top].LineStyle = BorderLineStyle.DoubleLine;
            Assert.Equal(BorderLineStyle.DoubleLine, table.Cell(0, 0).Borders[BorderPosition.Top].LineStyle);
        }
    }

    [Fact]
    public void BordersFromTheConditionalFormattingOfTheTableStyleAreReadWhenTheLookSwitchesThemOn()
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(_path, true, false))
        {
            StylesPart(document).Styles = new Styles(
                new Style(
                    new StyleName() { Val = "Header Grid" },
                    new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(new TableCellBorders(new BottomBorder() { Val = BorderValues.Double, Size = 6 })))
                    { Type = TableStyleOverrideValues.FirstRow })
                { Type = StyleValues.Table, StyleId = "HeaderGrid" });

            TableWrapper table = document.Tables.Add(document.Paragraphs[0].Range, 2, 2);
            table.Borders.Clear();
            TableProperties properties = table.Table.GetFirstChild<TableProperties>();
            properties.TableStyle = new TableStyle() { Val = "HeaderGrid" };

            Assert.Equal(BorderLineStyle.None, table.Cell(0, 0).Borders[BorderPosition.Bottom].LineStyle);

            properties.TableLook = new TableLook() { FirstRow = true };
            Assert.Equal(BorderLineStyle.DoubleLine, table.Cell(0, 0).Borders[BorderPosition.Bottom].LineStyle);
            Assert.Equal(BorderLineStyle.None, table.Cell(1, 0).Borders[BorderPosition.Bottom].LineStyle);
        }
    }

    [Fact]
    public void KeepWithNextOfARowAndACellCanBeRead()
    {
        TableCell first = new TableCell(new Paragraph());
        TableCell second = new TableCell(new Paragraph());
        TableRow row = new TableRow(first, second);
        Table table = new Table(row);
        TableRowWrapper wrapper = new TableRowWrapper(row);
        Assert.False(wrapper.KeepWithNext);

        new TableCellWrapper(first).KeepWithNext = true;
        Assert.True(new TableCellWrapper(first).KeepWithNext);
        Assert.False(wrapper.KeepWithNext);

        wrapper.KeepWithNext = true;
        Assert.True(wrapper.KeepWithNext);
    }

    private static StyleDefinitionsPart StylesPart(WordprocessingDocumentWrapper document)
    {
        return document.WordprocessingDocument.MainDocumentPart.StyleDefinitionsPart
            ?? document.WordprocessingDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
    }
}
