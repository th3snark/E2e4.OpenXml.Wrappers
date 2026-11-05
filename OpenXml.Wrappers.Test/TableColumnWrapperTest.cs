using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class TableColumnWrapperTest
{
    private static Table CreateTable()
    {
        return new Table(
            new TableGrid(new GridColumn() { Width = "1000" }, new GridColumn() { Width = "2000" }),
            new TableRow(new TableCell(new Paragraph()), new TableCell(new Paragraph())),
            new TableRow(new TableCell(new Paragraph())));
    }

    [Fact]
    public void ColumnWrappersOfTheSameGridColumnAreEqual()
    {
        Table table = CreateTable();
        TableWrapper wrapper = new TableWrapper(table);

        Assert.True(wrapper.Columns[0] == wrapper.Columns[0]);
        Assert.True(wrapper.Columns[0] != wrapper.Columns[1]);
    }

    [Fact]
    public void WidthIsReadFromTheGridColumn()
    {
        Table table = CreateTable();

        Assert.Equal(100, new TableColumnWrapper(table.Descendants<GridColumn>().Last()).Width);
    }

    [Fact]
    public void WidthsOverWhatWordAllowsAreRejected()
    {
        Table table = CreateTable();
        TableColumnWrapper column = new TableColumnWrapper(table.Descendants<GridColumn>().First());

        column.Width = 1584;

        Assert.Equal(1584, column.Width);
        Assert.Throws<ArgumentOutOfRangeException>(() => column.Width = 1585);
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableCellWrapper(table.Descendants<TableCell>().First()).Width = 1585);
        Assert.Throws<ArgumentOutOfRangeException>(() => column.PreferredWidth = 1585);
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableWrapper(table).PreferredWidth = 1585);
    }

    [Fact]
    public void ACellThatSpansColumnsIsAsWideAsThey()
    {
        Table table = new Table(
            new TableGrid(new GridColumn() { Width = "1000" }, new GridColumn() { Width = "2000" }),
            new TableRow(new TableCell(new Paragraph()), new TableCell(new Paragraph())),
            new TableRow(new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }), new Paragraph())));

        new TableColumnWrapper(table.Descendants<GridColumn>().First()).Width = 100;

        Assert.Equal(new[] { "2000", "2000", "4000" }, table.Descendants<TableCellWidth>().Select(w => w.Width.Value));
    }

    //Word's grid for a percentage differs from these by a twip here and there, so the scenarios cannot compare them. The values are Word's.
    [Theory]
    [InlineData(50, "200", 3885)]
    [InlineData(80, "40", 3111)]
    [InlineData(50, "add", 3334)]
    public void APercentageFollowsTheColumnsAsAShareOfTheSpace(int percent, string change, int expected)
    {
        using (MemoryStream stream = new MemoryStream())
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
        {
            TableWrapper table = document.Tables.Add(document.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
            table.PreferredWidthType = PreferredWidthType.Percent;
            table.PreferredWidth = percent;

            if (change == "200")
            {
                table.Columns[0].Width = 200;
            }
            else if (change == "40")
            {
                table.Columns[1].Width = 40;
            }
            else
            {
                table.Columns.Add();
            }

            TableWidth width = table.Table.GetFirstChild<TableProperties>().TableWidth;
            Assert.Equal(TableWidthUnitValues.Pct, width.Type.Value);
            Assert.Equal(expected.ToString(System.Globalization.CultureInfo.InvariantCulture), width.Width.Value);
        }
    }

    [Fact]
    public void ATableIsMeasuredAgainstTheSectionItIsIn()
    {
        using (MemoryStream stream = new MemoryStream())
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
        {
            document.Paragraphs[0].InsertSectionBreak(SectionStart.NewPage);
            document.Sections[0].PageSetup.Orientation = PageOrientation.Landscape;

            TableWrapper table = document.Tables.Add(document.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);

            //A4 landscape with margins of one inch, less the 10 twips Word takes off.
            Assert.Equal(13948 / 20.0, table.Width);
        }
    }

    [Fact]
    public void SettingTheWidthSkipsRowsWithTooFewCellsAndCellsWithoutProperties()
    {
        Table table = CreateTable();
        TableColumnWrapper column = new TableColumnWrapper(table.Descendants<GridColumn>().Last());

        column.Width = 150;

        Assert.Equal(150, column.Width);
        TableCell cell = table.Elements<TableRow>().First().Elements<TableCell>().Last();
        Assert.Equal("3000", cell.TableCellProperties.TableCellWidth.Width.Value);
        Assert.Equal(TableWidthUnitValues.Dxa, cell.TableCellProperties.TableCellWidth.Type.Value);
    }

    [Fact]
    public void SettingTheWidthUpdatesAnExistingCellWidth()
    {
        Table table = CreateTable();
        TableCell cell = table.Elements<TableRow>().First().Elements<TableCell>().First();
        new TableCellWrapper(cell).Width = 0.5;

        new TableColumnWrapper(table.Descendants<GridColumn>().First()).Width = 200;

        Assert.Equal(200, new TableCellWrapper(cell).Width);
    }
}
