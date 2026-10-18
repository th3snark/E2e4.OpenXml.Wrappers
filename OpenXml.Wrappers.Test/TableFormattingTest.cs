using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class TableFormattingTest
{
    private static Table CreateTable(int rows, int columns)
    {
        Table table = new Table(new TableProperties(), new TableGrid(Enumerable.Range(0, columns).Select(i => new GridColumn() { Width = "1000" })));
        for (int r = 0; r < rows; r++)
        {
            table.AppendChild(new TableRow(Enumerable.Range(0, columns).Select(c => new TableCell(new Paragraph(new Run(new Text(r + "," + c)))))));
        }
        return table;
    }

    [Fact]
    public void RowHeightIsInPointsAndDefaultsToAtLeast()
    {
        TableRow row = CreateTable(1, 1).Elements<TableRow>().Single();
        TableRowWrapper wrapper = new TableRowWrapper(row);
        Assert.Equal(0, wrapper.Height);
        Assert.Equal(RowHeightRule.Auto, wrapper.HeightRule);

        wrapper.Height = 18.5;

        Assert.Equal(18.5, wrapper.Height);
        Assert.Equal(RowHeightRule.AtLeast, wrapper.HeightRule);
        Assert.Equal(370u, row.TableRowProperties.GetFirstChild<TableRowHeight>().Val.Value);
        Assert.IsType<TableRowProperties>(row.FirstChild);

        wrapper.HeightRule = RowHeightRule.Exactly;
        Assert.Equal(RowHeightRule.Exactly, wrapper.HeightRule);

        wrapper.HeightRule = RowHeightRule.Auto;
        Assert.Equal(0, wrapper.Height);
    }

    [Fact]
    public void RowHeightRejectsNegativeValues()
    {
        TableRowWrapper wrapper = new TableRowWrapper(CreateTable(1, 1).Elements<TableRow>().Single());
        Assert.Throws<ArgumentOutOfRangeException>(() => wrapper.Height = -1);
    }

    [Fact]
    public void RowPropertiesGoAfterTablePropertyExceptions()
    {
        TableRow row = new TableRow(new TablePropertyExceptions(), new TableCell(new Paragraph()));
        new TableRowWrapper(row).HeadingFormat = true;

        Assert.Equal(new[] { "tblPrEx", "trPr", "tc" }, row.ChildElements.Select(e => e.LocalName).ToArray());
    }

    [Fact]
    public void HeadingFormatAndAllowBreakAcrossPagesToggle()
    {
        TableRow row = CreateTable(1, 1).Elements<TableRow>().Single();
        TableRowWrapper wrapper = new TableRowWrapper(row);
        Assert.False(wrapper.HeadingFormat);
        Assert.True(wrapper.AllowBreakAcrossPages);

        wrapper.HeadingFormat = true;
        wrapper.AllowBreakAcrossPages = false;
        Assert.True(wrapper.HeadingFormat);
        Assert.False(wrapper.AllowBreakAcrossPages);
        Assert.Single(row.TableRowProperties.Elements<TableHeader>());
        Assert.Single(row.TableRowProperties.Elements<CantSplit>());

        wrapper.HeadingFormat = true;
        Assert.Single(row.TableRowProperties.Elements<TableHeader>());

        wrapper.HeadingFormat = false;
        wrapper.AllowBreakAcrossPages = true;
        Assert.False(wrapper.HeadingFormat);
        Assert.True(wrapper.AllowBreakAcrossPages);
    }

    [Fact]
    public void CellVerticalAlignmentRoundTrips()
    {
        TableCell cell = CreateTable(1, 1).Descendants<TableCell>().Single();
        TableCellWrapper wrapper = new TableCellWrapper(cell);
        Assert.Equal(CellVerticalAlignment.Top, wrapper.VerticalAlignment);

        wrapper.VerticalAlignment = CellVerticalAlignment.Center;
        Assert.Equal(CellVerticalAlignment.Center, wrapper.VerticalAlignment);
        wrapper.VerticalAlignment = CellVerticalAlignment.Bottom;
        Assert.Equal(CellVerticalAlignment.Bottom, wrapper.VerticalAlignment);

        wrapper.VerticalAlignment = CellVerticalAlignment.Top;
        Assert.Null(cell.TableCellProperties.TableCellVerticalAlignment);
    }

    [Fact]
    public void CellPropertiesStayInSchemaOrder()
    {
        TableCell cell = CreateTable(1, 1).Descendants<TableCell>().Single();
        TableCellWrapper wrapper = new TableCellWrapper(cell);
        wrapper.VerticalAlignment = CellVerticalAlignment.Center;
        wrapper.BackgroundColor = "FFCC00";
        wrapper.Width = 100;

        Assert.Equal(new[] { "tcW", "shd", "vAlign" }, cell.TableCellProperties.ChildElements.Select(e => e.LocalName).ToArray());
    }

    [Fact]
    public void CellBackgroundColorRoundTripsAndValidates()
    {
        TableCell cell = CreateTable(1, 1).Descendants<TableCell>().Single();
        TableCellWrapper wrapper = new TableCellWrapper(cell);
        Assert.Null(wrapper.BackgroundColor);

        wrapper.BackgroundColor = "ffcc00";
        Assert.Equal("ffcc00", wrapper.BackgroundColor);
        Assert.Throws<ArgumentException>(() => wrapper.BackgroundColor = "red");

        wrapper.BackgroundColor = null;
        Assert.Null(wrapper.BackgroundColor);
        Assert.Null(cell.TableCellProperties.Shading);
    }

    [Fact]
    public void TableAlignmentRoundTrips()
    {
        Table table = CreateTable(1, 1);
        TableWrapper wrapper = new TableWrapper(table);
        Assert.Equal(RowAlignment.Left, wrapper.Alignment);

        wrapper.Alignment = RowAlignment.Right;
        Assert.Equal(RowAlignment.Right, wrapper.Alignment);
        wrapper.Alignment = RowAlignment.Center;
        Assert.Equal(RowAlignment.Center, wrapper.Alignment);

        wrapper.Alignment = RowAlignment.Left;
        Assert.Null(table.GetFirstChild<TableProperties>().TableJustification);
    }

    [Fact]
    public void TablePreferredWidthIsInPoints()
    {
        Table table = CreateTable(1, 1);
        TableWrapper wrapper = new TableWrapper(table);
        Assert.Equal(0, wrapper.PreferredWidth);
        Assert.Equal(PreferredWidthType.Auto, wrapper.PreferredWidthType);

        wrapper.PreferredWidth = 468;
        Assert.Equal(PreferredWidthType.Points, wrapper.PreferredWidthType);
        Assert.Equal(468, wrapper.PreferredWidth);
        Assert.Equal("9360", table.GetFirstChild<TableProperties>().TableWidth.Width.Value);

        wrapper.PreferredWidthType = PreferredWidthType.Percent;
        Assert.Equal(PreferredWidthType.Percent, wrapper.PreferredWidthType);
        wrapper.PreferredWidth = 50;
        Assert.Equal(50, wrapper.PreferredWidth);
        Assert.Equal("2500", table.GetFirstChild<TableProperties>().TableWidth.Width.Value);

        wrapper.PreferredWidthType = PreferredWidthType.Auto;
        Assert.Equal(0, wrapper.PreferredWidth);
        Assert.Throws<ArgumentOutOfRangeException>(() => wrapper.PreferredWidth = -1);
    }

    [Fact]
    public void ColumnIndexIsItsPositionInTheGrid()
    {
        Table table = CreateTable(1, 3);
        Assert.Equal(new[] { 0, 1, 2 }, new TableWrapper(table).Columns.Select(c => c.Index).ToArray());
    }

    [Fact]
    public void ColumnDeleteRemovesTheGridColumnAndTheCells()
    {
        Table table = CreateTable(2, 3);
        new TableWrapper(table).Columns[1].Delete();

        Assert.Equal(2, table.GetFirstChild<TableGrid>().Elements<GridColumn>().Count());
        Assert.All(table.Elements<TableRow>(), r => Assert.Equal(new[] { 0, 2 }, r.Elements<TableCell>().Select(c => Convert.ToInt32(c.InnerText.Substring(2), System.Globalization.CultureInfo.InvariantCulture)).ToArray()));
    }

    [Fact]
    public void DeletingTheOnlyColumnDeletesTheTable()
    {
        Body body = new Body();
        Table table = CreateTable(2, 1);
        body.AppendChild(table);

        new TableWrapper(table).Columns[0].Delete();

        Assert.Empty(body.Elements<Table>());
    }

    [Fact]
    public void ColumnDeleteNarrowsACellThatSpansTheColumn()
    {
        Table table = CreateTable(2, 3);
        TableCell spanning = table.Descendants<TableCell>().First();
        spanning.TableCellProperties = new TableCellProperties(new GridSpan() { Val = 2 });
        spanning.NextSibling<TableCell>().Remove();

        new TableWrapper(table).Columns[1].Delete();

        Assert.Equal(2, table.Descendants<GridColumn>().Count());
        Assert.Equal(1, spanning.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal(2, table.Elements<TableRow>().First().Elements<TableCell>().Count());
        Assert.Equal(2, table.Elements<TableRow>().Last().Elements<TableCell>().Count());
    }

    [Fact]
    public void ColumnDeletePassesTheStartOfAVerticalMergeToTheNextRow()
    {
        Table table = CreateTable(2, 2);
        new TableWrapper(table).Cell(0, 0).MergeDown();

        new TableWrapper(table).Columns[0].Delete();

        Assert.Single(table.Elements<TableRow>().First().Elements<TableCell>());
        Assert.Single(table.Elements<TableRow>().Last().Elements<TableCell>());
        Assert.Null(table.Elements<TableRow>().Last().Elements<TableCell>().Single().TableCellProperties?.VerticalMerge);
    }
}
