using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class TableColumnWrapperCollectionTest
{
    private static Table CreateTable()
    {
        return new Table(
            new TableGrid(new GridColumn() { Width = "1000" }, new GridColumn() { Width = "2000" }),
            new TableRow(
                new TableCell(new Paragraph(new Run(new Text("a")))),
                new TableCell(
                    new TableCellProperties(new TableCellWidth() { Width = "2000", Type = TableWidthUnitValues.Dxa }, new GridSpan() { Val = 2 }),
                    new Paragraph(new BookmarkStart() { Id = "1", Name = "b" }, new Run(new Text("b")), new BookmarkEnd() { Id = "1" }))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("c"))))));
    }

    [Fact]
    public void CountIsTheNumberOfGridColumns()
    {
        Assert.Equal(2, new TableWrapper(CreateTable()).Columns.Count);
    }

    [Fact]
    public void AddAppendsAGridColumnAndAnEmptyCellToEachRow()
    {
        Table table = CreateTable();
        TableColumnWrapper column = new TableWrapper(table).Columns.Add();

        Assert.Equal(3, table.Elements<TableGrid>().Single().Elements<GridColumn>().Count());
        Assert.Equal(100, column.Width);

        TableCell added = table.Elements<TableRow>().First().Elements<TableCell>().Last();
        Assert.Equal(String.Empty, new RangeWrapper(added).Text.TrimEnd('\r', '\a'));
    }

    [Fact]
    public void AddedCellsSpanOnlyTheNewColumnAndTakeItsWidth()
    {
        Table table = CreateTable();
        new TableWrapper(table).Columns.Add();

        TableCell added = table.Elements<TableRow>().First().Elements<TableCell>().Last();
        Assert.Null(added.TableCellProperties.GridSpan);
        Assert.Equal("2000", added.TableCellProperties.TableCellWidth.Width.Value);
    }

    [Fact]
    public void AddDoesNotDuplicateBookmarkAnchors()
    {
        Table table = CreateTable();
        new TableWrapper(table).Columns.Add();

        Assert.Single(table.Descendants<BookmarkStart>());
        Assert.Single(table.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void AddSkipsRowsWithoutCells()
    {
        Table table = CreateTable();
        table.AppendChild(new TableRow());

        new TableWrapper(table).Columns.Add();

        Assert.Empty(table.Elements<TableRow>().Last().Elements<TableCell>());
    }

    [Fact]
    public void AddOnAnEmptyGridThrows()
    {
        Table table = new Table(new TableGrid());

        Assert.Throws<InvalidOperationException>(() => new TableWrapper(table).Columns.Add());
    }
}
