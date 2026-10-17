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

        Assert.Equal(2000, new TableColumnWrapper(table.Descendants<GridColumn>().Last()).Width);
    }

    [Fact]
    public void SettingTheWidthSkipsRowsWithTooFewCellsAndCellsWithoutProperties()
    {
        Table table = CreateTable();
        TableColumnWrapper column = new TableColumnWrapper(table.Descendants<GridColumn>().Last());

        column.Width = 3000;

        Assert.Equal(3000, column.Width);
        TableCell cell = table.Elements<TableRow>().First().Elements<TableCell>().Last();
        Assert.Equal("3000", cell.TableCellProperties.TableCellWidth.Width.Value);
        Assert.Equal(TableWidthUnitValues.Dxa, cell.TableCellProperties.TableCellWidth.Type.Value);
    }

    [Fact]
    public void SettingTheWidthUpdatesAnExistingCellWidth()
    {
        Table table = CreateTable();
        TableCell cell = table.Elements<TableRow>().First().Elements<TableCell>().First();
        new TableCellWrapper(cell).CellWidth = 10;

        new TableColumnWrapper(table.Descendants<GridColumn>().First()).Width = 4000;

        Assert.Equal(4000, new TableCellWrapper(cell).CellWidth);
    }
}
