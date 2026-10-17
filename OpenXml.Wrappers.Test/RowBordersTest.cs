using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class RowBordersTest
{
    private static TableRow CreateRow(int cells)
    {
        return new TableRow(Enumerable.Range(0, cells).Select(i => new TableCell(new Paragraph(new Run(new Text(i.ToString(System.Globalization.CultureInfo.InvariantCulture)))))));
    }

    private static string[] BorderNames(TableCell cell)
    {
        return cell.TableCellProperties?.TableCellBorders?.ChildElements.Select(e => e.LocalName).ToArray() ?? Array.Empty<string>();
    }

    [Fact]
    public void TopAndBottomAreSetOnEveryCell()
    {
        TableRow row = CreateRow(3);
        TableRowWrapper wrapper = new TableRowWrapper(row);

        wrapper.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
        wrapper.Borders[BorderPosition.Bottom].LineWidth = 1.5;
        wrapper.Borders[BorderPosition.Bottom].Color = "FF0000";

        Assert.All(row.Elements<TableCell>(), c => Assert.Equal(new[] { "top", "bottom" }, BorderNames(c)));
        Assert.Equal(BorderLineStyle.SingleLine, wrapper.Borders[BorderPosition.Top].LineStyle);
        Assert.Equal(1.5, wrapper.Borders[BorderPosition.Bottom].LineWidth);
        Assert.Equal("FF0000", wrapper.Borders[BorderPosition.Bottom].Color);
    }

    [Fact]
    public void LeftIsTheFirstCellAndRightTheLast()
    {
        TableRow row = CreateRow(3);
        TableRowWrapper wrapper = new TableRowWrapper(row);

        wrapper.Borders[BorderPosition.Left].LineStyle = BorderLineStyle.DoubleLine;
        wrapper.Borders[BorderPosition.Right].LineStyle = BorderLineStyle.Dotted;

        TableCell[] cells = row.Elements<TableCell>().ToArray();
        Assert.Equal(new[] { "left" }, BorderNames(cells[0]));
        Assert.Empty(BorderNames(cells[1]));
        Assert.Equal(new[] { "right" }, BorderNames(cells[2]));
        Assert.Equal(BorderLineStyle.DoubleLine, wrapper.Borders[BorderPosition.Left].LineStyle);
        Assert.Equal(BorderLineStyle.Dotted, wrapper.Borders[BorderPosition.Right].LineStyle);
    }

    [Fact]
    public void VerticalIsTheLineBetweenTheCells()
    {
        TableRow row = CreateRow(3);
        TableRowWrapper wrapper = new TableRowWrapper(row);

        wrapper.Borders[BorderPosition.Vertical].LineStyle = BorderLineStyle.Dashed;

        TableCell[] cells = row.Elements<TableCell>().ToArray();
        Assert.Equal(new[] { "right" }, BorderNames(cells[0]));
        Assert.Equal(new[] { "left", "right" }, BorderNames(cells[1]));
        Assert.Equal(new[] { "left" }, BorderNames(cells[2]));
        Assert.Equal(BorderLineStyle.Dashed, wrapper.Borders[BorderPosition.Vertical].LineStyle);
    }

    [Fact]
    public void SetOutsideAndSetInsideUseTheRowRules()
    {
        TableRow row = CreateRow(2);
        TableRowWrapper wrapper = new TableRowWrapper(row);

        wrapper.Borders.SetOutside(BorderLineStyle.SingleLine, 2, "0000FF");
        wrapper.Borders.SetInside(BorderLineStyle.Dotted, 0.5, "auto");

        TableCell[] cells = row.Elements<TableCell>().ToArray();
        Assert.Equal(new[] { "top", "left", "bottom", "right" }, BorderNames(cells[0]));
        Assert.Equal(BorderLineStyle.Dotted, new TableCellWrapper(cells[0]).Borders[BorderPosition.Right].LineStyle);
        Assert.Equal(2, new TableCellWrapper(cells[0]).Borders[BorderPosition.Top].LineWidth);
    }

    [Fact]
    public void MixedValuesReadAsOther()
    {
        TableRow row = CreateRow(2);
        TableRowWrapper wrapper = new TableRowWrapper(row);
        new CellBorders(row, 0).Top.LineStyle = BorderLineStyle.SingleLine;
        new CellBorders(row, 1).Top.LineStyle = BorderLineStyle.Dotted;

        Assert.Equal(BorderLineStyle.Other, wrapper.Borders[BorderPosition.Top].LineStyle);
        Assert.Equal(0.5, wrapper.Borders[BorderPosition.Top].LineWidth);

        new CellBorders(row, 1).Top.Color = "FF0000";
        Assert.Null(wrapper.Borders[BorderPosition.Top].Color);
    }

    [Fact]
    public void ClearRemovesTheDirectBordersOfTheCells()
    {
        TableRow row = CreateRow(2);
        TableRowWrapper wrapper = new TableRowWrapper(row);
        wrapper.Borders.SetOutside(BorderLineStyle.SingleLine, 1, "auto");

        wrapper.Borders[BorderPosition.Top].Clear();
        Assert.All(row.Elements<TableCell>(), c => Assert.DoesNotContain("top", BorderNames(c)));

        wrapper.Borders.Clear();
        Assert.All(row.Elements<TableCell>(), c => Assert.Empty(BorderNames(c)));
        Assert.Equal(BorderLineStyle.None, wrapper.Borders[BorderPosition.Left].LineStyle);
    }

    [Fact]
    public void ARowHasNoHorizontalOrDiagonalBorders()
    {
        BordersWrapper borders = new TableRowWrapper(CreateRow(2)).Borders;

        Assert.Throws<NotSupportedException>(() => borders[BorderPosition.Horizontal]);
        Assert.Throws<NotSupportedException>(() => borders[BorderPosition.DiagonalDown]);
    }

    [Fact]
    public void ARowWithoutCellsReadsAsNoneAndIgnoresSets()
    {
        TableRowWrapper wrapper = new TableRowWrapper(new TableRow());

        wrapper.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(BorderLineStyle.None, wrapper.Borders[BorderPosition.Top].LineStyle);
        Assert.Equal("auto", wrapper.Borders[BorderPosition.Top].Color);
    }

    [Fact]
    public void ReadingDoesNotCreateCellProperties()
    {
        TableRow row = CreateRow(2);
        _ = new TableRowWrapper(row).Borders[BorderPosition.Top].LineStyle;

        Assert.All(row.Elements<TableCell>(), c => Assert.Null(c.TableCellProperties));
    }

    //Test helper to reach the borders of one cell of a row.
    private sealed class CellBorders
    {
        private readonly BordersWrapper _borders;

        public CellBorders(TableRow row, int index)
        {
            _borders = new TableCellWrapper(row.Elements<TableCell>().ElementAt(index)).Borders;
        }

        public BorderWrapper Top
        {
            get { return _borders[BorderPosition.Top]; }
        }
    }
}
