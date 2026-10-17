using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class NavigationTest
{
    private static Body CreateBody()
    {
        Table table = new Table(
            new TableRow(new TableCell(new Paragraph(new Run(new Text("a")))), new TableCell(new Paragraph(new Run(new Text("b"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("c")))), new TableCell(new Paragraph(new Run(new Text("d"))))));
        return new Body(new Paragraph(new Run(new Text("first"))), table, new Paragraph(new Run(new Text("last"))));
    }

    [Fact]
    public void ParagraphNextAndPreviousCrossTableCells()
    {
        Body body = CreateBody();
        ParagraphWrapper first = new ParagraphWrapper(body.Elements<Paragraph>().First());
        ParagraphWrapper last = new ParagraphWrapper(body.Elements<Paragraph>().Last());

        Assert.Null(first.Previous());
        Assert.StartsWith("a", first.Next().Text, System.StringComparison.Ordinal);
        Assert.Null(last.Next());
        Assert.StartsWith("d", last.Previous().Text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void RowNextPreviousAndEnds()
    {
        Table table = CreateBody().Elements<Table>().Single();
        TableRowWrapper first = new TableRowWrapper(table.Elements<TableRow>().First());
        TableRowWrapper last = new TableRowWrapper(table.Elements<TableRow>().Last());

        Assert.True(first.IsFirst);
        Assert.False(first.IsLast);
        Assert.Null(first.Previous());
        Assert.Equal(last, first.Next());
        Assert.Equal(first, last.Previous());
        Assert.True(last.IsLast);
        Assert.Null(last.Next());
    }

    [Fact]
    public void CellNextAndPreviousWrapAcrossRows()
    {
        Table table = CreateBody().Elements<Table>().Single();
        TableCellWrapper a = new TableCellWrapper(table.Descendants<TableCell>().First());
        TableCellWrapper b = new TableCellWrapper(table.Descendants<TableCell>().ElementAt(1));
        TableCellWrapper c = new TableCellWrapper(table.Descendants<TableCell>().ElementAt(2));
        TableCellWrapper d = new TableCellWrapper(table.Descendants<TableCell>().Last());

        Assert.Null(a.Previous());
        Assert.Equal(b, a.Next());
        Assert.Equal(c, b.Next());
        Assert.Equal(b, c.Previous());
        Assert.Null(d.Next());
    }

    [Fact]
    public void RangeDuplicateIsAnEqualButSeparateRange()
    {
        RangeWrapper range = new RangeWrapper(CreateBody());
        RangeWrapper copy = range.Duplicate();

        Assert.NotSame(range, copy);
        Assert.Equal(range, copy);
    }
}
