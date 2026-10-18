using System;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class LineWidthTest
{
    private static BorderWrapper CreateBorder()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        return new TableWrapper(table).Borders[BorderPosition.Top];
    }

    [Theory]
    [InlineData(BorderLineWidth.Pt025, 0.25)]
    [InlineData(BorderLineWidth.Pt050, 0.5)]
    [InlineData(BorderLineWidth.Pt075, 0.75)]
    [InlineData(BorderLineWidth.Pt100, 1)]
    [InlineData(BorderLineWidth.Pt150, 1.5)]
    [InlineData(BorderLineWidth.Pt225, 2.25)]
    [InlineData(BorderLineWidth.Pt300, 3)]
    [InlineData(BorderLineWidth.Pt450, 4.5)]
    [InlineData(BorderLineWidth.Pt600, 6)]
    public void NamedWidthsAreInPoints(BorderLineWidth named, double points)
    {
        LineWidth width = named;

        Assert.Equal(points, width.Points);
        Assert.Equal(new LineWidth(points), width);
    }

    [Fact]
    public void BordersTakeWidthsInPointsOrByName()
    {
        BorderWrapper border = CreateBorder();

        border.LineWidth = 1.5;
        Assert.Equal(1.5, border.LineWidth.Points);

        border.LineWidth = BorderLineWidth.Pt225;
        Assert.Equal(2.25, border.LineWidth.Points);
    }

    [Theory]
    [InlineData(0.1)]
    [InlineData(12.5)]
    [InlineData(-1)]
    [InlineData(double.NaN)]
    public void WidthsOutOfRangeAreRejected(double points)
    {
        BorderWrapper border = CreateBorder();

        Assert.Throws<ArgumentOutOfRangeException>(() => border.LineWidth = points);
    }

    [Fact]
    public void NoWidthCannotBeAssigned()
    {
        BorderWrapper border = CreateBorder();

        Assert.Equal(LineWidth.None, border.LineWidth);
        Assert.Throws<ArgumentOutOfRangeException>(() => border.LineWidth = LineWidth.None);
    }

    [Fact]
    public void WidthsRoundToEighthsOfAPoint()
    {
        Assert.Equal(0.375, new LineWidth(0.37).Points);
    }

    [Fact]
    public void BordersGroupsTakeNamedWidths()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        BordersWrapper borders = new TableWrapper(table).Borders;

        borders.SetOutside(BorderLineStyle.SingleLine, BorderLineWidth.Pt150, NamedColor.Red);
        borders.InsideLineWidth = BorderLineWidth.Pt050;

        Assert.Equal(1.5, borders.OutsideLineWidth.Points);
        Assert.Equal(0.5, borders.InsideLineWidth.Points);
    }
}
