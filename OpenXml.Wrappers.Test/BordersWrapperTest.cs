using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class BordersWrapperTest
{
    private static Table CreateTable()
    {
        return new Table(new TableProperties(new TableWidth() { Width = "100", Type = TableWidthUnitValues.Dxa }), new TableGrid(new GridColumn()), new TableRow(new TableCell(new Paragraph())));
    }

    private static string[] Names(TableBorders borders)
    {
        return borders.ChildElements.Select(e => e.LocalName).ToArray();
    }

    [Fact]
    public void AnUnsetBorderReadsAsNone()
    {
        BorderWrapper border = new TableWrapper(CreateTable()).Borders[BorderPosition.Top];

        Assert.Equal(BorderLineStyle.None, border.LineStyle);
        Assert.Equal(0, border.LineWidth);
        Assert.Equal("auto", border.Color);
    }

    [Fact]
    public void ReadingDoesNotCreateAnything()
    {
        Table table = CreateTable();
        _ = new TableWrapper(table).Borders[BorderPosition.Top].LineStyle;

        Assert.Null(table.GetFirstChild<TableProperties>().TableBorders);
    }

    [Fact]
    public void SettingTheStyleGivesWordsDefaults()
    {
        Table table = CreateTable();
        new TableWrapper(table).Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        TopBorder top = table.GetFirstChild<TableProperties>().TableBorders.TopBorder;
        Assert.Equal(BorderValues.Single, top.Val.Value);
        Assert.Equal(4u, top.Size.Value);
        Assert.Equal("auto", top.Color.Value);
    }

    [Fact]
    public void WidthIsInPointsAndValidated()
    {
        Table table = CreateTable();
        BorderWrapper border = new TableWrapper(table).Borders[BorderPosition.Bottom];
        border.LineWidth = 1.5;

        Assert.Equal(1.5, border.LineWidth);
        Assert.Equal(12u, table.GetFirstChild<TableProperties>().TableBorders.BottomBorder.Size.Value);
        Assert.Equal(BorderLineStyle.SingleLine, border.LineStyle);
        Assert.Throws<ArgumentOutOfRangeException>(() => border.LineWidth = 0.1);
        Assert.Throws<ArgumentOutOfRangeException>(() => border.LineWidth = 13);
    }

    [Fact]
    public void ColourIsValidated()
    {
        Table table = CreateTable();
        BorderWrapper border = new TableWrapper(table).Borders[BorderPosition.Left];
        border.Color = "C00000";

        Assert.Equal("C00000", border.Color);
        Assert.Equal(BorderLineStyle.SingleLine, border.LineStyle);
        Assert.Throws<ArgumentException>(() => border.Color = "red");
        Assert.Throws<ArgumentException>(() => border.Color = null);
        border.Color = "auto";
        Assert.Equal("auto", border.Color);
    }

    [Fact]
    public void BordersAreWrittenInSchemaOrderWhateverTheOrderTheyAreSet()
    {
        Table table = CreateTable();
        BordersWrapper borders = new TableWrapper(table).Borders;
        borders[BorderPosition.Vertical].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.Right].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.Horizontal].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.Left].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(new[] { "top", "left", "bottom", "right", "insideH", "insideV" }, Names(table.GetFirstChild<TableProperties>().TableBorders));
    }

    [Fact]
    public void TablePropertiesStayInSchemaOrder()
    {
        Table table = CreateTable();
        table.GetFirstChild<TableProperties>().AppendChild(new TableLayout() { Type = TableLayoutValues.Fixed });
        new TableWrapper(table).Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(new[] { "tblW", "tblBorders", "tblLayout" }, table.GetFirstChild<TableProperties>().ChildElements.Select(e => e.LocalName).ToArray());
    }

    [Fact]
    public void ANoneStyleIsAnExplicitNilAndClearRemovesIt()
    {
        Table table = CreateTable();
        BorderWrapper border = new TableWrapper(table).Borders[BorderPosition.Top];
        border.LineStyle = BorderLineStyle.DoubleLine;
        border.LineStyle = BorderLineStyle.None;

        Assert.Equal(BorderValues.Nil, table.GetFirstChild<TableProperties>().TableBorders.TopBorder.Val.Value);
        Assert.Equal(BorderLineStyle.None, border.LineStyle);

        border.Clear();
        Assert.Null(table.GetFirstChild<TableProperties>().TableBorders.TopBorder);
    }

    [Fact]
    public void OtherCannotBeSetButIsReadForUnknownStyles()
    {
        Table table = CreateTable();
        BorderWrapper border = new TableWrapper(table).Borders[BorderPosition.Top];
        Assert.Throws<ArgumentException>(() => border.LineStyle = BorderLineStyle.Other);

        border.LineStyle = BorderLineStyle.SingleLine;
        table.GetFirstChild<TableProperties>().TableBorders.TopBorder.Val = BorderValues.Thick;
        Assert.Equal(BorderLineStyle.Other, border.LineStyle);
    }

    [Fact]
    public void SetOutsideAndSetInsideSetAllTheBorders()
    {
        Table table = CreateTable();
        BordersWrapper borders = new TableWrapper(table).Borders;
        borders.SetOutside(BorderLineStyle.SingleLine, 1, "000000");
        borders.SetInside(BorderLineStyle.Dotted, 0.5, "808080");

        Assert.Equal(new[] { "top", "left", "bottom", "right", "insideH", "insideV" }, Names(table.GetFirstChild<TableProperties>().TableBorders));
        Assert.Equal(1, borders[BorderPosition.Right].LineWidth);
        Assert.Equal(BorderLineStyle.Dotted, borders[BorderPosition.Vertical].LineStyle);
        Assert.Equal("808080", borders[BorderPosition.Horizontal].Color);

        borders.Clear();
        Assert.Null(table.GetFirstChild<TableProperties>().TableBorders);
    }

    [Fact]
    public void ATableWithoutPropertiesGetsThemFirst()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        new TableWrapper(table).Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;

        Assert.IsType<TableProperties>(table.FirstChild);
    }

    [Fact]
    public void ATableHasNoDiagonalBorders()
    {
        Assert.Throws<NotSupportedException>(() => new TableWrapper(CreateTable()).Borders[BorderPosition.DiagonalDown]);
        Assert.Throws<ArgumentOutOfRangeException>(() => new TableWrapper(CreateTable()).Borders[(BorderPosition)99]);
    }

    [Fact]
    public void CellBordersIncludeDiagonalsAndKeepTheCellPropertiesInOrder()
    {
        TableCell cell = new TableCell(new TableCellProperties(new TableCellWidth() { Width = "10", Type = TableWidthUnitValues.Dxa }, new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }), new Paragraph());
        BordersWrapper borders = new TableCellWrapper(cell).Borders;
        borders[BorderPosition.DiagonalUp].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.DiagonalDown].LineStyle = BorderLineStyle.SingleLine;
        borders[BorderPosition.Top].LineStyle = BorderLineStyle.DashLargeGap;

        TableCellBorders cellBorders = cell.TableCellProperties.TableCellBorders;
        Assert.Equal(new[] { "top", "tl2br", "tr2bl" }, cellBorders.ChildElements.Select(e => e.LocalName).ToArray());
        Assert.Equal(new[] { "tcW", "tcBorders", "vAlign" }, cell.TableCellProperties.ChildElements.Select(e => e.LocalName).ToArray());
        Assert.Equal(BorderLineStyle.DashLargeGap, borders[BorderPosition.Top].LineStyle);
    }

    [Fact]
    public void ACellWithoutPropertiesGetsThemFirst()
    {
        TableCell cell = new TableCell(new Paragraph());
        new TableCellWrapper(cell).Borders[BorderPosition.Left].LineStyle = BorderLineStyle.SingleLine;

        Assert.IsType<TableCellProperties>(cell.FirstChild);
    }

    [Fact]
    public void StartAndEndElementsAreReadAsLeftAndRight()
    {
        TableCell cell = new TableCell(new TableCellProperties(new TableCellBorders(new StartBorder() { Val = BorderValues.Single, Size = 8 }, new EndBorder() { Val = BorderValues.Double, Size = 16 })), new Paragraph());
        BordersWrapper borders = new TableCellWrapper(cell).Borders;

        Assert.Equal(BorderLineStyle.SingleLine, borders[BorderPosition.Left].LineStyle);
        Assert.Equal(2, borders[BorderPosition.Right].LineWidth);

        borders[BorderPosition.Left].LineWidth = 3;
        Assert.Single(cell.TableCellProperties.TableCellBorders.Elements<StartBorder>());
        Assert.Empty(cell.TableCellProperties.TableCellBorders.Elements<LeftBorder>());
    }
}
