using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The borders of <c>Cells</c>, <c>Rows</c> and <c>Columns</c>. The markup is what Word 16 wrote (the cells-borders-table, cells-borders-row,
/// rows-borders and columns-borders scenarios of MarkupConformanceTest).
/// </summary>
public class TableCollectionBordersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly TableWrapper _table;

    public TableCollectionBordersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _table = _document.Tables.Add(_document.Range(0, 0), 3, 3);
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private TableCellBorders CellBorders(int row, int column)
    {
        return _table.Table.Elements<TableRow>().ElementAt(row).Elements<TableCell>().ElementAt(column).TableCellProperties?.TableCellBorders;
    }

    [Fact]
    public void CellsBorders_OfAWholeTable_AreTheBordersOfTheTable()
    {
        _table.Range.Cells.Borders.Enable = true;

        Assert.NotNull(_table.Table.GetFirstChild<TableProperties>().TableBorders);
        Assert.True(_table.Borders.Enable);
        Assert.Null(CellBorders(0, 0));
    }

    [Fact]
    public void RowsAndColumnsBorders_OfAWholeTable_AreTheBordersOfTheTable()
    {
        _table.Rows.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;
        _table.Columns.Borders.InsideLineStyle = BorderLineStyle.SingleLine;

        Assert.True(_table.Borders.Enable);
        Assert.Null(CellBorders(1, 1));
    }

    [Fact]
    public void CellsBorders_OfARow_AreItsEdges_AndTheNeighboursShareTheLines()
    {
        BordersWrapper borders = _table.Rows[1].Cells.Borders;

        borders.OutsideLineStyle = BorderLineStyle.SingleLine;

        Assert.NotNull(CellBorders(1, 0).TopBorder);
        Assert.NotNull(CellBorders(1, 0).BottomBorder);
        Assert.NotNull(CellBorders(1, 0).LeftBorder);
        Assert.Null(CellBorders(1, 0).RightBorder);
        Assert.NotNull(CellBorders(1, 2).RightBorder);
        Assert.NotNull(CellBorders(0, 1).BottomBorder);
        Assert.NotNull(CellBorders(2, 1).TopBorder);
        Assert.Equal(BorderLineStyle.SingleLine, borders.OutsideLineStyle);
    }

    [Fact]
    public void CellsBorders_Inside_AreTheLinesBetweenTheCells()
    {
        _table.Rows[1].Cells.Borders.InsideLineStyle = BorderLineStyle.SingleLine;

        Assert.NotNull(CellBorders(1, 0).RightBorder);
        Assert.NotNull(CellBorders(1, 1).LeftBorder);
        Assert.NotNull(CellBorders(1, 1).RightBorder);
        Assert.Null(CellBorders(1, 1).TopBorder);
    }

    [Fact]
    public void CellsBorders_Clear_RemovesTheirBorders()
    {
        BordersWrapper borders = _table.Rows[1].Cells.Borders;
        borders.OutsideLineStyle = BorderLineStyle.SingleLine;

        borders.Clear();

        Assert.Null(CellBorders(1, 0));
    }

    [Fact]
    public void CellsBorders_HaveNoDiagonals()
    {
        Assert.Throws<NotSupportedException>(() => _table.Rows[1].Cells.Borders[BorderPosition.DiagonalDown]);
    }
}
