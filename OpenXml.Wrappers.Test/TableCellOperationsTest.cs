using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>Merge</c>, <c>Split</c>, <c>Add</c> and <c>Delete</c> of <c>Cells</c>, and deleting a cell with the cells below it moving up. The markup
/// is what Word 16 wrote (the cells-merge-row, cell-delete-up, cells-delete-row-up, cells-add-*, and cells-split-* scenarios of
/// MarkupConformanceTest).
/// </summary>
public class TableCellOperationsTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly TableWrapper _table;

    public TableCellOperationsTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _table = _document.Tables.Add(_document.Range(0, 0), 3, 3);
        int number = 1;
        foreach (TableCellWrapper cell in _table.Range.Cells)
        {
            cell.Range.Text = "c" + number.ToString(CultureInfo.InvariantCulture);
            number++;
        }
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private string Texts(int row)
    {
        return string.Join("|", _table.Table.Elements<TableRow>().ElementAt(row).Elements<TableCell>().Select(c => string.Concat(c.Descendants<Text>().Select(t => t.Text))));
    }

    private int RowCount
    {
        get { return _table.Table.Elements<TableRow>().Count(); }
    }

    [Fact]
    public void CellsMerge_MakesOneCellOfTheRectangleTheCellsFill()
    {
        _table.Rows[1].Cells.Merge();

        Assert.Equal("c4c5c6", Texts(1));
        Assert.Equal(3, _table.Table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().First().TableCellProperties.GridSpan.Val.Value);
        Assert.Equal("c1|c2|c3", Texts(0));
    }

    [Fact]
    public void CellsMerge_OfAWholeTable_MakesOneCell()
    {
        _table.Range.Cells.Merge();

        Assert.Equal(1, _table.Table.Descendants<TableCell>().Count(c => c.TableCellProperties?.VerticalMerge == null || c.TableCellProperties.VerticalMerge.Val?.Value == MergedCellValues.Restart));
        Assert.Contains("c9", Texts(0));
    }

    [Fact]
    public void DeleteUp_MovesTheCellsBelowUp_AndLeavesTheBottomOneEmpty()
    {
        _table.Cell(1, 1).Delete(ShiftCells.Up);

        Assert.Equal("c1|c2|c3", Texts(0));
        Assert.Equal("c4|c8|c6", Texts(1));
        Assert.Equal("c7||c9", Texts(2));
        Assert.True(_table.Table.Descendants<TableCell>().All(c => c.Elements<Paragraph>().Any()));
    }

    [Fact]
    public void CellsDeleteUp_OfARow_MovesTheRowBelowUp()
    {
        _table.Rows[1].Cells.Delete(ShiftCells.Up);

        Assert.Equal("c1|c2|c3", Texts(0));
        Assert.Equal("c7|c8|c9", Texts(1));
        Assert.Equal("||", Texts(2));
    }

    [Fact]
    public void CellsDelete_EntireRow_DeletesTheRows()
    {
        _table.Rows[1].Cells.Delete(ShiftCells.EntireRow);

        Assert.Equal(2, RowCount);
        Assert.Equal("c7|c8|c9", Texts(1));
    }

    [Fact]
    public void CellsDelete_Left_OfAWholeRow_DeletesTheRow()
    {
        _table.Rows[1].Cells.Delete();

        Assert.Equal(2, RowCount);
    }

    [Fact]
    public void CellsDelete_Left_OfAllTheCells_DeletesTheColumns_AndSoTheTable()
    {
        _table.Range.Cells.Delete(ShiftCells.Left);

        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Table>());
    }

    [Fact]
    public void CellsAdd_AtTheEnd_AddsACellAndGrowsTheGrid()
    {
        TableCellWrapper added = _table.Rows[1].Cells.Add();

        Assert.Equal(4, _table.Rows[1].Cells.Count);
        Assert.Equal(added, _table.Rows[1].Cells[3]);
        Assert.Equal(4, _table.Columns.Count);
        Assert.Equal(1, _table.Table.Elements<TableRow>().First().TableRowProperties.GetFirstChild<GridAfter>().Val.Value);
    }

    [Fact]
    public void CellsAdd_BeforeACell_PutsTheNewCellInFrontOfIt()
    {
        _table.Rows[1].Cells.Add(_table.Cell(1, 1));

        Assert.Equal("c4||c5|c6", Texts(1));
    }

    [Fact]
    public void CellsAdd_Rejects_ACellOfAnotherRow_AndCellsThatAreNotARow()
    {
        Assert.Throws<ArgumentException>(() => _table.Rows[1].Cells.Add(_table.Cell(0, 0)));
        Assert.Throws<InvalidOperationException>(() => _table.Range.Cells.Add());
    }

    [Fact]
    public void CellsSplit_IntoColumns_SplitsEveryCell()
    {
        _table.Rows[1].Cells.Split(1, 2);

        Assert.Equal(6, _table.Rows[1].Cells.Count);
        Assert.Equal(6, _table.Columns.Count);
    }

    [Fact]
    public void CellsSplit_WholeRowsIntoRows_AddsRows()
    {
        _table.Rows[1].Cells.Split(2, 1, false);

        Assert.Equal(4, RowCount);
        Assert.Equal("c4|c5|c6", Texts(1));
        Assert.Equal("||", Texts(2));
        Assert.All(_table.Table.Descendants<TableCell>(), c => Assert.Null(c.TableCellProperties?.VerticalMerge));
    }

    [Fact]
    public void CellsSplit_WithMergeBeforeSplit_MergesTheCellsFirst()
    {
        _table.Rows[1].Cells.Split(1, 3, true);

        Assert.Equal(3, _table.Rows[1].Cells.Count);
        Assert.Equal(3, _table.Columns.Count);
    }

    [Fact]
    public void CellsSplit_RejectsLessThanOneRowOrColumn()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _table.Rows[1].Cells.Split(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => _table.Rows[1].Cells.Split(1, 0));
    }
}
