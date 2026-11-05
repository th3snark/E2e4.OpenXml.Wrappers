using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The width of a single cell, <c>SetWidth</c> and <c>DistributeWidth</c>. The markup is what Word 16 wrote (the cell-width-ragged,
/// cell-set-width-*, cells-set-width-row, cells-distribute-width, columns-set-width and columns-distribute-width scenarios of
/// MarkupConformanceTest).
/// </summary>
public class TableWidthsTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly TableWrapper _table;

    public TableWidthsTest()
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

    private int[] Grid
    {
        get { return _table.Table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(c => int.Parse(c.Width.Value, CultureInfo.InvariantCulture)).ToArray(); }
    }

    private int[] RowWidths(int row)
    {
        return _table.Table.Elements<TableRow>().ElementAt(row).Elements<TableCell>().Select(c => int.Parse(c.TableCellProperties.TableCellWidth.Width.Value, CultureInfo.InvariantCulture)).ToArray();
    }

    private int[] Spans(int row)
    {
        return _table.Table.Elements<TableRow>().ElementAt(row).Elements<TableCell>().Select(c => c.TableCellProperties?.GridSpan?.Val?.Value ?? 1).ToArray();
    }

    private TableRowProperties RowProperties(int row)
    {
        return _table.Table.Elements<TableRow>().ElementAt(row).TableRowProperties;
    }

    [Fact]
    public void CellWidth_OfOneCell_SplitsTheGridWhereTheCellsEnd()
    {
        _table.Cell(1, 1).Width = 100;

        Assert.Equal(new[] { 3008, 2000, 1009, 2000, 1009 }, Grid);
        Assert.Equal(new[] { 3008, 2000, 3009 }, RowWidths(1));
        Assert.Equal(new[] { 1, 1, 2 }, Spans(1));
        Assert.Equal(new[] { 1, 2, 2 }, Spans(0));
    }

    [Fact]
    public void CellWidth_OfAShortRow_LeavesTheGridColumnsAfterItUnused()
    {
        _table.Cell(1, 1).Width = 100;

        Assert.Equal(1, RowProperties(1).GetFirstChild<GridAfter>().Val.Value);
        Assert.Equal("1009", RowProperties(1).GetFirstChild<WidthAfterTableRow>().Width.Value);
        Assert.Null(RowProperties(0));
    }

    [Fact]
    public void CellWidth_SetBackToTheOthers_JoinsTheGridAgain()
    {
        _table.Cell(1, 1).Width = 100;
        _table.Cell(1, 1).Width = _table.Cell(0, 1).Width;

        Assert.Equal(new[] { 3008, 3009, 3009 }, Grid);
        Assert.All(Enumerable.Range(0, 3), r => Assert.All(Spans(r), s => Assert.Equal(1, s)));
        Assert.Null(RowProperties(1)?.GetFirstChild<GridAfter>());
    }

    [Fact]
    public void CellWidth_OfACellOutsideATable_JustSetsTheWidth()
    {
        TableCellWrapper cell = new TableCellWrapper(new TableCell(new Paragraph()));

        cell.Width = 50;

        Assert.Equal(50, cell.Width);
    }

    [Fact]
    public void CellSetWidth_Proportional_SharesTheDifferenceAmongTheCellsAfter()
    {
        _table.Cell(1, 0).SetWidth(100, RulerStyle.Proportional);

        Assert.Equal(new[] { 2000, 3513, 3513 }, RowWidths(1));
        Assert.Equal(new[] { 3008, 3009, 3009 }, RowWidths(0));
    }

    [Fact]
    public void CellSetWidth_FirstColumn_GivesTheDifferenceToTheNextCell()
    {
        _table.Cell(1, 0).SetWidth(100, RulerStyle.FirstColumn);

        Assert.Equal(new[] { 2000, 4017, 3009 }, RowWidths(1));
    }

    [Fact]
    public void CellSetWidth_OfTheLastCell_ChangesNothingElse()
    {
        _table.Cell(1, 2).SetWidth(100, RulerStyle.SameWidth);

        Assert.Equal(new[] { 3008, 3009, 2000 }, RowWidths(1));
    }

    [Fact]
    public void CellsSetWidth_None_GivesEveryCellTheWidth_AndTheGridFollows()
    {
        _table.Rows[1].Cells.SetWidth(100, RulerStyle.None);

        Assert.Equal(new[] { 2000, 2000, 2000 }, RowWidths(1));
        Assert.Equal(new[] { 2000, 1008, 992, 2000, 17, 3009 }, Grid);
    }

    [Fact]
    public void CellsWidth_ReadsNullWhereTheCellsDiffer()
    {
        _table.Range.Cells.Width = 80;
        Assert.Equal(80, _table.Range.Cells.Width);
        Assert.Equal(new[] { 1600, 1600, 1600 }, Grid);

        _table.Cell(0, 0).Width = 90;
        Assert.Null(_table.Range.Cells.Width);
    }

    [Fact]
    public void CellsDistributeWidth_SharesEachRowsWidthEvenly_LastCellTakingTheRemainder()
    {
        _table.Cell(1, 1).Width = 100;

        _table.Range.Cells.DistributeWidth();

        Assert.Equal(new[] { 2672, 2672, 2673 }, RowWidths(1));
        Assert.Equal(new[] { 3008, 3009, 3009 }, RowWidths(0));
    }

    [Fact]
    public void ColumnSetWidth_GivesTheDifferenceToTheColumnsAfter()
    {
        _table.Columns[0].SetWidth(100, RulerStyle.FirstColumn);
        Assert.Equal(new[] { 2000, 4017, 3009 }, Grid);

        _table.Columns[1].SetWidth(120, RulerStyle.Proportional);
        Assert.Equal(new[] { 2000, 2400, 4626 }, Grid);
        Assert.All(Enumerable.Range(0, 3), r => Assert.Equal(new[] { 2000, 2400, 4626 }, RowWidths(r)));
    }

    [Fact]
    public void ColumnSetWidth_OfTheLastColumn_ChangesNothingElse()
    {
        _table.Columns[2].SetWidth(100, RulerStyle.Proportional);

        Assert.Equal(new[] { 3008, 3009, 2000 }, Grid);
    }

    [Fact]
    public void ColumnSetWidth_None_IsTheSameAsWidth()
    {
        _table.Columns[1].SetWidth(100, RulerStyle.None);

        Assert.Equal(new[] { 3008, 2000, 3009 }, Grid);
    }

    [Fact]
    public void ColumnsDistributeWidth_SharesTheWidthInWholeTwips()
    {
        _table.Columns[0].Width = 100;

        _table.Columns.DistributeWidth();

        Assert.Equal(new[] { 2672, 2673, 2673 }, Grid);
    }

    [Fact]
    public void TableWidth_StaysAutomaticUntilTheColumnsAreWiderThanTheMargins()
    {
        _table.Columns.Width = 150.4;
        Assert.Equal(TableWidthUnitValues.Auto, _table.Table.GetFirstChild<TableProperties>().TableWidth.Type.Value);

        _table.Columns.Width = 150.45;
        Assert.Equal(TableWidthUnitValues.Dxa, _table.Table.GetFirstChild<TableProperties>().TableWidth.Type.Value);
    }

    [Fact]
    public void SetWidth_RejectsWidthsOutsideTheRange()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => _table.Cell(0, 0).SetWidth(-1, RulerStyle.None));
        Assert.Throws<ArgumentOutOfRangeException>(() => _table.Columns[0].SetWidth(2000, RulerStyle.Proportional));
        Assert.Throws<ArgumentOutOfRangeException>(() => _table.Range.Cells.SetWidth(double.NaN, RulerStyle.None));
    }
}
