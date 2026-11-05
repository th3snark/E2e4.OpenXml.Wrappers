using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The formatting that <c>Cells</c>, <c>Rows</c> and <c>Columns</c> apply to all their members, the nesting level, and the columns that
/// follow each other. The markup is what Word 16 wrote (the cells-collection-format, columns-collection-format and rows-set-height scenarios
/// of MarkupConformanceTest).
/// </summary>
public class TableCollectionsTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly TableWrapper _table;

    public TableCollectionsTest()
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

    private Body Body
    {
        get { return _document.WordprocessingDocument.MainDocumentPart.Document.Body; }
    }

    [Fact]
    public void CellsShading_OfAllTheCellsOfATable_ShadesTheTableToo()
    {
        _table.Range.Cells.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);

        Assert.Equal(10, Body.Descendants<Shading>().Count());
        Assert.NotNull(_table.Table.GetFirstChild<TableProperties>().Shading);
        Assert.Equal(ColorHelper.RGB(255, 0, 0), (int)_table.Shading.BackgroundPatternColor.Value.Value);
    }

    [Fact]
    public void CellsShading_OfSomeCells_ShadesOnlyThose()
    {
        _table.Rows[1].Cells.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);

        Assert.Equal(3, Body.Descendants<Shading>().Count());
        Assert.Null(_table.Table.GetFirstChild<TableProperties>().Shading);
    }

    [Fact]
    public void ColumnsShading_OfAllTheColumns_ShadesTheTableToo()
    {
        _table.Columns.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);

        Assert.Equal(10, Body.Descendants<Shading>().Count());
        Assert.NotNull(_table.Table.GetFirstChild<TableProperties>().Shading);
    }

    [Fact]
    public void CellsVerticalAlignment_ReadsNullWhereTheCellsDiffer()
    {
        _table.Cell(0, 0).VerticalAlignment = CellVerticalAlignment.Center;

        Assert.Null(_table.Range.Cells.VerticalAlignment);

        _table.Range.Cells.VerticalAlignment = CellVerticalAlignment.Bottom;
        Assert.Equal(CellVerticalAlignment.Bottom, _table.Range.Cells.VerticalAlignment);
        Assert.All(_table.Range.Cells, c => Assert.Equal(CellVerticalAlignment.Bottom, c.VerticalAlignment));
    }

    [Fact]
    public void CellsHeight_IsTheHeightOfTheirRows()
    {
        _table.Rows[0].Height = 20;

        Assert.Null(_table.Range.Cells.Height);

        _table.Range.Cells.Height = 30;
        Assert.Equal(30, _table.Range.Cells.Height);
        Assert.All(_table.Rows, r => Assert.Equal(30, r.Height));
    }

    [Fact]
    public void CellsSetHeight_SetsTheRowsOfTheCellsOnce()
    {
        _table.Rows[1].Cells.SetHeight(40, RowHeightRule.Exactly);

        Assert.Equal(40, _table.Rows[1].Height);
        Assert.Equal(RowHeightRule.Exactly, _table.Rows[1].HeightRule);
        Assert.Equal(RowHeightRule.Exactly, _table.Rows[1].Cells.HeightRule);
        Assert.NotEqual(40, _table.Rows[0].Height);
    }

    [Fact]
    public void CellsPreferredWidth_InPoints_BringsTheGridInLine_WhenTheyAreWholeColumns()
    {
        TableCellWrapperCollection cells = _table.Range.Cells;

        cells.PreferredWidthType = PreferredWidthType.Points;
        cells.PreferredWidth = 80;

        Assert.Equal(80, cells.PreferredWidth);
        Assert.All(_table.Columns, c => Assert.Equal(80, c.Width));
        Assert.All(_table.Range.Cells, c => Assert.Equal(1600, c.WidthTwips));
    }

    [Fact]
    public void CellsPreferredWidth_OfPartOfAColumn_LeavesTheGrid()
    {
        double before = _table.Columns[0].Width;

        _table.Rows[0].Cells.PreferredWidthType = PreferredWidthType.Points;
        _table.Rows[0].Cells.PreferredWidth = 50;

        Assert.Equal(before, _table.Columns[0].Width);
        Assert.Equal(50, _table.Cell(0, 0).PreferredWidth);
    }

    [Fact]
    public void ColumnsWidth_SetsEveryColumn_AndReadsNullWhereTheyDiffer()
    {
        _table.Columns[1].Width = 120;
        Assert.Null(_table.Columns.Width);

        _table.Columns.Width = 100;

        Assert.Equal(100, _table.Columns.Width);
        Assert.All(_table.Range.Cells, c => Assert.Equal(2000, c.WidthTwips));
    }

    [Fact]
    public void ColumnsPreferredWidth_InPoints_IsTheWidthOfTheColumns()
    {
        _table.Columns.PreferredWidthType = PreferredWidthType.Points;
        _table.Columns.PreferredWidth = 90;

        Assert.All(_table.Columns, c => Assert.Equal(90, c.Width));
        Assert.Equal(90, _table.Columns.PreferredWidth);
    }

    [Fact]
    public void ColumnsPreferredWidth_InPercent_LeavesTheGrid()
    {
        double before = _table.Columns[0].Width;

        _table.Columns.PreferredWidthType = PreferredWidthType.Percent;
        _table.Columns.PreferredWidth = 25;

        Assert.Equal(25, _table.Columns.PreferredWidth);
        Assert.Equal(PreferredWidthType.Percent, _table.Columns.PreferredWidthType);
        Assert.Equal(before, _table.Columns[0].Width);
    }

    [Fact]
    public void RowsSetHeight_SetsEveryRow()
    {
        _table.Rows.SetHeight(40, RowHeightRule.Exactly);
        _table.Cell(2, 0).SetHeight(25, RowHeightRule.AtLeast);

        Assert.Equal(40, _table.Rows[0].Height);
        Assert.Equal(RowHeightRule.Exactly, _table.Rows[1].HeightRule);
        Assert.Equal(25, _table.Rows[2].Height);
        Assert.Equal(RowHeightRule.AtLeast, _table.Rows[2].HeightRule);
    }

    [Fact]
    public void RowsDelete_DeletesEveryRow_AndSoTheTable()
    {
        _table.Rows.Delete();

        Assert.Empty(Body.Descendants<Table>());
    }

    [Fact]
    public void ColumnsDelete_DeletesEveryColumn_AndSoTheTable()
    {
        _table.Columns.Delete();

        Assert.Empty(Body.Descendants<Table>());
    }

    [Fact]
    public void Columns_KnowTheColumnsBesideThem()
    {
        TableColumnWrapper middle = _table.Columns[1];

        Assert.Equal(_table.Columns[2], middle.Next());
        Assert.Equal(_table.Columns[0], middle.Previous());
        Assert.Null(_table.Columns[2].Next());
        Assert.Null(_table.Columns[0].Previous());
    }

    [Fact]
    public void NestingLevel_IsThatOfTheTable()
    {
        TableWrapper inner = _document.Tables.Add(_table.Cell(0, 0).Range, 1, 2);

        Assert.Equal(1, _table.Cell(0, 1).NestingLevel);
        Assert.Equal(1, _table.Rows[0].NestingLevel);
        Assert.Equal(1, _table.Columns[0].NestingLevel);
        Assert.Equal(1, _table.Columns.NestingLevel);
        Assert.Equal(2, inner.Cell(0, 0).NestingLevel);
        Assert.Equal(2, inner.Rows[0].NestingLevel);
        Assert.Equal(2, inner.Columns[1].NestingLevel);
    }
}
