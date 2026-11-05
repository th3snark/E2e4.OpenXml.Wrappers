using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>Cell.Formula</c>, <c>Row.LeftIndent</c> and <c>SetLeftIndent</c>, and the conversion of older horizontal merges. The markup of the formulas
/// and indents is what Word 16 wrote (the cell-formula-*, cell-autosum and row-*-left-indent* scenarios of MarkupConformanceTest).
/// </summary>
public class TableFormulaIndentMergeTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableFormulaIndentMergeTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private TableWrapper Fill(int rows, int columns, params string[] values)
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), rows, columns);
        for (int index = 0; index < values.Length; index++)
        {
            if (values[index] != null)
            {
                table.Cell(index / columns, index % columns).Range.Text = values[index];
            }
        }
        return table;
    }

    private static string CellText(TableWrapper table, int row, int column)
    {
        return string.Concat(table.Cell(row, column).TableCell.Descendants<Text>().Select(t => t.Text));
    }

    private static string FieldResult(TableWrapper table, int row, int column)
    {
        TableCell cell = table.Cell(row, column).TableCell;
        SimpleField simple = cell.Descendants<SimpleField>().FirstOrDefault();
        Run result = simple != null ? simple.Elements<Run>().First() : cell.Descendants<Run>().First(r => r.Elements<Text>().Any());
        return result.GetFirstChild<Text>().Text;
    }

    [Fact]
    public void Formula_SumAbove_PutsTheSumInTheCell_AndFormatsIt()
    {
        TableWrapper table = Fill(3, 1, "1", "2", null);

        table.Cell(2, 0).Formula("=SUM(ABOVE)", "0.00");

        Assert.Equal("3.00", FieldResult(table, 2, 0));
        Assert.Contains("SUM(ABOVE)", table.Cell(2, 0).TableCell.Descendants<FieldCode>().Single().Text);
    }

    [Theory]
    [InlineData("AVERAGE", "3")]
    [InlineData("MAX", "6")]
    [InlineData("MIN", "1")]
    [InlineData("COUNT", "3")]
    [InlineData("PRODUCT", "12")]
    public void Formula_Functions_WorkOutTheirResult(string function, string expected)
    {
        TableWrapper table = Fill(4, 1, "1", "2", "6", null);

        table.Cell(3, 0).Formula("=" + function + "(ABOVE)");

        Assert.Equal(expected, FieldResult(table, 3, 0));
    }

    [Fact]
    public void Formula_StopsAtACellWithoutANumber_AsWordDoes()
    {
        TableWrapper table = Fill(4, 1, "1", "x", "2", null);

        table.Cell(3, 0).Formula("=SUM(ABOVE)");

        Assert.Equal("2", FieldResult(table, 3, 0));
    }

    [Fact]
    public void Formula_LeftRightAndBelow()
    {
        TableWrapper table = Fill(2, 3, "1", "2", null, null, null, null);
        table.Cell(0, 2).Formula("=SUM(LEFT)");
        Assert.Equal("3", FieldResult(table, 0, 2));

        TableWrapper below = Fill(3, 1, null, "4", "5");
        below.Cell(0, 0).Formula("=SUM(BELOW)");
        Assert.Equal("9", FieldResult(below, 0, 0));
    }

    [Fact]
    public void AutoSum_UsesTheNumbersAbove_OtherwiseThoseToTheLeft_OtherwiseWordsError()
    {
        TableWrapper table = Fill(3, 2, "1", null, "2", null, null, null);

        table.Cell(2, 0).AutoSum();
        table.Cell(0, 1).AutoSum();

        Assert.Equal("3", FieldResult(table, 2, 0));
        Assert.Equal("1", FieldResult(table, 0, 1));

        TableWrapper alone = Fill(1, 1);
        alone.Cell(0, 0).Formula();
        Assert.Equal("!Unexpected End of Formula", FieldResult(alone, 0, 0).Trim());
    }

    [Fact]
    public void Formula_ThatCannotBeWorkedOutHere_IsLeftForWord()
    {
        TableWrapper table = Fill(2, 1, "2", null);

        table.Cell(1, 0).Formula("=A1*2");

        FieldChar begin = table.Cell(1, 0).TableCell.Descendants<FieldChar>().First(f => f.FieldCharType.Value == FieldCharValues.Begin);
        Assert.True(begin.Dirty.Value);
    }

    [Fact]
    public void RowLeftIndent_LeavesSpaceBeforeTheRow_AndMakesTheTableRagged()
    {
        TableWrapper table = Fill(3, 3);

        table.Rows[1].LeftIndent = 20;

        Assert.Equal(20, table.Rows[1].LeftIndent);
        Assert.Equal(0, table.Rows[0].LeftIndent);
        Assert.Equal(7, table.Columns.Count);
        Assert.Equal(1, table.Table.Elements<TableRow>().ElementAt(1).TableRowProperties.GetFirstChild<GridBefore>().Val.Value);
        Assert.Equal(TableWidthUnitValues.Dxa, table.Table.GetFirstChild<TableProperties>().TableWidth.Type.Value);
    }

    [Fact]
    public void RowSetLeftIndent_GivesTheDifferenceToTheCellsOfTheRow()
    {
        TableWrapper table = Fill(3, 3);

        table.Rows[1].SetLeftIndent(30, RulerStyle.FirstColumn);

        Assert.Equal(new[] { 2408, 3009, 3009 }, table.Rows[1].Cells.Select(c => (int)Math.Round(c.Width * 20)).ToArray());
        Assert.Equal(30, table.Rows[1].LeftIndent);
    }

    [Fact]
    public void RowSetLeftIndent_ToZero_RemovesTheSpaceAgain()
    {
        TableWrapper table = Fill(3, 3);
        table.Rows[1].LeftIndent = 20;

        table.Rows[1].LeftIndent = 0;

        Assert.Equal(0, table.Rows[1].LeftIndent);
        Assert.Equal(3, table.Columns.Count);
    }

    [Fact]
    public void RowSetLeftIndent_RejectsWidthsOutsideTheRange()
    {
        TableWrapper table = Fill(2, 2);

        Assert.Throws<ArgumentOutOfRangeException>(() => table.Rows[0].SetLeftIndent(-1, RulerStyle.None));
    }

    [Fact]
    public void CellWidths_WithinATwipOrTwoOfAnEdge_DoNotMakeAColumnOfASliver()
    {
        TableWrapper table = Fill(2, 2);
        table.Cell(0, 0).Width = 225.4;

        table.Cell(1, 0).Width = 225.35;

        Assert.Equal(2, table.Columns.Count);
    }

    //A table whose first row has its first two cells merged the older way.
    private TableWrapper LegacyMergedTable()
    {
        TableWrapper table = Fill(2, 3, "a", null, "c", "d", "e", "f");
        TableCell first = table.Cell(0, 0).TableCell;
        TableCell second = table.Cell(0, 1).TableCell;
        first.TableCellProperties.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Restart });
        second.TableCellProperties.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Continue });
        return table;
    }

    [Fact]
    public void Merge_OfATableWithOlderHorizontalMerges_ConvertsThemToGridSpans()
    {
        TableWrapper table = LegacyMergedTable();

        table.Cell(1, 0).Merge(table.Cell(1, 1));

        Assert.Empty(table.Table.Descendants<HorizontalMerge>());
        Assert.Equal(2, table.Table.Elements<TableRow>().First().Elements<TableCell>().Count());
        Assert.Equal(2, table.Cell(0, 0).TableCell.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal(6017, table.Cell(0, 0).TableCell.TableCellProperties.TableCellWidth == null ? 0 : int.Parse(table.Cell(0, 0).TableCell.TableCellProperties.TableCellWidth.Width.Value, CultureInfo.InvariantCulture));
        Assert.Equal("a", CellText(table, 0, 0));
    }

    [Fact]
    public void ColumnDelete_AndSplit_WorkOnATableWithOlderHorizontalMerges()
    {
        TableWrapper table = LegacyMergedTable();
        table.Columns[2].Delete();
        Assert.Empty(table.Table.Descendants<HorizontalMerge>());
        Assert.Equal(2, table.Columns.Count);

        TableWrapper other = LegacyMergedTable();
        other.Cell(1, 2).Split(1, 2);
        Assert.Empty(other.Table.Descendants<HorizontalMerge>());
    }
}
