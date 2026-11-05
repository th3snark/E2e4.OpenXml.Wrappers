using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>ConvertToText</c>, <c>ConvertToTable</c> and <c>Sort</c>. The markup is what Word 16 wrote (the table-convert-to-text-*, row-convert-to-text,
/// range-convert-to-table-* and table-sort-* scenarios of MarkupConformanceTest).
/// </summary>
public class TableConvertSortTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableConvertSortTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
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

    private TableWrapper Fill(int rows, int columns, params string[] values)
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), rows, columns);
        int index = 0;
        foreach (TableCellWrapper cell in table.Range.Cells)
        {
            cell.Range.Text = values[index++];
        }
        return table;
    }

    private string[] Paragraphs()
    {
        return Body.Elements<Paragraph>().Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text))).ToArray();
    }

    private string[] Column(TableWrapper table, int column)
    {
        return table.Table.Elements<TableRow>().Select(r => string.Concat(r.Elements<TableCell>().ElementAt(column).Descendants<Text>().Select(t => t.Text))).ToArray();
    }

    [Fact]
    public void ConvertToText_WithTabs_MakesAParagraphOfEachRow_WithTabStopsAtTheCells()
    {
        TableWrapper table = Fill(2, 3, "a", "b", "c", "d", "e", "f");

        RangeWrapper result = table.ConvertToText(TableFieldSeparator.Tabs);

        Assert.Empty(Body.Descendants<Table>());
        Assert.Equal(new[] { "abc", "def", string.Empty }, Paragraphs());
        Paragraph first = Body.Elements<Paragraph>().First();
        Assert.Equal(2, first.ParagraphProperties.GetFirstChild<Tabs>().Elements<TabStop>().Count());
        Assert.Equal(2, first.Descendants<TabChar>().Count());
        Assert.Contains("a", result.Text);
    }

    [Fact]
    public void ConvertToText_WithParagraphs_MakesAParagraphOfEachCellParagraph()
    {
        TableWrapper table = Fill(2, 2, "a", "b", "c", "d");

        table.ConvertToText(TableFieldSeparator.Paragraphs);

        Assert.Equal(new[] { "a", "b", "c", "d", string.Empty }, Paragraphs());
    }

    [Theory]
    [InlineData(TableFieldSeparator.Commas, "a, b")]
    [InlineData(TableFieldSeparator.Default, "a-b")]
    public void ConvertToText_UsesTheSeparatorWordUses(TableFieldSeparator separator, string expected)
    {
        TableWrapper table = Fill(1, 2, "a", "b");

        table.ConvertToText(separator);

        Assert.Equal(expected, Paragraphs()[0]);
    }

    [Fact]
    public void ConvertToText_JoinsTheParagraphsOfACellWithASpace()
    {
        TableWrapper table = Fill(1, 2, "a", "b");
        table.Cell(0, 0).Range.InsertParagraphAfter();
        table.Cell(0, 0).Range.InsertAfter("more");

        table.ConvertToText(";");

        string[] fields = Paragraphs()[0].Split(';');
        Assert.Equal(2, fields.Length);
        Assert.Contains("more", fields[0]);
        Assert.Equal("b", fields[1]);
    }

    [Fact]
    public void ConvertToText_RejectsAnEmptySeparator_AndNestedTables()
    {
        TableWrapper table = Fill(1, 2, "a", "b");
        Assert.Throws<ArgumentException>(() => table.ConvertToText(string.Empty));

        _document.Tables.Add(table.Cell(0, 0).Range, 1, 1);
        Assert.Throws<NotSupportedException>(() => table.ConvertToText(TableFieldSeparator.Tabs));
    }

    [Fact]
    public void RowsConvertToText_SplitsTheTableAroundTheRows()
    {
        TableWrapper table = Fill(3, 1, "a", "b", "c");

        table.Rows[1].ConvertToText(TableFieldSeparator.Tabs);

        Assert.Equal(2, Body.Elements<Table>().Count());
        Assert.Equal(new[] { "b", string.Empty }, Paragraphs());
    }

    [Fact]
    public void RowsConvertToText_OfTheFirstRows_LeavesTheRestOfTheTable()
    {
        TableWrapper table = Fill(3, 1, "a", "b", "c");

        table.Rows.ConvertToText();

        Assert.Empty(Body.Descendants<Table>());
        Assert.Equal(new[] { "a", "b", "c", string.Empty }, Paragraphs());
    }

    [Fact]
    public void ConvertToTable_MakesARowOfEachParagraph_AndPadsShortOnes()
    {
        _document.Content.InsertAfter("a\tb\tc\rd\te");

        TableWrapper table = _document.Content.ConvertToTable(TableFieldSeparator.Tabs);

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal(new[] { "c", string.Empty }, Column(table, 2));
    }

    [Fact]
    public void ConvertToTable_WithParagraphs_FillsRowsOfNumColumns()
    {
        _document.Content.InsertAfter("a\rb\rc\rd\re");

        TableWrapper table = _document.Content.ConvertToTable(TableFieldSeparator.Paragraphs, null, 2);

        Assert.Equal(3, table.Rows.Count);
        Assert.Equal(new[] { "a", "c", "e" }, Column(table, 0));
        Assert.Equal(new[] { "b", "d", string.Empty }, Column(table, 1));
    }

    [Fact]
    public void ConvertToTable_BySeparatorText_AndMoreRowsThanTheTextNeeds()
    {
        _document.Content.InsertAfter("a;b\rc;d");

        TableWrapper table = _document.Content.ConvertToTable(";", 4);

        Assert.Equal(4, table.Rows.Count);
        Assert.Equal(new[] { "a", "c", string.Empty, string.Empty }, Column(table, 0));
    }

    [Fact]
    public void ConvertToTable_Rejects_BadArguments_AndAnEmptyRange()
    {
        Assert.Throws<ArgumentException>(() => _document.Content.ConvertToTable(string.Empty));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Content.ConvertToTable(TableFieldSeparator.Tabs, 0));
        Assert.Throws<InvalidOperationException>(() => _document.Content.ConvertToTable(TableFieldSeparator.Tabs));
    }

    [Fact]
    public void Sort_OrdersTheRowsByAColumn_WithoutRegardToCase_AndKeepsEqualRowsInOrder()
    {
        TableWrapper table = Fill(4, 2, "b", "1", "A", "2", "a", "3", "C", "4");

        table.Sort();

        Assert.Equal(new[] { "A", "a", "b", "C" }, Column(table, 0));
        Assert.Equal(new[] { "2", "3", "1", "4" }, Column(table, 1));
    }

    [Fact]
    public void Sort_Descending_AndByAnotherColumn()
    {
        TableWrapper table = Fill(3, 2, "a", "2", "b", "10", "c", "1");

        table.Sort(false, 1, SortFieldType.Numeric, SortOrder.Descending);

        Assert.Equal(new[] { "b", "a", "c" }, Column(table, 0));
    }

    [Fact]
    public void Sort_CanLeaveTheHeaderRowWhereItIs()
    {
        TableWrapper table = Fill(3, 1, "z", "b", "a");

        table.Sort(true);

        Assert.Equal(new[] { "z", "a", "b" }, Column(table, 0));
    }

    [Fact]
    public void Sort_ByDate_PutsTextThatIsNoDateFirst()
    {
        TableWrapper table = Fill(3, 1, "2025-03-01", "none", "2024-12-31");

        table.Sort(false, 0, SortFieldType.Date);

        Assert.Equal(new[] { "none", "2024-12-31", "2025-03-01" }, Column(table, 0));
    }

    [Fact]
    public void Sort_Rejects_ABadColumn_AndMergedCells()
    {
        TableWrapper table = Fill(2, 2, "a", "b", "c", "d");
        Assert.Throws<ArgumentOutOfRangeException>(() => table.Sort(false, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.Sort(false, -1));

        table.Cell(0, 0).Merge(table.Cell(1, 0));
        Assert.Throws<NotSupportedException>(() => table.Sort());
    }
}
