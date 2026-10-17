using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class TableWrapperCollectionTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public TableWrapperCollectionTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    [Fact]
    public void TablesOfABookmarkRangeAreTheTablesItSpans()
    {
        Table inner = new Table(new TableRow(new TableCell(new Paragraph())));
        Table outer = new Table(new TableRow(new TableCell(inner, new Paragraph())));
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "a" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Body body = new Body(new Paragraph(start), outer, new Paragraph(end));

        TableWrapperCollection tables = new RangeWrapper(start, end).Tables;

        Assert.Equal(1, tables.Count);
        Assert.Same(outer, tables[0].Table);
    }

    [Fact]
    public void TablesOfABookmarkRangeWithNoTablesIsEmpty()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "a" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Body body = new Body(new Paragraph(start, new Run(new Text("a")), end));

        Assert.Equal(0, new RangeWrapper(start, end).Tables.Count);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void CountPropertyIsCorrect()
    {
        Assert.Equal(2, _document1.Tables.Count);
        Assert.Equal(0, _document1.Tables[0].Tables.Count);
        Assert.Equal(3, _document1.Tables[1].Tables.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[0].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[0].Tables[0].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[0].Tables[1].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[0].Tables[2].Tables.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[1].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[1].Tables[0].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[1].Tables[1].Tables.Count);
        Assert.Equal(4, _document1.Tables[1].Tables[2].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[2].Tables[0].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[2].Tables[1].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[2].Tables[2].Tables.Count);
        Assert.Equal(0, _document1.Tables[1].Tables[2].Tables[3].Tables.Count);
    }

    [Fact]
    public void NestingLevelPropertyIsCorrect()
    {
        Assert.Equal(2, _document1.Tables[0].Tables.NestingLevel);
        Assert.Equal(2, _document1.Tables[1].Tables.NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[0].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[0].Tables[0].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[0].Tables[1].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[0].Tables[2].Tables.NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[1].Tables[0].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[1].Tables[1].Tables.NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[2].Tables[0].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[2].Tables[1].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[2].Tables[2].Tables.NestingLevel);
        Assert.Equal(4, _document1.Tables[1].Tables[2].Tables[3].Tables.NestingLevel);
    }

    [Fact]
    public void _TestTablesRows()
    {
        int tableCount = 0;
        foreach (Table table in _document1.WordprocessingDocument.MainDocumentPart.Document.Descendants<Table>())
        {
            tableCount++;
            TableWrapper tableWrapper = new TableWrapper(table);
            int rowCount = 0;
            if (tableWrapper.Parent == null || tableWrapper.Tables.Parent == null)
            {
                Assert.Null(tableWrapper.Parent);
                Assert.Null(tableWrapper.Tables.Parent);
            }
            else
            {
                Assert.Same(tableWrapper.Parent.Table, tableWrapper.Tables.Parent.Table);
            }
            Assert.Same(tableWrapper.Table, tableWrapper.Rows.Parent.Table);

            foreach (TableRowWrapper row in tableWrapper.Rows)
            {
                rowCount++;
                Assert.Same(tableWrapper.Table, row.Parent.Table);
                foreach (TableCellWrapper cell in row.Cells)
                {
                    Assert.Same(row.TableRow, cell.Parent.TableRow);
                }
            }

            for (int rowIndex = 0; rowIndex < tableWrapper.Rows.Count; rowIndex++)
            {
                TableRowWrapper row = tableWrapper.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
                {
                    TableCellWrapper cell = row.Cells[columnIndex];
                    Assert.Same(row.TableRow, cell.Parent.TableRow);
                }
            }
        }
    }
}
