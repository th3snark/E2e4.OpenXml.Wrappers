using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class TableCellWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;
    private WordprocessingDocumentWrapper _document2;

    public TableCellWrapperTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
        _document2 = WordprocessingDocumentWrapper.Open(TestFile.Document2);
    }

    public void Dispose()
    {
        _document1.Close();
        _document2.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void RangeTextPropertyIsCorrect()
    {
        Assert.Equal("R1C1\r\a", _document1.Tables[0].Rows[0].Cells[0].Range.Text);
        Assert.Equal("R1C2\r\a", _document1.Tables[0].Rows[0].Cells[1].Range.Text);
        Assert.Equal("R1C3\rR1C4\r\a", _document1.Tables[0].Rows[0].Cells[2].Range.Text);
        Assert.Equal("R5SomeBoldTextC6\r\a", _document1.Tables[0].Rows[4].Cells[2].Range.Text);

        Assert.Equal("R1C1\r\aR1C2\r\aR1C3\rR1C4\r\aR1C5\r\aR1C6\r\aR1C7\r\aR1C8\r\a\r\a",
           _document1.Tables[0].Rows[0].Range.Text);
        Assert.Equal("R2C1\r\aR2C2\r\aR2C3\r\aR2C4\r\aR2C5\r\aR2C6\r\aR2C7\r\aR2C8\r\a\r\a",
           _document1.Tables[0].Rows[1].Range.Text);
        Assert.Equal("R3C1\r\aR3C2\r\aR3C3\r\aR3C4\r\aR3C5\r\aR3C6\r\aR3C7\r\aR3C8\r\a\r\a",
           _document1.Tables[0].Rows[2].Range.Text);
        Assert.Equal("R4C1\r\aR4C2\r\aR4C3\r\aR4C4\r\aR4C5\r\aR4C6\r\aR4C7\r\aR4C8\r\a\r\a",
           _document1.Tables[0].Rows[3].Range.Text);
        Assert.Equal("R5C1\r\aR5C2\r\aR5SomeBoldTextC6\r\aR5C4\r\aR5C5\r\aR5C6\r\aR5C7\r\aR5C8\r\a\r\a",
           _document1.Tables[0].Rows[4].Range.Text);

        Assert.Equal(
           "R1C1\r\aR1C2\r\aR1C3\rR1C4\r\aR1C5\r\aR1C6\r\aR1C7\r\aR1C8\r\a\r\a" +
           "R2C1\r\aR2C2\r\aR2C3\r\aR2C4\r\aR2C5\r\aR2C6\r\aR2C7\r\aR2C8\r\a\r\a" +
           "R3C1\r\aR3C2\r\aR3C3\r\aR3C4\r\aR3C5\r\aR3C6\r\aR3C7\r\aR3C8\r\a\r\a" +
           "R4C1\r\aR4C2\r\aR4C3\r\aR4C4\r\aR4C5\r\aR4C6\r\aR4C7\r\aR4C8\r\a\r\a" +
           "R5C1\r\aR5C2\r\aR5SomeBoldTextC6\r\aR5C4\r\aR5C5\r\aR5C6\r\aR5C7\r\aR5C8\r\a\r\a",
           _document1.Tables[0].Range.Text);
    }

    [Fact]
    public void RowIndexPropertyIsCorrect()
    {
        foreach (TableWrapper table in _document1.Tables)
        {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                TableRowWrapper row = table.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
                {
                    TableCellWrapper cell = row.Cells[columnIndex];
                    Assert.Equal(rowIndex, cell.RowIndex);
                }
            }
        }
    }

    [Fact]
    public void ColumnIndexPropertyIsCorrect()
    {
        foreach (TableWrapper table in _document1.Tables)
        {
            foreach (TableRowWrapper row in table.Rows)
            {
                for (int index = 0; index < row.Cells.Count; index++)
                {
                    Assert.Equal(index, row.Cells[index].ColumnIndex);
                }
            }
        }
    }

    [Fact]
    public void foo()
    {
        Assert.Equal(1, _document2.Bookmarks["StudyNumber"].Range.Cells.Count);
        Assert.Equal(1, _document2.Bookmarks["StudyNumber"].Range.Cells[0].RowIndex);
        Assert.Equal(2, _document2.Bookmarks["StudyNumber"].Range.Cells[0].ColumnIndex);
    }
}
