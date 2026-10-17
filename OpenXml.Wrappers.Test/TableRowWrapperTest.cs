using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Word;
using OpenXml.Wrappers.Test.Utils;

namespace OpenXml.Wrappers.Test;

public class TableRowWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public TableRowWrapperTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void RowWrappersOfTheSameRowAreEqual()
    {
        Assert.True(_document1.Tables[0].Rows[0] == _document1.Tables[0].Rows[0]);
        Assert.True(_document1.Tables[0].Rows[0] != _document1.Tables[0].Rows[1]);
        Assert.True(_document1.Tables[0].Rows[1].Cells[0].Parent == _document1.Tables[0].Rows[1]);
    }

    [Fact]
    public void CellsPropertyIsNotNull()
    {
        Assert.NotNull(_document1.Tables[0].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[0].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[0].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[0].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[0].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[0].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[0].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[0].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[0].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[0].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[1].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[1].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[1].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[2].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[2].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[2].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[2].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[2].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[0].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[0].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[0].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[1].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[1].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[1].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[1].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[1].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[0].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[0].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[0].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[0].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[0].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[1].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[1].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[1].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[1].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[1].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[2].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[2].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[2].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[2].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[2].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows[0].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows[1].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows[2].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows[3].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows[4].Cells);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows[0].Cells);
    }

    [Fact]
    public void IndexPropertyIsCorrect()
    {
        foreach (TableWrapper table in _document1.Tables)
        {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                Assert.Equal(rowIndex, table.Rows[rowIndex].Index);
            }
        }
    }
}
