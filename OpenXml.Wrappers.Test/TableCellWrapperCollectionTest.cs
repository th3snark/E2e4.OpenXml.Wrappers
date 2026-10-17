using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class TableCellWrapperCollectionTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public TableCellWrapperCollectionTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void CountPropertyIsCorrect()
    {
        Assert.Equal(7, _document1.Tables[0].Rows[0].Cells.Count);
        Assert.Equal(8, _document1.Tables[0].Rows[1].Cells.Count);
        Assert.Equal(8, _document1.Tables[0].Rows[2].Cells.Count);
        Assert.Equal(8, _document1.Tables[0].Rows[3].Cells.Count);
        Assert.Equal(8, _document1.Tables[0].Rows[4].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Rows[0].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Rows[1].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[0].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[0].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[0].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[0].Rows[3].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[0].Rows[4].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[1].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[1].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[1].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[2].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[2].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[2].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[2].Rows[3].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Tables[2].Rows[4].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[1].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[1].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[1].Rows[2].Cells.Count);
        Assert.Equal(4, _document1.Tables[1].Tables[1].Tables[0].Rows[0].Cells.Count);
        Assert.Equal(4, _document1.Tables[1].Tables[1].Tables[0].Rows[1].Cells.Count);
        Assert.Equal(4, _document1.Tables[1].Tables[1].Tables[0].Rows[2].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[1].Rows[0].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[1].Rows[1].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[1].Rows[2].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[1].Rows[3].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[1].Rows[4].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[0].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[0].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[0].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[0].Rows[3].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[0].Rows[4].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[1].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[1].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[1].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[1].Rows[3].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[1].Rows[4].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[2].Rows[0].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[2].Rows[1].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[2].Rows[2].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[2].Rows[3].Cells.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Tables[2].Rows[4].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].Rows[0].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].Rows[1].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].Rows[2].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].Rows[3].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].Rows[4].Cells.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].Rows[0].Cells.Count);
    }
}
