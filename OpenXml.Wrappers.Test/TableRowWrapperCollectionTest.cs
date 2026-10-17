using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Word;
using OpenXml.Wrappers.Test.Utils;

namespace OpenXml.Wrappers.Test;

public class TableRowWrapperCollectionTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public TableRowWrapperCollectionTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    public void CountPropertyIsCorrect()
    {
        Assert.Equal(5, _document1.Tables[0].Rows.Count);
        Assert.Equal(3, _document1.Tables[1].Rows.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[0].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[0].Tables[0].Rows.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[0].Tables[1].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[0].Tables[2].Rows.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Rows.Count);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[0].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[1].Tables[1].Rows.Count);
        Assert.Equal(2, _document1.Tables[1].Tables[2].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[2].Tables[0].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[2].Tables[1].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[2].Tables[2].Rows.Count);
        Assert.Equal(5, _document1.Tables[1].Tables[2].Tables[3].Rows.Count);
    }
}
