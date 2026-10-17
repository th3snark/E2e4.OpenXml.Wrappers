using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Word;
using OpenXml.Wrappers.Test.Utils;

namespace OpenXml.Wrappers.Test;

public class TableWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public TableWrapperTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void NestingLevelPropertyIsCorrect()
    {
        Assert.Equal(1, _document1.Tables[0].NestingLevel);
        Assert.Equal(1, _document1.Tables[1].NestingLevel);
        Assert.Equal(2, _document1.Tables[1].Tables[0].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[0].Tables[0].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[0].Tables[1].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[0].Tables[2].NestingLevel);
        Assert.Equal(2, _document1.Tables[1].Tables[1].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[0].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[1].Tables[1].NestingLevel);
        Assert.Equal(2, _document1.Tables[1].Tables[2].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[0].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[1].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[2].NestingLevel);
        Assert.Equal(3, _document1.Tables[1].Tables[2].Tables[3].NestingLevel);
    }

    [Fact]
    public void RowsPropertyIsNotNull()
    {
        Assert.NotNull(_document1.Tables[0].Rows);
        Assert.NotNull(_document1.Tables[1].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[0].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[0].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[1].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[0].Tables[2].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[1].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[0].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[1].Tables[1].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[2].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[0].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[1].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[2].Rows);
        Assert.NotNull(_document1.Tables[1].Tables[2].Tables[3].Rows);
    }
}
