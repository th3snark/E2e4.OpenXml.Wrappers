using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
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
    public void ParentIsTheDirectlyEnclosingTable()
    {
        Table inner = new Table(new TableRow(new TableCell(new Paragraph())));
        Table middle = new Table(new TableRow(new TableCell(inner, new Paragraph())));
        Table outer = new Table(new TableRow(new TableCell(middle, new Paragraph())));

        Assert.Null(new TableWrapper(outer).Parent);
        Assert.Same(outer, new TableWrapper(middle).Parent.Table);
        Assert.Same(middle, new TableWrapper(inner).Parent.Table);
    }

    [Fact]
    public void ParentOfTheTablesCollectionOfATableIsTheEnclosingTable()
    {
        Table inner = new Table(new TableRow(new TableCell(new Paragraph())));
        Table outer = new Table(new TableRow(new TableCell(inner, new Paragraph())));

        Assert.Null(new TableWrapper(outer).Tables.Parent);
        Assert.Same(outer, new TableWrapper(inner).Range.Tables.Parent.Table);
        Assert.Same(outer, new TableWrapperCollection(new RangeWrapper(outer.Descendants<TableCell>().First())).Parent.Table);
    }

    [Fact]
    public void WidthOfATableWithoutPropertiesIsZero()
    {
        Assert.Equal(0, new TableWrapper(new Table()).Width);
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
