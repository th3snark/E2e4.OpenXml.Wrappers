using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
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

    [Fact]
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

    [Fact]
    public void AddDoesNotDuplicateBookmarkOrCommentAnchors()
    {
        Table table = new Table(
            new TableGrid(new GridColumn() { Width = "1000" }),
            new TableRow(new BookmarkStart() { Id = "1", Name = "row" },
                new TableCell(new Paragraph(
                    new BookmarkStart() { Id = "2", Name = "cell" },
                    new CommentRangeStart() { Id = "0" },
                    new Run(new Text("a")),
                    new CommentRangeEnd() { Id = "0" },
                    new Run(new CommentReference() { Id = "0" }),
                    new BookmarkEnd() { Id = "2" })),
                new BookmarkEnd() { Id = "1" }));

        TableRowWrapper added = new TableWrapper(table).Rows.Add();

        Assert.Equal(2, table.Elements<TableRow>().Count());
        Assert.Equal(2, table.Descendants<BookmarkStart>().Count());
        Assert.Equal(2, table.Descendants<BookmarkEnd>().Count());
        Assert.Single(table.Descendants<CommentRangeStart>());
        Assert.Single(table.Descendants<CommentReference>());
        Assert.Equal("\r\a\r\a".Length, new RangeWrapper(added.TableRow).Text.Length);
    }

    [Fact]
    public void AddToATableWithoutRowsThrows()
    {
        Assert.Throws<InvalidOperationException>(() => new TableWrapper(new Table()).Rows.Add());
    }

    [Fact]
    public void AddToACollectionThatIsNotATableThrows()
    {
        Assert.Throws<InvalidOperationException>(() => new TableRowWrapperCollection(new RangeWrapper(new TableCell(new Paragraph()))).Add());
    }
}
