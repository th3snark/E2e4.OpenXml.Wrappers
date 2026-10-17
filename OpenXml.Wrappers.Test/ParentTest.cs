using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class ParentTest
{
    private static Body CreateBody()
    {
        Table table = new Table(
            new TableGrid(new GridColumn(), new GridColumn()),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("a")))), new TableCell(new Paragraph(new Run(new Text("b"))))));
        return new Body(
            new Paragraph(new BookmarkStart() { Id = "1", Name = "bm" }, new Run(new Text("first")), new BookmarkEnd() { Id = "1" }),
            table);
    }

    [Fact]
    public void ColumnAndColumnCollectionParentIsTheTable()
    {
        Table table = CreateBody().Elements<Table>().Single();
        TableWrapper wrapper = new TableWrapper(table);

        Assert.Equal(wrapper, new TableColumnWrapperCollection(new RangeWrapper(table)).Parent);
        Assert.Equal(wrapper, new TableColumnWrapper(table.GetFirstChild<TableGrid>().Elements<GridColumn>().Last()).Parent);
    }

    [Fact]
    public void ParagraphParentIsTheContainingBlock()
    {
        Body body = CreateBody();
        Table table = body.Elements<Table>().Single();

        Assert.Equal(new RangeWrapper(body), new ParagraphWrapper(body.Elements<Paragraph>().Single()).Parent);
        TableCell cell = table.Descendants<TableCell>().First();
        Assert.Equal(new RangeWrapper(cell), new ParagraphWrapper(cell.Elements<Paragraph>().Single()).Parent);
    }

    [Fact]
    public void BookmarkAndCollectionParents()
    {
        Body body = CreateBody();
        RangeWrapper range = new RangeWrapper(body);
        BookmarkWrapperCollection bookmarks = range.Bookmarks;

        Assert.Equal(range, bookmarks.Parent);
        Assert.Equal(new RangeWrapper(body.Elements<Paragraph>().First()), bookmarks[0].Parent);
        Assert.Equal(range, range.Comments.Parent);
    }
}
