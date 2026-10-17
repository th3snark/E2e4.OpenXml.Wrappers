using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class BookmarkTableTest
{
    [Fact]
    public void ATableCanBeBookmarkedWithMarkersAroundIt()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph(new Run(new Text("cell"))))));
        Body body = new Body(new Paragraph(new Run(new Text("before"))), table, new Paragraph(new Run(new Text("after"))));
        RangeWrapper range = new RangeWrapper(body);

        BookmarkWrapper bookmark = range.Bookmarks.Add("the_table", new TableWrapper(table).Range);

        Assert.IsType<BookmarkStart>(table.PreviousSibling());
        Assert.IsType<BookmarkEnd>(table.NextSibling());
        Assert.Equal("the_table", bookmark.Name);
        Assert.Single(range.Bookmarks);

        bookmark.Delete();
        Assert.Empty(body.Descendants<BookmarkStart>());
        Assert.Empty(body.Descendants<BookmarkEnd>());
    }
}
