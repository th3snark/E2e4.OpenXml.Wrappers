using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperInsertParagraphTest
{
    private static Body CreateBody()
    {
        return new Body(
            new Paragraph(new Run(new Text("first"))),
            new Table(new TableRow(new TableCell(new TableCellProperties(), new Paragraph(new Run(new Text("cell")))))),
            new SectionProperties());
    }

    [Fact]
    public void AfterTheDocumentAppendsAfterATrailingTableAndBeforeTheSectionProperties()
    {
        Body body = CreateBody();

        ParagraphWrapper paragraph = new RangeWrapper(body).InsertParagraphAfter();

        Assert.Same(paragraph.Paragraph, body.Elements().ElementAt(2));
        Assert.IsType<SectionProperties>(body.LastChild);
        Assert.Single(body.Descendants<TableCell>().Single().Elements<Paragraph>());
    }

    [Fact]
    public void BeforeTheDocumentPrepends()
    {
        Body body = CreateBody();

        ParagraphWrapper paragraph = new RangeWrapper(body).InsertParagraphBefore();

        Assert.Same(paragraph.Paragraph, body.FirstChild);
    }

    [Fact]
    public void AroundAParagraphAndATableAddsSiblings()
    {
        Body body = CreateBody();
        Paragraph first = body.Elements<Paragraph>().First();
        Table table = body.Elements<Table>().Single();

        ParagraphWrapper afterFirst = new RangeWrapper(first).InsertParagraphAfter();
        ParagraphWrapper beforeTable = new RangeWrapper(table).InsertParagraphBefore();

        Assert.Same(first.NextSibling(), afterFirst.Paragraph);
        Assert.Same(table.PreviousSibling(), beforeTable.Paragraph);
        Assert.Equal(2, body.Elements<Paragraph>().Count() - 1);
    }

    [Fact]
    public void InATableCellTheParagraphIsFirstOrLastAndAfterTheProperties()
    {
        TableCell cell = CreateBody().Descendants<TableCell>().Single();

        ParagraphWrapper last = new RangeWrapper(cell).InsertParagraphAfter();
        ParagraphWrapper first = new RangeWrapper(cell).InsertParagraphBefore();

        Assert.IsType<TableCellProperties>(cell.FirstChild);
        Assert.Same(first.Paragraph, cell.Elements<Paragraph>().First());
        Assert.Same(last.Paragraph, cell.Elements<Paragraph>().Last());
    }

    [Fact]
    public void AroundABookmarkRangeTheParagraphFollowsItsParagraph()
    {
        Paragraph paragraph = new Paragraph(new BookmarkStart() { Id = "1", Name = "b" }, new Run(new Text("x")), new BookmarkEnd() { Id = "1" });
        Body body = new Body(paragraph);
        RangeWrapper range = new RangeWrapper(body).Bookmarks["b"].Range;

        ParagraphWrapper after = range.InsertParagraphAfter();
        ParagraphWrapper before = range.InsertParagraphBefore();

        Assert.Same(paragraph.NextSibling(), after.Paragraph);
        Assert.Same(paragraph.PreviousSibling(), before.Paragraph);
    }

    [Fact]
    public void InAHeaderAppends()
    {
        Header header = new Header(new Paragraph());

        ParagraphWrapper added = new RangeWrapper(header).InsertParagraphAfter();

        Assert.Same(added.Paragraph, header.LastChild);
    }

    [Fact]
    public void ARowIsNotSupported()
    {
        RangeWrapper range = new RangeWrapper(CreateBody().Descendants<TableRow>().Single());

        Assert.Throws<NotSupportedException>(() => range.InsertParagraphAfter());
    }
}
