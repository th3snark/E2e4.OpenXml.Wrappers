using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperInsertTest
{
    //The text of the last paragraph of a table cell includes the end of paragraph and end of cell markers.
    private static readonly char[] CellEnd = { (char)13, (char)7 };

    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    [Fact]
    public void InsertBeforeAParagraphAddsTextAtTheStart()
    {
        Paragraph paragraph = CreateParagraph("b");
        new RangeWrapper(paragraph).InsertBefore("a");

        Assert.Equal("ab\r", new RangeWrapper(paragraph).Text);
    }

    [Fact]
    public void InsertBeforeAParagraphWithPropertiesKeepsThePropertiesFirst()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(), new Run(new Text("b")));
        new RangeWrapper(paragraph).InsertBefore("a");

        Assert.IsType<ParagraphProperties>(paragraph.FirstChild);
        Assert.Equal("ab\r", new RangeWrapper(paragraph).Text);
    }

    [Fact]
    public void InsertAfterAParagraphAddsTextBeforeTheParagraphMark()
    {
        Paragraph paragraph = CreateParagraph("a");
        new RangeWrapper(paragraph).InsertAfter("b");

        Assert.Equal("ab\r", new RangeWrapper(paragraph).Text);
    }

    [Fact]
    public void InsertingTextKeepsTheOrderOfSeveralRuns()
    {
        Paragraph paragraph = CreateParagraph("c");
        RangeWrapper range = new RangeWrapper(paragraph);
        range.InsertBefore("a\tb");
        range.InsertAfter("d\ne");

        Assert.Equal("a\tbcd\ve\r", range.Text);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void InsertingNothingDoesNothing(string text)
    {
        Paragraph paragraph = CreateParagraph("a");
        RangeWrapper range = new RangeWrapper(paragraph);
        range.InsertBefore(text);
        range.InsertAfter(text);

        Assert.Single(paragraph.Elements<Run>());
    }

    [Fact]
    public void InsertedRunsTakeTheFormattingOfTheParagraphMark()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new ParagraphMarkRunProperties(new Bold())), CreateParagraph("x").FirstChild.CloneNode(true));
        new RangeWrapper(paragraph).InsertAfter("y");

        Run inserted = paragraph.Elements<Run>().Last();
        Assert.Single(inserted.RunProperties.Elements<Bold>());
    }

    [Fact]
    public void InsertIntoATableCellUsesTheFirstAndLastParagraphs()
    {
        TableCell cell = new TableCell(CreateParagraph("1"), CreateParagraph("2"));
        RangeWrapper range = new RangeWrapper(cell);
        range.InsertBefore("a");
        range.InsertAfter("z");

        Paragraph[] paragraphs = cell.Elements<Paragraph>().ToArray();
        Assert.Equal("a1", new ParagraphWrapper(paragraphs[0]).Text);
        Assert.Equal("2z", new ParagraphWrapper(paragraphs[1]).Text.TrimEnd(CellEnd));
    }

    [Fact]
    public void InsertIntoAnEmptyTableCellAddsAParagraph()
    {
        TableCell cell = new TableCell();
        new RangeWrapper(cell).InsertAfter("a");

        Assert.Equal("a", new ParagraphWrapper(cell.Elements<Paragraph>().Single()).Text.TrimEnd(CellEnd));
    }

    [Fact]
    public void InsertIntoABookmarkRangeStaysInsideTheBookmark()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph paragraph = new Paragraph(new Run(new Text("x")), start, new Run(new Text("b")), end, new Run(new Text("y")));
        Body body = new Body(paragraph);
        RangeWrapper range = new RangeWrapper(start, end);

        range.InsertBefore("[");
        range.InsertAfter("]");

        Assert.Equal("[b]", range.Text);
        Assert.Equal("x[b]y", new ParagraphWrapper(paragraph).Text);
    }

    [Fact]
    public void InsertIntoAnEmptyBookmarkRangeAddsTheText()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph paragraph = new Paragraph(start, end);
        Body body = new Body(paragraph);
        RangeWrapper range = new RangeWrapper(start, end);

        range.InsertAfter("b");
        range.InsertBefore("a");

        Assert.Equal("ab", range.Text);
    }

    [Fact]
    public void InsertIntoACommentRangeStaysInsideTheRange()
    {
        CommentRangeStart start = new CommentRangeStart() { Id = "1" };
        CommentRangeEnd end = new CommentRangeEnd() { Id = "1" };
        Paragraph paragraph = new Paragraph(start, new Run(new Text("b")), end);
        Body body = new Body(paragraph);
        RangeWrapper range = new RangeWrapper(start, end);

        range.InsertBefore("a");
        range.InsertAfter("c");

        Assert.Equal("abc", range.Text.TrimEnd('\r'));
    }

    [Fact]
    public void InsertIntoARangeWhoseMarkersAreNotInAParagraphIsNotSupported()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Body body = new Body(start, CreateParagraph("a"), end);

        Assert.Throws<NotSupportedException>(() => new RangeWrapper(start, end).InsertBefore("x"));
        Assert.Throws<NotSupportedException>(() => new RangeWrapper(start, end).InsertAfter("x"));
    }

    [Fact]
    public void InsertAfterTheDocumentAppendsToTheLastParagraph()
    {
        Paragraph last = CreateParagraph("b");
        Document document = new Document(new Body(CreateParagraph("a"), last, new SectionProperties()));
        new RangeWrapper(document).InsertAfter("!");

        Assert.Equal("b!", new ParagraphWrapper(last).Text);
    }

    [Fact]
    public void InsertBeforeTheDocumentPrependsToTheFirstParagraph()
    {
        Paragraph first = CreateParagraph("a");
        Body body = new Body(first, CreateParagraph("b"));
        new RangeWrapper(body).InsertBefore(">");

        Assert.Equal(">a", new ParagraphWrapper(first).Text);
    }

    [Fact]
    public void InsertAfterADocumentEndingInATableAddsAParagraphAfterTheTable()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        SectionProperties sectionProperties = new SectionProperties();
        Document document = new Document(new Body(table, sectionProperties));
        new RangeWrapper(document).InsertAfter("end");

        Paragraph added = (Paragraph)table.NextSibling();
        Assert.Equal("end", new ParagraphWrapper(added).Text);
        Assert.Same(sectionProperties, added.NextSibling());
    }

    [Fact]
    public void InsertBeforeADocumentStartingWithATableAddsAParagraphBeforeTheTable()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        Body body = new Body(table, new SectionProperties());
        new RangeWrapper(body).InsertBefore("start");

        Assert.Equal("start", new ParagraphWrapper((Paragraph)body.FirstChild).Text);
        Assert.Same(table, body.FirstChild.NextSibling());
    }

    [Fact]
    public void InsertIntoAnEmptyDocumentAddsAParagraphBeforeTheSectionProperties()
    {
        SectionProperties sectionProperties = new SectionProperties();
        Body body = new Body(sectionProperties);
        new RangeWrapper(body).InsertAfter("text");

        Assert.Equal("text", new ParagraphWrapper((Paragraph)body.FirstChild).Text);
        Assert.Same(sectionProperties, body.LastChild);
    }

    [Fact]
    public void InsertIntoADocumentWithoutABodyThrows()
    {
        Assert.Throws<InvalidOperationException>(() => new RangeWrapper(new Document()).InsertAfter("a"));
    }

    [Fact]
    public void InsertIntoATableOrRowIsNotSupported()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));

        Assert.Throws<NotSupportedException>(() => new RangeWrapper(table).InsertAfter("a"));
        Assert.Throws<NotSupportedException>(() => new RangeWrapper(table.Elements<TableRow>().First()).InsertBefore("a"));
    }
}
