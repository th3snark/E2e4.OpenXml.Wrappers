using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperDeleteTest
{
    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    [Fact]
    public void DeletingAParagraphRemovesIt()
    {
        Paragraph middle = CreateParagraph("b");
        Body body = new Body(CreateParagraph("a"), middle, CreateParagraph("c"));
        new RangeWrapper(middle).Delete();

        Assert.Equal(2, body.Elements<Paragraph>().Count());
        Assert.Null(middle.Parent);
    }

    [Fact]
    public void DeletingTheOnlyParagraphOfTheBodyEmptiesItAndKeepsItsProperties()
    {
        Paragraph only = new Paragraph(new ParagraphProperties(new KeepNext()), new Run(new Text("a")));
        Body body = new Body(only, new SectionProperties());
        new RangeWrapper(only).Delete();

        Assert.Same(only, body.Elements<Paragraph>().Single());
        Assert.Empty(only.Elements<Run>());
        Assert.NotNull(only.ParagraphProperties.KeepNext);
    }

    [Fact]
    public void DeletingTheLastParagraphAfterATableEmptiesItSoTheTableStaysFollowedByAParagraph()
    {
        Paragraph last = CreateParagraph("a");
        Body body = new Body(new Table(new TableRow(new TableCell(new Paragraph()))), last);
        new RangeWrapper(last).Delete();

        Assert.Same(last, body.LastChild);
        Assert.Empty(last.Elements<Run>());
    }

    [Fact]
    public void DeletingTheOnlyParagraphOfATableCellEmptiesIt()
    {
        Paragraph only = CreateParagraph("a");
        TableCell cell = new TableCell(only);
        new RangeWrapper(only).Delete();

        Assert.Same(only, cell.Elements<Paragraph>().Single());
        Assert.Empty(only.Elements<Run>());
    }

    [Fact]
    public void DeletingAParagraphOfATableCellWithSeveralParagraphsRemovesIt()
    {
        Paragraph first = CreateParagraph("a");
        TableCell cell = new TableCell(first, CreateParagraph("b"));
        new RangeWrapper(first).Delete();

        Assert.Single(cell.Elements<Paragraph>());
    }

    [Fact]
    public void DeletingATableCellEmptiesItButKeepsItsPropertiesAndParagraphFormatting()
    {
        TableCellProperties properties = new TableCellProperties(new TableCellWidth() { Width = "100", Type = TableWidthUnitValues.Dxa });
        TableCell cell = new TableCell(properties, new Paragraph(new ParagraphProperties(new KeepNext()), new Run(new Text("a"))), CreateParagraph("b"));
        new RangeWrapper(cell).Delete();

        Assert.Same(properties, cell.FirstChild);
        Paragraph paragraph = cell.Elements<Paragraph>().Single();
        Assert.Empty(paragraph.Elements<Run>());
        Assert.NotNull(paragraph.ParagraphProperties.KeepNext);
    }

    [Fact]
    public void DeletingATableRowRemovesIt()
    {
        TableRow second = new TableRow(new TableCell(new Paragraph()));
        Table table = new Table(new TableRow(new TableCell(new Paragraph())), second);
        new RangeWrapper(second).Delete();

        Assert.Single(table.Elements<TableRow>());
    }

    [Fact]
    public void DeletingTheLastRowOfATableRemovesTheTable()
    {
        TableRow only = new TableRow(new TableCell(new Paragraph()));
        Table table = new Table(only);
        Body body = new Body(table, new Paragraph());
        new RangeWrapper(only).Delete();

        Assert.Empty(body.Elements<Table>());
    }

    [Fact]
    public void DeletingATableRemovesIt()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        Body body = new Body(table, new Paragraph());
        new RangeWrapper(table).Delete();

        Assert.Empty(body.Elements<Table>());
    }

    [Fact]
    public void DeletingABookmarkRangeRemovesItsTextButKeepsTheBookmark()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph paragraph = new Paragraph(new Run(new Text("x")), start, new Run(new Text("b")), end, new Run(new Text("y")));
        Body body = new Body(paragraph);
        new RangeWrapper(start, end).Delete();

        Assert.Equal("xy\r", new ParagraphWrapper(paragraph).Text);
        Assert.Same(start, paragraph.Descendants<BookmarkStart>().Single());
        Assert.Equal(String.Empty, new RangeWrapper(start, end).Text);
    }

    [Fact]
    public void DeletingAParagraphRemovesTheBookmarkWhoseOtherHalfIsOutsideIt()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph first = new Paragraph(new Run(new Text("a")), start);
        Paragraph second = new Paragraph(new Run(new Text("b")), end);
        Body body = new Body(first, second, CreateParagraph("c"));
        new RangeWrapper(first).Delete();

        Assert.Empty(body.Descendants<BookmarkStart>());
        Assert.Empty(body.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void DeletingAParagraphKeepsBookmarksThatAreEntirelyOutsideIt()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph first = CreateParagraph("a");
        Body body = new Body(first, new Paragraph(start, new Run(new Text("b")), end));
        new RangeWrapper(first).Delete();

        Assert.Single(body.Descendants<BookmarkStart>());
        Assert.Single(body.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void DeletingContentRemovesTheCommentAnchoredInIt()
    {
        string path = System.IO.Path.GetTempFileName();
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, true))
            {
                Body body = document.WordprocessingDocument.MainDocumentPart.Document.Body;
                Paragraph keep = CreateParagraph("keep");
                Paragraph doomed = CreateParagraph("doomed");
                body.InsertAt(keep, 0);
                body.InsertAt(doomed, 1);
                document.Comments.Add(new ParagraphWrapper(keep).Range, "stays");
                document.Comments.Add(new ParagraphWrapper(doomed).Range, "goes");
                Assert.Equal(2, document.Comments.Count);

                new RangeWrapper(doomed).Delete();

                Assert.Equal(1, document.Comments.Count);
                Comments comments = document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments;
                Assert.Single(comments.Elements<Comment>());
                Assert.Single(body.Descendants<CommentReference>());
                Assert.Equal("stays", document.Comments[0].Text.TrimEnd('\r'));
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void DeletingTheDocumentLeavesOneEmptyParagraphBeforeTheSectionProperties()
    {
        SectionProperties sectionProperties = new SectionProperties();
        Body body = new Body(CreateParagraph("a"), new Table(new TableRow(new TableCell(new Paragraph()))), CreateParagraph("b"), sectionProperties);
        new RangeWrapper(new Document(body)).Delete();

        Assert.Equal(2, body.ChildElements.Count);
        Assert.IsType<Paragraph>(body.FirstChild);
        Assert.Empty(body.FirstChild.Elements<Run>());
        Assert.Same(sectionProperties, body.LastChild);
    }

    [Fact]
    public void DeletingACommentTextRangeIsNotSupported()
    {
        Assert.Throws<NotSupportedException>(() => new RangeWrapper(new Comment()).Delete());
    }
}
