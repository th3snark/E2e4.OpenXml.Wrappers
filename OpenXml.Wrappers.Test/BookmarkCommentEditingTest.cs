using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class BookmarkCommentEditingTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly Body _body;

    public BookmarkCommentEditingTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, true);
        _body = _document.WordprocessingDocument.MainDocumentPart.Document.Body;
        _body.RemoveAllChildren<Paragraph>();
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private Paragraph AddParagraph(string text)
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new KeepNext()), new Run(new Text(text)));
        _body.AppendChild(paragraph);
        return paragraph;
    }

    [Fact]
    public void AddingABookmarkToAParagraphWrapsItsContent()
    {
        Paragraph paragraph = AddParagraph("hello");
        BookmarkWrapper bookmark = _document.Bookmarks.Add("Greeting", new RangeWrapper(paragraph));

        Assert.Equal("Greeting", bookmark.Name);
        Assert.Equal(new[] { "pPr", "bookmarkStart", "r", "bookmarkEnd" }, paragraph.ChildElements.Select(e => e.LocalName).ToArray());
        Assert.Equal("hello", bookmark.Range.Text);
        Assert.Same(paragraph.Elements<BookmarkStart>().Single(), _document.Bookmarks["Greeting"].Range.StartElement);
    }

    [Fact]
    public void AddingABookmarkToATableCellSpansItsParagraphs()
    {
        TableCell cell = new TableCell(new Paragraph(new Run(new Text("a"))), new Paragraph(new Run(new Text("b"))));
        _body.AppendChild(new Table(new TableRow(cell)));

        _document.Bookmarks.Add("Cell", new RangeWrapper(cell));

        Assert.Single(cell.Elements<Paragraph>().First().Elements<BookmarkStart>());
        Assert.Single(cell.Elements<Paragraph>().Last().Elements<BookmarkEnd>());
    }

    [Fact]
    public void AddingABookmarkToABookmarkRangeUsesTheSameMarkers()
    {
        Paragraph paragraph = AddParagraph("x");
        BookmarkWrapper first = _document.Bookmarks.Add("First", new RangeWrapper(paragraph));

        BookmarkWrapper second = _document.Bookmarks.Add("Second", first.Range);

        Assert.NotEqual(first.Id, second.Id);
        Assert.Equal(2, paragraph.Elements<BookmarkStart>().Count());
        Assert.Equal(2, paragraph.Elements<BookmarkEnd>().Count());
    }

    [Fact]
    public void AddingAnExistingNameMovesTheBookmark()
    {
        Paragraph one = AddParagraph("one");
        Paragraph two = AddParagraph("two");
        _document.Bookmarks.Add("Mark", new RangeWrapper(one));

        _document.Bookmarks.Add("mark", new RangeWrapper(two));

        Assert.Empty(one.Elements<BookmarkStart>());
        Assert.Empty(one.Elements<BookmarkEnd>());
        Assert.Single(two.Elements<BookmarkStart>());
        Assert.Single(_body.Descendants<BookmarkStart>());
    }

    [Fact]
    public void AddingABookmarkToItsOwnRangeIsANoOp()
    {
        Paragraph paragraph = AddParagraph("x");
        BookmarkWrapper bookmark = _document.Bookmarks.Add("Same", new RangeWrapper(paragraph));

        Assert.Equal(bookmark, _document.Bookmarks.Add("same", bookmark.Range));
        Assert.Single(paragraph.Elements<BookmarkStart>());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("1abc")]
    [InlineData("has space")]
    public void InvalidBookmarkNamesAreRejected(string name)
    {
        Paragraph paragraph = AddParagraph("x");

        Assert.Throws<ArgumentException>(() => _document.Bookmarks.Add(name, new RangeWrapper(paragraph)));
        Assert.Empty(paragraph.Elements<BookmarkStart>());
    }

    [Fact]
    public void ANameOverFortyCharactersIsRejected()
    {
        Paragraph paragraph = AddParagraph("x");

        Assert.Throws<ArgumentException>(() => _document.Bookmarks.Add(new string('a', 41), new RangeWrapper(paragraph)));
        _document.Bookmarks.Add(new string('a', 40), new RangeWrapper(paragraph));
    }

    [Fact]
    public void ABookmarkCannotBeAddedToATableRow()
    {
        Table table = new Table(new TableRow(new TableCell(new Paragraph())));
        _body.AppendChild(table);

        Assert.Throws<NotSupportedException>(() => _document.Bookmarks.Add("T", new RangeWrapper(table.Elements<TableRow>().Single())));
    }

    [Fact]
    public void ExistsAndLookupIgnoreCase()
    {
        _document.Bookmarks.Add("Alpha", new RangeWrapper(AddParagraph("x")));

        Assert.True(_document.Bookmarks.Exists("ALPHA"));
        Assert.False(_document.Bookmarks.Exists("Beta"));
        Assert.False(_document.Bookmarks.Exists(null));
        Assert.NotNull(_document.Bookmarks["alpha"]);
    }

    [Fact]
    public void ABookmarkSurvivesASaveAndReload()
    {
        _document.Bookmarks.Add("Kept", new RangeWrapper(AddParagraph("text")));
        string copy = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            _document.SaveAs(copy).Close();
            using WordprocessingDocumentWrapper reopened = WordprocessingDocumentWrapper.Open(copy, false);
            Assert.NotNull(reopened.Bookmarks["Kept"]);
        }
        finally
        {
            File.Delete(copy);
        }
    }

    [Fact]
    public void CommentTextCanBeReplaced()
    {
        CommentWrapper comment = _document.Comments.Add(new RangeWrapper(AddParagraph("body")), "old");

        comment.Text = "new";
        Assert.Equal("new\r", comment.Text.TrimEnd('\a'));

        comment.Text = "one\rtwo\r";
        Comment element = _document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>().Single();
        Assert.Equal(2, element.Elements<Paragraph>().Count());
        Assert.Single(element.Descendants<AnnotationReferenceMark>());
        Assert.Equal("onetwo", element.InnerText);

        comment.Text = null;
        Assert.Single(element.Elements<Paragraph>());
        Assert.Equal(string.Empty, element.InnerText);
    }

    [Fact]
    public void CommentDateRoundTrips()
    {
        CommentWrapper comment = _document.Comments.Add(new RangeWrapper(AddParagraph("body")), "c");
        Assert.NotNull(comment.Date);

        DateTime date = new DateTime(2024, 5, 6, 7, 8, 9, DateTimeKind.Utc);
        comment.Date = date;
        Assert.Equal(date, comment.Date);

        comment.Date = null;
        Assert.Null(comment.Date);
    }

    [Fact]
    public void DeletingACommentRemovesAllOfItsParts()
    {
        Paragraph paragraph = AddParagraph("body");
        CommentWrapper keep = _document.Comments.Add(new RangeWrapper(paragraph), "keep");
        CommentWrapper remove = _document.Comments.Add(new RangeWrapper(paragraph), "remove");

        remove.Delete();

        Assert.Equal(1, _document.Comments.Count);
        Assert.Single(_body.Descendants<CommentRangeStart>());
        Assert.Single(_body.Descendants<CommentRangeEnd>());
        Assert.Single(_body.Descendants<CommentReference>());
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>());
        Assert.Equal(keep.Index, _document.Comments[0].Index);
    }
}
