using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class CommentWordMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public CommentWordMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "alpha beta gamma";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        string[] errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010).Validate(_document.WordprocessingDocument).Select(e => e.Description + " at " + e.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, string.Join("; ", errors));
    }

    //The text of a comment ends with the mark of its paragraph.
    private static string Plain(CommentWrapper comment)
    {
        return comment.Text.TrimEnd('\r');
    }

    [Fact]
    public void TheIndexIsThePositionAmongTheComments()
    {
        CommentWrapper second = _document.Comments.Add(_document.Range(6, 10), "second");
        CommentWrapper first = _document.Comments.Add(_document.Range(0, 5), "first");

        Assert.Equal(0, first.Index);
        Assert.Equal(1, second.Index);
        Assert.Equal("first", Plain(_document.Comments[0]));
    }

    [Fact]
    public void ACommentIsMarkedDoneAndNotDone()
    {
        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 5), "check this");
        Assert.False(comment.Done);

        comment.Done = true;
        Assert.True(comment.Done);
        Assert.True(_document.Comments[0].Done);
        AssertValid();

        comment.Done = false;
        Assert.False(comment.Done);
    }

    [Fact]
    public void ThreadInformationSurvivesChangingTheText()
    {
        CommentWrapper comment = _document.Comments.Add(_document.Range(0, 5), "check this");
        comment.Done = true;

        comment.Text = "changed";

        Assert.True(comment.Done);
        Assert.Equal("changed", Plain(comment));
        AssertValid();
    }

    [Fact]
    public void RepliesAreCommentsThatKnowTheirAncestor()
    {
        CommentWrapper comment = _document.Comments.Add(_document.Range(6, 10), "question");
        CommentWrapper first = comment.Replies.Add("answer one");
        CommentWrapper second = comment.Replies.Add("answer two");

        Assert.Equal(2, comment.Replies.Count);
        Assert.Equal(first, comment.Replies[0]);
        Assert.Equal(second, comment.Replies[1]);
        Assert.Equal(comment, first.Ancestor);
        Assert.Null(comment.Ancestor);
        Assert.Equal(0, first.Replies.Count);
        Assert.Equal(3, _document.Comments.Count);
        Assert.Equal(new[] { "question", "answer one", "answer two" }, _document.Comments.Select(Plain).ToArray());
        Assert.Equal("beta", first.Scope.Text);
        AssertValid();
    }

    [Fact]
    public void DeletingACommentDeletesItsReplies()
    {
        CommentWrapper comment = _document.Comments.Add(_document.Range(6, 10), "question");
        comment.Replies.Add("answer");
        _document.Comments.Add(_document.Range(11, 16), "other");

        comment.Delete();

        Assert.Equal(1, _document.Comments.Count);
        Assert.Equal("other", Plain(_document.Comments[0]));
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>());
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsExPart.CommentsEx.ChildElements);
        AssertValid();
    }

    [Fact]
    public void ACommentCanBeAddedToATableRowAndTheWholeDocument()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(16, 16), 2, 2);
        table.Cell(0, 0).Range.Text = "a";

        CommentWrapper onRow = _document.Comments.Add(table.Rows[0].Range, "row comment");
        CommentWrapper onTable = _document.Comments.Add(table.Range, "table comment");
        CommentWrapper onDocument = _document.Comments.Add(_document.Content, "document comment");

        Assert.Equal(3, _document.Comments.Count);
        Assert.Equal("table comment", Plain(onTable));
        Assert.Equal("document comment", Plain(onDocument));
        Assert.NotNull(onRow.Scope.Text);
        AssertValid();
    }

    [Fact]
    public void DeleteAllRemovesTheComments()
    {
        _document.Comments.Add(_document.Range(0, 5), "one");
        _document.Comments.Add(_document.Range(6, 10), "two");

        _document.Comments.DeleteAll();

        Assert.Equal(0, _document.Comments.Count);
        AssertValid();
    }
}
