using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class CommentWrapperCollectionTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document3;

    public CommentWrapperCollectionTest()
    {
        _document3 = WordprocessingDocumentWrapper.Open(TestFile.Document3, true);
    }

    public void Dispose()
    {
        _document3.Close();
        GC.SuppressFinalize(this);
    }

    private static WordprocessingDocumentWrapper CreateTemporaryDocument(string path)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, true);
        document.InitializeMainDocumentPart();
        return document;
    }

    [Fact]
    public void AddToAnUnsupportedRangeThrowsAndLeavesNoOrphanComment()
    {
        string path = System.IO.Path.GetTempFileName();
        try
        {
            WordprocessingDocumentWrapper document = CreateTemporaryDocument(path);
            try
            {
                Assert.Throws<NotSupportedException>(() => document.Comments.Add(new RangeWrapper(new Run()), "text"));
                WordprocessingCommentsPart part = document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart;
                Assert.True(part?.Comments == null || !part.Comments.Elements<Comment>().Any());
            }
            finally
            {
                document.Close();
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void AddToAnEmptyParagraphWrapsTheReferenceInARunAndCreatesTheCommentsRoot()
    {
        string path = System.IO.Path.GetTempFileName();
        try
        {
            WordprocessingDocumentWrapper document = CreateTemporaryDocument(path);
            try
            {
                Paragraph paragraph = document.WordprocessingDocument.MainDocumentPart.Document.Body.Elements<Paragraph>().First();
                CommentWrapper comment = document.Comments.Add(new ParagraphWrapper(paragraph).Range, "text");

                Assert.Equal("text\r", comment.Text);
                Assert.Equal(1, document.Comments.Count);
                CommentReference reference = paragraph.Descendants<CommentReference>().Single();
                Assert.IsType<Run>(reference.Parent);
                Assert.NotNull(comment.Author);
            }
            finally
            {
                document.Close();
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void AddUsesOneMoreThanTheLargestExistingId()
    {
        string path = System.IO.Path.GetTempFileName();
        try
        {
            WordprocessingDocumentWrapper document = CreateTemporaryDocument(path);
            try
            {
                Body body = document.WordprocessingDocument.MainDocumentPart.Document.Body;
                Paragraph paragraph = body.Elements<Paragraph>().First();
                paragraph.AppendChild(new Run(new Text("a")));
                CommentWrapper first = document.Comments.Add(new ParagraphWrapper(paragraph).Range, "one");
                CommentWrapper second = document.Comments.Add(new ParagraphWrapper(paragraph).Range, "two");
                //Make the ids unordered: the last comment in the part no longer has the largest id.
                Comments comments = document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments;
                Comment last = comments.Elements<Comment>().Last();
                last.Id = "0";
                comments.Elements<Comment>().First().Id = "5";

                CommentWrapper third = document.Comments.Add(new ParagraphWrapper(paragraph).Range, "three");

                Assert.Equal("6", comments.Elements<Comment>().Last().Id.Value);
                Assert.Equal(2, third.Index);
            }
            finally
            {
                document.Close();
            }
        }
        finally
        {
            System.IO.File.Delete(path);
        }
    }

    [Fact]
    public void AddCommentIsCorrect()
    {
        CommentWrapper comment = _document3.Comments.Add(_document3.Content.LastParagraph.Range, "**Comment Text**");
    }
}
