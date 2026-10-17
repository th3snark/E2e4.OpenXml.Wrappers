using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
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

    [Fact]
    public void AddCommentIsCorrect()
    {
        CommentWrapper comment = _document3.Comments.Add(_document3.Range.LastParagraph.Range, "**Comment Text**");
    }
}
