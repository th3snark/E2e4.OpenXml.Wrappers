using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class CommentWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document3;

    public CommentWrapperTest()
    {
        _document3 = WordprocessingDocumentWrapper.Open(TestFile.Document3);
    }

    public void Dispose()
    {
        _document3.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void CommentWrappersOfTheSameCommentAreEqual()
    {
        Assert.True(_document3.Comments[0] == _document3.Comments[0]);
        Assert.Equal(_document3.Comments[0].GetHashCode(), _document3.Comments[0].GetHashCode());
        if (_document3.Comments.Count > 1)
        {
            Assert.True(_document3.Comments[0] != _document3.Comments[1]);
        }
    }

    [Fact]
    public void IndexPropertyIsCorrect()
    {
        for (int index = 0; index < _document3.Comments.Count; index++)
        {
            Assert.Equal(index, _document3.Comments[index].Index);
        }
    }

    [Fact]
    public void TextPropertyIsCorrect()
    {
        Assert.Equal("Comment text\r\tPara2\rPara3\r", _document3.Comments[0].Text);
        Assert.Equal("Comment 2 - Nice\r", _document3.Comments[1].Text);
        Assert.Equal("Multi line comment\r", _document3.Comments[2].Text);
        Assert.Equal("A cell comment\r", _document3.Comments[3].Text);
        Assert.Equal("Multi cell comment.\r", _document3.Comments[4].Text);
    }

    [Fact]
    public void AuthorPropertyIsCorrect()
    {
        Assert.Equal("nobody", _document3.Comments[0].Author);
        Assert.Equal("nobody", _document3.Comments[1].Author);
        Assert.Equal("nobody", _document3.Comments[2].Author);
        Assert.Equal("nobody", _document3.Comments[3].Author);
        Assert.Equal("nobody", _document3.Comments[4].Author);
    }

    [Fact]
    public void InitialPropertyIsCorrect()
    {
        Assert.Equal("n", _document3.Comments[0].Initial);
        Assert.Equal("n", _document3.Comments[1].Initial);
        Assert.Equal("n", _document3.Comments[2].Initial);
        Assert.Equal("n", _document3.Comments[3].Initial);
        Assert.Equal("n", _document3.Comments[4].Initial);
    }

    [Fact]
    public void RangeTextPropertyIsCorrect()
    {
        Assert.Equal("text", _document3.Comments[0].Scope.Text);
        Assert.Equal("Another paragraph with a comment\r", _document3.Comments[1].Scope.Text);
        Assert.Equal("Line 2\rLine 3\rLin", _document3.Comments[2].Scope.Text);
        Assert.Equal("R2C2\r", _document3.Comments[3].Scope.Text);
        Assert.Equal("R3C1\rR3C2\r", _document3.Comments[4].Scope.Text);
    }
}
