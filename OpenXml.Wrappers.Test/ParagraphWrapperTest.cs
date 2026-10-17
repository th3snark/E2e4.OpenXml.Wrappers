using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Word;
using OpenXml.Wrappers.Test.Utils;

namespace OpenXml.Wrappers.Test;

public class ParagraphWrapperTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public ParagraphWrapperTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void RangeTextPropertyIsCorrect()
    {
        Assert.Equal("Test Document\r", _document1.Range.FirstParagraph.Range.Text);
        Assert.Equal("Test Document", _document1.Range.FirstParagraph.Text);
    }

    [Fact]
    public void RangeFirstParagraphRangeTextPropertyIsCorrect()
    {
        Assert.Equal("Test Document\r", _document1.Range.FirstParagraph.Range.Text);
        Assert.Equal("Test Document", _document1.Range.FirstParagraph.Text);
    }
}
