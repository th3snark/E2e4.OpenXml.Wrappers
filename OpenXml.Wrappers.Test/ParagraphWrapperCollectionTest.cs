using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Xunit;
using OpenXml.Wrappers.Word;
using OpenXml.Wrappers.Test.Utils;

namespace OpenXml.Wrappers.Test;

public class ParagraphWrapperCollectionTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public ParagraphWrapperCollectionTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void CountPropertyIsCorrectForDocument()
    {
        Assert.Equal(184, _document1.Paragraphs.Count);
    }

    [Fact]
    public void CountPropertyIsCorrectForTable()
    {
        Assert.Equal(40, _document1.Tables[0].Range.Paragraphs.Count);
        Assert.Equal(121, _document1.Tables[1].Range.Paragraphs.Count);
    }
}
