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
    public void CollectionsCanBeUsedWithLinqWithoutACast()
    {
        IEnumerable<ParagraphWrapper> paragraphs = _document1.Paragraphs;

        Assert.Equal(_document1.Paragraphs.Count, paragraphs.Count());
        Assert.Equal("Test Document", paragraphs.Select(p => p.Text).First());
    }

    [Fact]
    public void EnumeratingVisitsEachParagraphOnceInOrder()
    {
        List<ParagraphWrapper> visited = new List<ParagraphWrapper>();
        foreach (ParagraphWrapper paragraph in _document1.Paragraphs)
        {
            visited.Add(paragraph);
        }

        Assert.Equal(_document1.Paragraphs.Count, visited.Count);
        for (int i = 0; i < visited.Count; i++)
        {
            Assert.True(visited[i] == _document1.Paragraphs[i]);
        }
    }

    [Fact]
    public void NonGenericEnumerationStillWorks()
    {
        System.Collections.IEnumerable paragraphs = _document1.Paragraphs;
        int count = 0;
        foreach (object paragraph in paragraphs)
        {
            Assert.IsType<ParagraphWrapper>(paragraph);
            count++;
        }

        Assert.Equal(_document1.Paragraphs.Count, count);
    }

    [Fact]
    public void DeletingWhileEnumeratingDoesNotSkipItems()
    {
        DocumentFormat.OpenXml.Wordprocessing.Body body = new DocumentFormat.OpenXml.Wordprocessing.Body(
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(),
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(),
            new DocumentFormat.OpenXml.Wordprocessing.Paragraph());
        ParagraphWrapperCollection paragraphs = new RangeWrapper(body).Paragraphs;

        int visited = 0;
        foreach (ParagraphWrapper paragraph in paragraphs)
        {
            paragraph.Delete();
            visited++;
        }

        Assert.Equal(3, visited);
        //The last paragraph of a body stays, as in Word, so that the body is never left without one.
        Assert.Equal(1, paragraphs.Count);
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
