using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class OpenXmlExtensionsTest : System.IDisposable
{
    private WordprocessingDocumentWrapper _document1;

    public OpenXmlExtensionsTest()
    {
        _document1 = WordprocessingDocumentWrapper.Open(TestFile.Document1);
    }

    public void Dispose()
    {
        _document1.Close();
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void RootElementTableNestingLevelIsZero()
    {
        Assert.Equal(0, _document1.WordprocessingDocument.MainDocumentPart.Document.NestingLevel<Table>());
    }
}
