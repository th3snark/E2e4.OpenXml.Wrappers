using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class DocumentIntegrityTest
{
    private static IList<string> CheckAfter(System.Action<WordprocessingDocument> damage)
    {
        using (MemoryStream stream = new MemoryStream())
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
            {
                document.Content.InsertAfter("Text");
                damage(document.WordprocessingDocument);
            }
            stream.Position = 0;
            using (WordprocessingDocument document = WordprocessingDocument.Open(stream, false))
            {
                return DocumentIntegrity.CheckAll(document);
            }
        }
    }

    [Fact]
    public void ANewDocumentPassesEveryCheck()
    {
        Assert.Empty(CheckAfter(d => { }));
    }

    [Fact]
    public void BrokenReferencesAndRepeatedIdsAreFound()
    {
        IList<string> problems = CheckAfter(d =>
        {
            Body body = d.MainDocumentPart.Document.Body;
            Paragraph paragraph = body.GetFirstChild<Paragraph>();
            paragraph.PrependChild(new ParagraphProperties(new ParagraphStyleId() { Val = "Missing" }, new NumberingProperties(new NumberingId() { Val = 7 })));
            paragraph.ParagraphId = "00000001";
            paragraph.Append(new BookmarkStart() { Id = "1", Name = "a" }, new BookmarkEnd() { Id = "1" }, new BookmarkStart() { Id = "1", Name = "b" },
                new BookmarkEnd() { Id = "1" });
            body.Append(new Paragraph() { ParagraphId = "00000001" });
            body.GetFirstChild<SectionProperties>().PrependChild(new HeaderReference() { Id = "rId99", Type = HeaderFooterValues.Default });
        });

        Assert.Contains("Styles: pStyle Missing in /word/document.xml names no style", problems);
        Assert.Contains("Numbering: numId 7 in /word/document.xml names no num", problems);
        Assert.Contains("Relationships: r:id=\"rId99\" on headerReference in /word/document.xml names no relationship", problems);
        Assert.Contains("BookmarkIds: bookmark id 1 is used 2 times (/word/document.xml)", problems);
        Assert.Contains("ParagraphIds: w14:paraId 00000001 is used 2 times (/word/document.xml)", problems);
    }
}
