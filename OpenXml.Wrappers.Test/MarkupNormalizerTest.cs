using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using OpenXml.Wrappers.Test.Utils;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class MarkupNormalizerTest
{
    private const string Namespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" "
        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
        + "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\"";

    private static SortedDictionary<string, string> Normalize(params string[] entries)
    {
        using (MemoryStream stream = new MemoryStream())
        {
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
            {
                for (int i = 0; i < entries.Length; i += 2)
                {
                    using (StreamWriter writer = new StreamWriter(archive.CreateEntry(entries[i]).Open(), new UTF8Encoding(false)))
                    {
                        writer.Write(entries[i + 1]);
                    }
                }
            }
            stream.Position = 0;
            return MarkupNormalizer.NormalizePackage(stream);
        }
    }

    [Fact]
    public void VolatileValuesAreRemovedOrNumberedAndTheMarkupIsCanonical()
    {
        SortedDictionary<string, string> parts = Normalize(
            "word/document.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<w:document " + Namespaces + " mc:Ignorable=\"wp14 w14\"><w:body>"
            + "<w:p w14:paraId=\"1A2B3C4D\" w14:textId=\"77777777\" w:rsidR=\"00AB12CD\" w:rsidRDefault=\"00AB12CD\">"
            + "<w:bookmarkStart w:name=\"_GoBack\" w:id=\"0\"/><w:bookmarkStart w:name=\"Mark\" w:id=\"5\"/><w:proofErr w:type=\"spellStart\"/>"
            + "<w:r w:rsidR=\"00AB12CD\"><w:t xml:space=\"preserve\"> a </w:t></w:r><w:bookmarkEnd w:id=\"0\"/><w:bookmarkEnd w:id=\"5\"/>"
            + "<w:r><w:rPr><w:sz w:val=\"20\"/></w:rPr><w:lastRenderedPageBreak/><w:t> </w:t></w:r></w:p>"
            + "<w:sectPr><w:headerReference r:id=\"rId9\" w:type=\"default\"/><w:pgMar w:top=\"1\" w:left=\"2\"/></w:sectPr></w:body></w:document>",
            "word/_rels/document.xml.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
            + "<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"/word/header1.xml\"/>"
            + "</Relationships>",
            "docProps/app.xml",
            "<Properties/>");

        Assert.Equal(new[] { "document.xml", "document.xml.rels" }, parts.Keys);
        Assert.Equal(
            "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
            + "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n"
            + "  <w:body>\n"
            + "    <w:p>\n"
            + "      <w:bookmarkStart w:id=\"{bookmark1}\" w:name=\"Mark\" />\n"
            + "      <w:r>\n"
            + "        <w:t xml:space=\"preserve\"> a </w:t>\n"
            + "      </w:r>\n"
            + "      <w:bookmarkEnd w:id=\"{bookmark1}\" />\n"
            + "      <w:r>\n"
            + "        <w:rPr>\n"
            + "          <w:sz w:val=\"20\" />\n"
            + "        </w:rPr>\n"
            + "        <w:t> </w:t>\n"
            + "      </w:r>\n"
            + "    </w:p>\n"
            + "    <w:sectPr>\n"
            + "      <w:headerReference r:id=\"{rId1}\" w:type=\"default\" />\n"
            + "      <w:pgMar w:left=\"2\" w:top=\"1\" />\n"
            + "    </w:sectPr>\n"
            + "  </w:body>\n"
            + "</w:document>\n",
            parts["document.xml"]);
        Assert.Equal(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            + "  <Relationship Id=\"{rId1}\" Target=\"/word/header1.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" />\n"
            + "  <Relationship Id=\"{rId2}\" Target=\"/word/styles.xml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" />\n"
            + "</Relationships>\n",
            parts["document.xml.rels"]);
    }

    [Fact]
    public void IgnorableKeepsOnlyThePrefixesStillUsedAndCommentParagraphIdsStayLinked()
    {
        SortedDictionary<string, string> parts = Normalize(
            "word/comments.xml",
            "<w:comments " + Namespaces + " mc:Ignorable=\"wp14 w14\"><w:comment w:id=\"3\" w:author=\"A\" w:date=\"2026-11-04T09:00:00Z\">"
            + "<w:p w14:paraId=\"0000ABCD\"/></w:comment></w:comments>",
            "word/commentsExtended.xml",
            "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:commentEx w15:paraId=\"0000ABCD\" w15:done=\"0\"/></w15:commentsEx>");

        Assert.Contains("mc:Ignorable=\"w14\"", parts["comments.xml"]);
        Assert.DoesNotContain("wp14", parts["comments.xml"]);
        Assert.Contains("<w:comment w:author=\"A\" w:date=\"{date}\" w:id=\"{comment1}\">", parts["comments.xml"]);
        Assert.Contains("<w:p w14:paraId=\"{para1}\" />", parts["comments.xml"]);
        Assert.Contains("<w15:commentEx w15:done=\"0\" w15:paraId=\"{para1}\" />", parts["commentsExtended.xml"]);
    }

    [Fact]
    public void PartNamesAreFlattenedAndUncomparedPartsAreLeftOut()
    {
        Assert.Equal("content-types.xml", MarkupNormalizer.GetPartFileName("[Content_Types].xml"));
        Assert.Equal("package.rels", MarkupNormalizer.GetPartFileName("_rels/.rels"));
        Assert.Equal("document.xml.rels", MarkupNormalizer.GetPartFileName("word/_rels/document.xml.rels"));
        Assert.Equal("theme.theme1.xml", MarkupNormalizer.GetPartFileName("word/theme/theme1.xml"));
        Assert.Null(MarkupNormalizer.GetPartFileName("docProps/core.xml"));
        Assert.Null(MarkupNormalizer.GetPartFileName("word/media/image1.png"));
    }
}
