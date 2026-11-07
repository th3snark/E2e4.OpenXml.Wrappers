using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// CleanString, the list galleries and XPath relative to a custom XML node. The expected values are those Word 16 gave for the same calls
/// (the application-clean-string, application-list-galleries and custom-xml-node-select scenarios of tools/WordReference/scenarios.ps1,
/// which print them).
/// </summary>
public class ApplicationExtrasTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ApplicationExtrasTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void CleanStringReplacesAndDropsTheCharactersWordDoes()
    {
        ApplicationWrapper application = _document.Application;
        Assert.Equal("a b c", application.CleanString("a" + (char)160 + "b" + (char)1 + "c"));
        Assert.Equal("ab", application.CleanString("a" + (char)7 + (char)31 + (char)172 + (char)182 + "b"));
        Assert.Equal("a\rb-c", application.CleanString("a\nb" + (char)30 + "c"));
        //Tab, paragraph mark and optional hyphen stay.
        Assert.Equal("a\tb\rc" + (char)173, application.CleanString("a\tb\rc" + (char)173));
        Assert.Equal(string.Empty, application.CleanString(string.Empty));
        Assert.Throws<ArgumentNullException>(() => application.CleanString(null));
    }

    [Fact]
    public void TheListGalleriesHaveSevenTemplatesEach()
    {
        ListGalleryWrapperCollection galleries = _document.Application.ListGalleries;
        Assert.Equal(3, galleries.Count);
        Assert.Equal(new[] { 7, 7, 7 }, galleries.Select(g => g.ListTemplates.Count).ToArray());

        //Word's values: the second number template is "1)", the fourth outline is "Article 1." linked to Heading 1.
        Assert.Equal("%1)", galleries[ListGalleryType.Number].ListTemplates[1].ListLevels[0].NumberFormat);
        ListTemplateWrapper outline = galleries[ListGalleryType.Outline].ListTemplates[3];
        Assert.True(outline.OutlineNumbered);
        Assert.Equal(9, outline.ListLevels.Count);
        Assert.Equal("Article %1.", outline.ListLevels[0].NumberFormat);
        Assert.Equal("Section %1.%2", outline.ListLevels[1].NumberFormat);
        Assert.Equal(ListNumberStyle.UppercaseRoman, outline.ListLevels[0].NumberStyle);
        Assert.False(galleries[ListGalleryType.Bullet].Modified(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => galleries[ListGalleryType.Bullet].Modified(7));
        Assert.Throws<ArgumentOutOfRangeException>(() => galleries[(ListGalleryType)9]);
        Assert.Throws<ArgumentOutOfRangeException>(() => galleries[3]);
        Assert.Throws<ArgumentOutOfRangeException>(() => galleries[0].ListTemplates[7]);

        //Applying a template of a gallery adds it to the document.
        Assert.Equal(0, _document.ListTemplates.Count);
        _document.Paragraphs[0].Range.ListFormat.ApplyListTemplate(outline, false);
        Assert.Equal(1, _document.ListTemplates.Count);
        Assert.Equal("Heading 1", _document.ListTemplates[0].ListLevels[0].LinkedStyle);
        Assert.Equal("Heading 1", _document.Paragraphs[0].Range.ParagraphStyle);
        Assert.Null(DocumentIntegrity.CheckAll(_document.WordprocessingDocument).FirstOrDefault());
    }

    [Fact]
    public void XPathOfANodeStartsAtTheNode()
    {
        CustomXMLPart part = _document.CustomXMLParts.Add("<root xmlns=\"urn:test\"><a><b>1</b><b>2</b></a><c id=\"x\"><b>3</b></c></root>");
        part.NamespaceManager.AddNamespace("ns", "urn:test");
        CustomXMLNode a = part.SelectSingleNode("/ns:root/ns:a");

        Assert.Equal(new[] { "1", "2" }, a.SelectNodes("ns:b").Select(n => n.Text).ToArray());
        Assert.Equal(new[] { "1", "2" }, a.SelectNodes("./ns:b").Select(n => n.Text).ToArray());
        Assert.Equal(new[] { "1", "2" }, a.SelectNodes(".//ns:b").Select(n => n.Text).ToArray());
        Assert.Equal(new[] { "3" }, a.SelectNodes("/ns:root/ns:c/ns:b").Select(n => n.Text).ToArray());
        Assert.Equal("root", a.SelectNodes("..").Single().BaseName);
        Assert.Equal("id", part.SelectSingleNode("/ns:root/ns:c").SelectNodes("@id").Single().BaseName);
        Assert.Equal("2", a.SelectSingleNode("ns:b[2]").Text);
        Assert.Null(a.SelectSingleNode("ns:zzz"));
        //A name without a prefix has no namespace, so it matches nothing in this part.
        Assert.Equal(0, a.SelectNodes("b").Count);
        Assert.Throws<ArgumentNullException>(() => a.SelectNodes(null));
    }
}
