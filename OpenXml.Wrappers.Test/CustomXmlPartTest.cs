using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Custom XML parts and the binding of content controls to them. The markup is Word 16's (the custom-xml-* and
/// content-control-xml-mapping* scenarios of MarkupConformanceTest).
/// </summary>
public class CustomXmlPartTest : IDisposable
{
    private const string Xml = "<root xmlns=\"urn:test\"><name>Bob</name><city>Oslo</city></root>";
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public CustomXmlPartTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    [Fact]
    public void APartIsAddedFoundReplacedAndDeleted()
    {
        Assert.Equal(0, _document.CustomXMLParts.Count);
        CustomXMLPart part = _document.CustomXMLParts.Add(Xml);

        Assert.Equal(1, _document.CustomXMLParts.Count);
        Assert.Equal(part, _document.CustomXMLParts[0]);
        Assert.Equal("urn:test", part.NamespaceURI);
        Assert.Equal(Xml, part.XML);
        Assert.False(part.BuiltIn);
        Assert.Matches("^\\{[0-9A-F-]{36}\\}$", part.Id);
        Assert.Equal(part, _document.CustomXMLParts.SelectByID(part.Id.ToLowerInvariant()));
        Assert.Null(_document.CustomXMLParts.SelectByID("{none}"));
        Assert.Single(_document.CustomXMLParts.SelectByNamespace("urn:test"));
        Assert.Empty(_document.CustomXMLParts.SelectByNamespace("urn:other"));

        part.LoadXML("<a xmlns=\"urn:new\"/>");
        Assert.Equal("urn:new", part.NamespaceURI);
        AssertValid();

        part.Delete();
        Assert.Equal(0, _document.CustomXMLParts.Count);
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.CustomXMLParts[0]);
        AssertValid();
    }

    [Fact]
    public void AControlShowsTheTextOfTheElementItIsBoundTo()
    {
        CustomXMLPart part = _document.CustomXMLParts.Add(Xml);
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));

        Assert.False(control.XMLMapping.IsMapped);
        Assert.True(control.XMLMapping.SetMapping("/ns:root/ns:city", "xmlns:ns='urn:test'", part));

        Assert.True(control.XMLMapping.IsMapped);
        Assert.Equal("/ns:root/ns:city", control.XMLMapping.XPath);
        Assert.Equal("xmlns:ns='urn:test'", control.XMLMapping.PrefixMappings);
        Assert.Equal(part, control.XMLMapping.CustomXMLPart);
        Assert.Equal("Oslo", control.Range.Text);
        Assert.False(control.ShowingPlaceholderText);
        AssertValid();
    }

    [Fact]
    public void WithoutAPartTheFirstOneThatMatchesIsUsed()
    {
        _document.CustomXMLParts.Add("<other xmlns=\"urn:other\"/>");
        CustomXMLPart part = _document.CustomXMLParts.Add(Xml);
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));

        Assert.True(control.XMLMapping.SetMapping("/a:root/a:name", "xmlns:a=\"urn:test\""));
        Assert.Equal(part, control.XMLMapping.CustomXMLPart);
        Assert.Equal("Bob", control.Range.Text);
    }

    [Fact]
    public void AMappingThatMatchesNothingChangesNothing()
    {
        CustomXMLPart part = _document.CustomXMLParts.Add(Xml);
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));

        Assert.False(control.XMLMapping.SetMapping("/ns:root/ns:missing", "xmlns:ns='urn:test'", part));

        Assert.False(control.XMLMapping.IsMapped);
        Assert.True(control.ShowingPlaceholderText);
    }

    [Fact]
    public void DeletingTheMappingOrThePartKeepsTheText()
    {
        CustomXMLPart part = _document.CustomXMLParts.Add(Xml);
        ContentControlWrapper first = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));
        first.XMLMapping.SetMapping("/ns:root/ns:name", "xmlns:ns='urn:test'", part);
        ContentControlWrapper second = _document.ContentControls.Add(ContentControlType.Text, _document.Range(_document.ContentControls[0].Range.End + 7, _document.ContentControls[0].Range.End + 7));
        second.XMLMapping.SetMapping("/ns:root/ns:city", "xmlns:ns='urn:test'", part);

        first.XMLMapping.Delete();
        Assert.False(first.XMLMapping.IsMapped);
        Assert.Equal("Bob", first.Range.Text);
        Assert.Null(first.XMLMapping.CustomXMLPart);

        part.Delete();
        Assert.False(second.XMLMapping.IsMapped);
        Assert.Equal("Oslo", second.Range.Text);
        AssertValid();
    }
}
