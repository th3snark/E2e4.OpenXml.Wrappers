using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The placeholder text of content controls and the editing of custom XML nodes. The markup of the placeholders is Word 16's (the
/// content-control-placeholder-text scenario of MarkupConformanceTest); the values are those Word printed for the same calls.
/// </summary>
public class ContentControlPlaceholderTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ContentControlPlaceholderTest()
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
        IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    [Fact]
    public void AControlShowingItsPlaceholderShowsTheNewText()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));
        Assert.Equal("Click or tap here to enter text.", control.PlaceholderText.Value);

        control.SetPlaceholderText(null, null, "Type here");

        Assert.True(control.ShowingPlaceholderText);
        Assert.Equal("Type here", control.Range.Text);
        Assert.Equal("Type here", control.PlaceholderText.Value);
        Assert.Equal(32, control.PlaceholderText.Name.Length);

        control.SetPlaceholderText(null, null, "Changed");
        Assert.Equal("Changed", control.Range.Text);
        AssertValid();
    }

    [Fact]
    public void AControlHoldingTextKeepsItAndRemembersThePlaceholder()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(12, 17));

        control.SetPlaceholderText(null, null, "Other");

        Assert.False(control.ShowingPlaceholderText);
        Assert.Equal("again", control.Range.Text);
        Assert.Equal("Other", control.PlaceholderText.Value);
        AssertValid();
    }

    [Fact]
    public void ThePlaceholderCanComeFromARangeAndTheArgumentsAreChecked()
    {
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(12, 17));

        control.SetPlaceholderText(null, _document.Range(0, 5));

        Assert.Equal("Hello", control.PlaceholderText.Value);
        Assert.Throws<ArgumentException>(() => control.SetPlaceholderText());
        Assert.Throws<NotSupportedException>(() => control.SetPlaceholderText(control.PlaceholderText, null, null));
    }

    [Fact]
    public void NodesAreAddedInsertedAndReplaced()
    {
        CustomXMLPart part = _document.CustomXMLParts.Add("<root xmlns=\"urn:test\"><name>Bob</name></root>");
        part.NamespaceManager.AddNamespace("ns", "urn:test");
        CustomXMLNode root = part.DocumentElement;

        root.AppendChildNode("age", "urn:test", CustomXMLNodeType.Element, "42");
        root.AppendChildNode("note", string.Empty, CustomXMLNodeType.Element, "plain");
        root.AppendChildNode("id", string.Empty, CustomXMLNodeType.Attribute, "7");
        root.AppendChildSubtree("<city xmlns=\"urn:test\">Paris</city>");
        CustomXMLNode name = part.SelectSingleNode("/ns:root/ns:name");
        root.InsertNodeBefore("first", "urn:test", CustomXMLNodeType.Element, "x", name);
        root.ReplaceChildNode(name, "surname", "urn:test", CustomXMLNodeType.Element, "Smith");
        part.AddNode(root, "extra", "urn:test", CustomXMLNodeType.Element, "e");

        XNamespace ns = "urn:test";
        XElement document = XDocument.Parse(part.XML).Root;
        Assert.Equal(new[] { "first", "surname", "age", "note", "city", "extra" }, document.Elements().Select(e => e.Name.LocalName).ToArray());
        Assert.Equal(string.Empty, document.Element("note").Name.NamespaceName);
        Assert.Equal("7", (string)document.Attribute("id"));
        Assert.Equal("Smith", (string)document.Element(ns + "surname"));

        root.ReplaceChildSubtree("<town xmlns=\"urn:test\">Rome</town>", part.SelectSingleNode("/ns:root/ns:city"));
        root.InsertSubtreeBefore("<zero xmlns=\"urn:test\"/>", part.SelectSingleNode("/ns:root/ns:first"));
        Assert.Equal(new[] { "zero", "first", "surname", "age", "note", "town", "extra" }, XDocument.Parse(part.XML).Root.Elements().Select(e => e.Name.LocalName).ToArray());
        Assert.Throws<ArgumentNullException>(() => root.ReplaceChildNode(null, "x"));
        Assert.Throws<ArgumentNullException>(() => part.AddNode(null, "x"));
    }

    [Fact]
    public void ABoundControlFollowsAnAddedNode()
    {
        CustomXMLPart part = _document.CustomXMLParts.Add("<root xmlns=\"urn:test\"><name>Bob</name></root>");
        part.NamespaceManager.AddNamespace("ns", "urn:test");
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));
        control.XMLMapping.SetMappingByNode(part.SelectSingleNode("/ns:root/ns:name"));

        part.DocumentElement.ReplaceChildNode(part.SelectSingleNode("/ns:root/ns:name"), "name", "urn:test", CustomXMLNodeType.Element, "Carol");

        Assert.Equal("Carol", control.Range.Text);
    }
}
