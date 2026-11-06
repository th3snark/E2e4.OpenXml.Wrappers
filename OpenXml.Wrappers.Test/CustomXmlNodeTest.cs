using System;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The nodes of a custom XML part. The markup is Word 16's (the content-control-xml-node scenario of MarkupConformanceTest); the values are
/// those Word printed for the same calls.
/// </summary>
public class CustomXmlNodeTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;
    private readonly CustomXMLPart _part;

    public CustomXmlNodeTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
        _part = _document.CustomXMLParts.Add("<root xmlns=\"urn:test\" id=\"7\"><name>Bob</name><age>42</age></root>");
        _part.NamespaceManager.AddNamespace("ns", "urn:test");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void ANodeKnowsItsNamePositionAndText()
    {
        CustomXMLNode node = _part.SelectSingleNode("/ns:root/ns:name");

        Assert.Equal("name", node.BaseName);
        Assert.Equal("urn:test", node.NamespaceURI);
        Assert.Equal("Bob", node.Text);
        Assert.Equal("<name xmlns=\"urn:test\">Bob</name>", node.XML);
        Assert.Equal(CustomXMLNodeType.Element, node.NodeType);
        Assert.Equal("/ns0:root[1]/ns0:name[1]", node.XPath);
        Assert.Equal(1, node.ChildNodes.Count);
        Assert.True(node.HasChildNodes());
        Assert.Equal("root", node.ParentNode.BaseName);
        Assert.Equal("age", node.NextSibling.BaseName);
        Assert.Null(node.PreviousSibling);
        Assert.Equal(_part, node.OwnerPart);
    }

    [Fact]
    public void TheRootHasChildrenAndAttributes()
    {
        CustomXMLNode root = _part.DocumentElement;

        Assert.Equal("root", root.BaseName);
        Assert.Equal(2, root.ChildNodes.Count);
        Assert.Equal(1, root.Attributes.Count);
        Assert.Equal("name", root.FirstChild.BaseName);
        Assert.Equal("age", root.LastChild.BaseName);
        Assert.Equal("id", root.Attributes[0].BaseName);
        Assert.Equal("7", root.Attributes[0].Text);
        Assert.Equal(CustomXMLNodeType.Attribute, root.Attributes[0].NodeType);
        Assert.Null(root.ParentNode);
    }

    [Fact]
    public void NodesAreSelectedByExpression()
    {
        Assert.Equal(1, _part.SelectNodes("//ns:name").Count);
        Assert.Equal(2, _part.SelectNodes("/ns:root/*").Count);
        Assert.Equal(0, _part.SelectNodes("//ns:zzz").Count);
        Assert.Null(_part.SelectSingleNode("//ns:zzz"));
        Assert.Throws<ArgumentNullException>(() => _part.SelectNodes(null));
    }

    [Fact]
    public void ChangingANodeChangesThePartAndTheControlsBoundToIt()
    {
        CustomXMLNode node = _part.SelectSingleNode("/ns:root/ns:name");
        ContentControlWrapper control = _document.ContentControls.Add(ContentControlType.Text, _document.Range(5, 5));

        Assert.True(control.XMLMapping.SetMappingByNode(node));
        Assert.True(control.XMLMapping.IsMapped);
        Assert.Equal("/ns0:root[1]/ns0:name[1]", control.XMLMapping.XPath);
        Assert.Equal("xmlns:ns0='urn:test' ", control.XMLMapping.PrefixMappings);
        Assert.Equal(node, control.XMLMapping.CustomXMLNode);
        Assert.Equal("Bob", control.Range.Text);
        Assert.Equal(1, _document.SelectLinkedControls(node).Count);
        Assert.Equal(0, _document.SelectUnlinkedControls(_part).Count);

        node.Text = "Alice";
        Assert.Equal("Alice", control.Range.Text);
        Assert.Contains("<name>Alice</name>", _part.XML, StringComparison.Ordinal);

        node.Delete();
        Assert.DoesNotContain("<name>", _part.XML, StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => node.Text);
        Assert.Null(control.XMLMapping.CustomXMLNode);
        Assert.Throws<ArgumentNullException>(() => control.XMLMapping.SetMappingByNode(null));
    }

    [Fact]
    public void AnAttributeCanBeChanged()
    {
        CustomXMLNode id = _part.DocumentElement.Attributes[0];

        id.Text = "8";

        Assert.Contains("id=\"8\"", _part.XML, StringComparison.Ordinal);
        Assert.Equal("/ns0:root[1]/@id", id.XPath);
        Assert.Equal("urn:test", _part.NamespaceManager.LookupNamespace("ns"));
        Assert.Equal("ns", _part.NamespaceManager.LookupPrefix("urn:test"));
        Assert.Equal(1, _part.NamespaceManager.Count);
        Assert.True(_document.SelectUnlinkedControls().Count == 0);
    }
}
