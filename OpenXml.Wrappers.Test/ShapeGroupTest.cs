using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Grouping and ungrouping shapes. The expected values are those Word 16 gave for the same calls (the shape-group and shape-ungroup
/// scenarios of tools/WordReference/scenarios.ps1, which print them).
/// </summary>
public class ShapeGroupTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ShapeGroupTest()
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

    private ShapeWrapper Rectangle(string name, double left, double top, double width, double height)
    {
        ShapeWrapper shape = _document.Shapes.AddShape(AutoShapeType.Rectangle, left, top, width, height, _document.Range(5, 5));
        shape.Name = name;
        return shape;
    }

    [Fact]
    public void AGroupTakesThePlaceOfItsShapes()
    {
        ShapeWrapper a = Rectangle("A", 20, 30, 80, 40);
        ShapeWrapper b = Rectangle("B", 120, 60, 60, 60);
        ShapeWrapper group = _document.Shapes.Group(new[] { b, a });

        Assert.Equal(1, _document.Shapes.Count);
        Assert.Equal(ShapeType.Group, group.Type);
        Assert.StartsWith("Group ", group.Name, StringComparison.Ordinal);
        Assert.Equal((160.2, 90), (group.Width, group.Height));
        Assert.Equal(2, group.GroupItems.Count);
        Assert.Equal(new[] { "A", "B" }, group.GroupItems.Select(s => s.Name).ToArray());
        Assert.All(group.GroupItems, item =>
        {
            Assert.True(item.Child);
            Assert.Equal(group, item.ParentGroup);
            Assert.Equal(ShapeType.AutoShape, item.Type);
            Assert.Equal(AutoShapeType.Rectangle, item.AutoShapeType);
        });
        Assert.Equal((80, 40), (group.GroupItems[0].Width, group.GroupItems[0].Height));
        Assert.NotEqual(group.GroupItems[0], group.GroupItems[1]);
        Assert.NotEqual(group.GroupItems[0].ID, group.GroupItems[1].ID);
        Assert.Empty(DocumentIntegrity.CheckAll(_document.WordprocessingDocument));
    }

    [Fact]
    public void ResizingAGroupScalesItsShapes()
    {
        ShapeWrapper group = _document.Shapes.Group(new[] { Rectangle("A", 20, 30, 80, 40), Rectangle("B", 120, 60, 60, 60) });
        group.Width = group.Width * 2;
        group.Height = group.Height / 2;
        Assert.Equal((160, 20), (group.GroupItems[0].Width, group.GroupItems[0].Height));
        group.GroupItems[0].Name = "First";
        Assert.Equal("First", group.GroupItems[0].Name);
        Assert.Throws<NotSupportedException>(() => group.GroupItems[0].Left);
        Assert.Throws<NotSupportedException>(() => group.GroupItems[0].Width = 10);
        Assert.Throws<NotSupportedException>(() => group.GroupItems[0].Delete());
        Assert.Throws<NotSupportedException>(() => group.Rotation = 10);
        Assert.Throws<NotSupportedException>(() => group.Flip(FlipCommand.Horizontal));
        Assert.Throws<InvalidOperationException>(() => group.Fill);
        Assert.Throws<InvalidOperationException>(() => group.ParentGroup);
        Assert.Empty(DocumentIntegrity.CheckAll(_document.WordprocessingDocument));
    }

    [Fact]
    public void UngroupingGivesTheShapesBackAtTheirPlaces()
    {
        ShapeWrapper group = _document.Shapes.Group(new[] { Rectangle("A", 20, 30, 80, 40), Rectangle("B", 120, 60, 60, 60) });
        System.Collections.Generic.IList<ShapeWrapper> shapes = group.Ungroup();

        Assert.Equal(new[] { "A", "B" }, shapes.Select(s => s.Name).ToArray());
        Assert.Equal(2, _document.Shapes.Count);
        Assert.Equal((-52.2, 48), (shapes[0].Left, shapes[1].Left));
        Assert.Equal((-42, -12), (shapes[0].Top, shapes[1].Top));
        Assert.True(shapes[1].ZOrderPosition > shapes[0].ZOrderPosition);
        Assert.NotEqual(shapes[0].ID, shapes[1].ID);
        Assert.Empty(DocumentIntegrity.CheckAll(_document.WordprocessingDocument));
        Assert.Throws<InvalidOperationException>(() => shapes[0].Ungroup());
    }

    [Fact]
    public void OnlySomeShapesCanBeGrouped()
    {
        ShapeWrapper a = Rectangle("A", 20, 30, 80, 40);
        ShapeWrapper line = _document.Shapes.AddLine(20, 90, 100, 130, _document.Range(5, 5));
        ShapeWrapper box = _document.Shapes.AddTextbox(TextOrientation.Horizontal, 120, 30, 60, 30, _document.Range(5, 5));
        ShapeWrapper group = _document.Shapes.Group(new[] { a, line, box });
        Assert.Equal(new[] { ShapeType.AutoShape, ShapeType.Line, ShapeType.TextBox }, group.GroupItems.Select(s => s.Type).ToArray());

        Assert.Throws<ArgumentException>(() => _document.Shapes.Group(new[] { group }));
        ShapeWrapper other = Rectangle("Other", 10, 10, 10, 10);
        Assert.Throws<NotSupportedException>(() => _document.Shapes.Group(new[] { group, other }));
        Assert.Throws<ArgumentNullException>(() => _document.Shapes.Group(null));
    }
}
