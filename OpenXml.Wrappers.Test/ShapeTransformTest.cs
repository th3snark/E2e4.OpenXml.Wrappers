using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Flipping, moving, scaling and duplicating shapes, and the text frame settings. The expected values are those Word 16 gave for the same
/// calls (the shape-flip, shape-increments, shape-scale, shape-textframe, shape-lines-directions, shape-duplicate and shape-layout-in-cell
/// scenarios of tools/WordReference/scenarios.ps1, which print them).
/// </summary>
public class ShapeTransformTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public ShapeTransformTest()
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

    private ShapeWrapper Rectangle(double left = 20, double top = 30)
    {
        return _document.Shapes.AddShape(AutoShapeType.Rectangle, left, top, 80, 40, _document.Range(5, 5));
    }

    [Fact]
    public void FlippingTurnsTheRotationAround()
    {
        ShapeWrapper shape = Rectangle();
        Assert.False(shape.HorizontalFlip);
        shape.Flip(FlipCommand.Horizontal);
        Assert.True(shape.HorizontalFlip);
        shape.Flip(FlipCommand.Horizontal);
        Assert.False(shape.HorizontalFlip);

        shape.Rotation = 30;
        shape.Flip(FlipCommand.Horizontal);
        Assert.True(shape.HorizontalFlip);
        Assert.Equal(330, shape.Rotation, 3);

        ShapeWrapper other = Rectangle();
        other.Rotation = 100;
        other.Flip(FlipCommand.Horizontal);
        other.Flip(FlipCommand.Vertical);
        Assert.True(other.HorizontalFlip);
        Assert.True(other.VerticalFlip);
        Assert.Equal(100, other.Rotation, 3);
        Assert.Throws<ArgumentOutOfRangeException>(() => other.Flip((FlipCommand)5));
    }

    [Fact]
    public void IncrementsMoveAndTurn()
    {
        ShapeWrapper shape = Rectangle();
        double left = shape.Left;
        double top = shape.Top;
        shape.IncrementLeft(10);
        shape.IncrementTop(-5);
        shape.IncrementRotation(30);
        shape.IncrementRotation(350);
        Assert.Equal(left + 10, shape.Left, 3);
        Assert.Equal(top - 5, shape.Top, 3);
        Assert.Equal(20, shape.Rotation, 3);
    }

    [Fact]
    public void ScalingKeepsTheNamedPoint()
    {
        ShapeWrapper a = Rectangle();
        double left = a.Left;
        double top = a.Top;
        a.ScaleWidth(2);
        a.ScaleHeight(0.5);
        Assert.Equal((160, 20), (a.Width, a.Height));
        Assert.Equal((left, top), (a.Left, a.Top));

        ShapeWrapper b = Rectangle(20, 100);
        top = b.Top;
        b.ScaleHeight(2, false, ScaleFrom.BottomRight);
        Assert.Equal(80, b.Height, 3);
        Assert.Equal(top - 40, b.Top, 3);

        ShapeWrapper c = Rectangle(20, 170);
        left = c.Left;
        c.ScaleWidth(1.5, false, ScaleFrom.Middle);
        Assert.Equal(120, c.Width, 3);
        Assert.Equal(left - 20, c.Left, 3);

        Assert.Throws<NotSupportedException>(() => a.ScaleWidth(2, true));
        Assert.Throws<ArgumentOutOfRangeException>(() => a.ScaleWidth(0));
    }

    [Fact]
    public void LinesThatGoUpOrLeftAreTurnedOver()
    {
        ShapeWrapper a = _document.Shapes.AddLine(100, 70, 20, 30, _document.Range(5, 5));
        ShapeWrapper b = _document.Shapes.AddLine(20, 70, 100, 30, _document.Range(5, 5));
        ShapeWrapper c = _document.Shapes.AddLine(100, 30, 20, 70, _document.Range(5, 5));
        Assert.Equal((true, true), (a.HorizontalFlip, a.VerticalFlip));
        Assert.Equal((false, true), (b.HorizontalFlip, b.VerticalFlip));
        Assert.Equal((true, false), (c.HorizontalFlip, c.VerticalFlip));
        Assert.All(new[] { a, b, c }, s => Assert.Equal((80, 40), (s.Width, s.Height)));
        Assert.Equal(a.Left, b.Left);
        Assert.Equal(a.Top, c.Top);
    }

    [Fact]
    public void TextFrameSettings()
    {
        ShapeWrapper box = _document.Shapes.AddTextbox(TextOrientation.Horizontal, 20, 30, 120, 60, _document.Range(5, 5));
        TextFrameWrapper frame = box.TextFrame;
        Assert.Equal((7.2, 7.2, 3.6, 3.6), (frame.MarginLeft, frame.MarginRight, frame.MarginTop, frame.MarginBottom));
        Assert.Equal((VerticalAnchor.Top, HorizontalAnchor.None, true, false, TextOrientation.Horizontal),
            (frame.VerticalAnchor, frame.HorizontalAnchor, frame.WordWrap, frame.AutoSize, frame.Orientation));

        frame.MarginLeft = 10;
        frame.MarginRight = 12;
        frame.MarginTop = 3;
        frame.MarginBottom = 4;
        frame.VerticalAnchor = VerticalAnchor.Middle;
        frame.WordWrap = false;
        frame.AutoSize = true;
        frame.HorizontalAnchor = HorizontalAnchor.Center;
        frame.Orientation = TextOrientation.Vertical;
        Assert.Equal((10, 12, 3, 4), (frame.MarginLeft, frame.MarginRight, frame.MarginTop, frame.MarginBottom));
        Assert.Equal((VerticalAnchor.Middle, HorizontalAnchor.Center, false, true, TextOrientation.Vertical),
            (frame.VerticalAnchor, frame.HorizontalAnchor, frame.WordWrap, frame.AutoSize, frame.Orientation));
        Assert.Same(box, frame.Parent);

        //The shapes of the gallery start with their text in the middle.
        ShapeWrapper rectangle = Rectangle();
        Assert.Equal(VerticalAnchor.Middle, rectangle.TextFrame.VerticalAnchor);

        frame.TextRange.Text = "Boxed";
        Assert.True(frame.HasText);
        frame.DeleteText();
        Assert.False(frame.HasText);
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.MarginLeft = -1);
        Assert.Throws<ArgumentOutOfRangeException>(() => frame.Orientation = (TextOrientation)9);
    }

    [Fact]
    public void ALabelIsATextBox()
    {
        ShapeWrapper label = _document.Shapes.AddLabel(TextOrientation.Horizontal, 20, 30, 120, 40, _document.Range(5, 5));
        Assert.Equal(ShapeType.TextBox, label.Type);
        Assert.False(label.Fill.Visible);
        Assert.True(label.Line.Visible);

        ShapeWrapper vertical = _document.Shapes.AddTextbox(TextOrientation.Vertical, 20, 30, 120, 60, _document.Range(5, 5));
        Assert.Equal(TextOrientation.Vertical, vertical.TextFrame.Orientation);
    }

    [Fact]
    public void ADuplicateIsInFrontOfTheShape()
    {
        ShapeWrapper shape = Rectangle();
        shape.Fill.ForeColor.RGB = 255;
        ShapeWrapper copy = shape.Duplicate();
        Assert.Equal(2, _document.Shapes.Count);
        Assert.Equal(shape.Name, copy.Name);
        Assert.NotEqual(shape.ID, copy.ID);
        Assert.Equal((shape.Left, shape.Top, shape.Width, shape.Height), (copy.Left, copy.Top, copy.Width, copy.Height));
        Assert.Equal(255, copy.Fill.ForeColor.RGB);
        Assert.True(copy.ZOrderPosition > shape.ZOrderPosition);
        Assert.Equal(2, _document.Shapes.Select(s => s.ID).Distinct().Count());
    }

    [Fact]
    public void LayoutInCellAndDecorative()
    {
        ShapeWrapper shape = Rectangle();
        Assert.True(shape.LayoutInCell);
        Assert.False(shape.Decorative);
        shape.LayoutInCell = false;
        shape.Decorative = true;
        Assert.False(shape.LayoutInCell);
        Assert.True(shape.Decorative);
        shape.Decorative = false;
        Assert.False(shape.Decorative);
    }

    [Fact]
    public void TheMarkupIsValid()
    {
        ShapeWrapper rectangle = Rectangle();
        rectangle.Flip(FlipCommand.Horizontal);
        rectangle.ScaleWidth(2, false, ScaleFrom.Middle);
        rectangle.TextFrame.TextRange.Text = "Rect";
        rectangle.TextFrame.Orientation = TextOrientation.Upward;
        rectangle.Decorative = true;
        rectangle.Duplicate();
        ShapeWrapper box = _document.Shapes.AddTextbox(TextOrientation.Vertical, 20, 30, 120, 60, _document.Range(5, 5));
        box.TextFrame.AutoSize = true;
        box.TextFrame.MarginLeft = 10;
        _document.Shapes.AddLine(100, 70, 20, 30, _document.Range(5, 5));
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }
}
