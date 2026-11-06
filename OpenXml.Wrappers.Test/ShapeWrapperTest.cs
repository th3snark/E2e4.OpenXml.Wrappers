using System;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Floating pictures. The markup is Word 16's (the shape-* scenarios of MarkupConformanceTest); the values of the members are those Word
/// gave for the same calls.
/// </summary>
public class ShapeWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly string _picture = System.IO.Path.GetFullPath(TestFile.Picture);
    private readonly WordprocessingDocumentWrapper _document;

    public ShapeWrapperTest()
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

    private ShapeWrapper Add(double left = 20, double top = 30)
    {
        return _document.Shapes.AddPicture(_picture, false, true, left, top, 60, 30, _document.Range(5, 5));
    }

    [Fact]
    public void APictureFloatsWithTheDefaultsWordGives()
    {
        ShapeWrapper shape = Add();

        Assert.Equal(1, _document.Shapes.Count);
        Assert.Equal(shape, _document.Shapes[0]);
        Assert.Equal(ShapeType.Picture, shape.Type);
        Assert.Equal("Picture 1", shape.Name);
        Assert.Equal(20, shape.Left, 3);
        Assert.Equal(30, shape.Top, 3);
        Assert.Equal(60, shape.Width, 3);
        Assert.Equal(30, shape.Height, 3);
        Assert.Equal(RelativeHorizontalPosition.Column, shape.RelativeHorizontalPosition);
        Assert.Equal(RelativeVerticalPosition.Paragraph, shape.RelativeVerticalPosition);
        Assert.Equal(WrapType.None, shape.WrapFormat.Type);
        Assert.Equal(WrapSideType.Both, shape.WrapFormat.Side);
        Assert.Equal(9, shape.WrapFormat.DistanceLeft, 3);
        Assert.Equal(0, shape.WrapFormat.DistanceTop, 3);
        Assert.True(shape.WrapFormat.AllowOverlap);
        Assert.False(shape.LockAnchor);
        Assert.False(shape.LockAspectRatio);
        Assert.True(shape.Visible);
        Assert.Equal(0, shape.Rotation, 3);
        Assert.Equal(1, shape.ZOrderPosition);
        Assert.Equal(0, shape.Anchor.Start);
        Assert.Equal(0, _document.InlineShapes.Count);
        AssertValid();
    }

    [Fact]
    public void ThePictureTakesItsOwnSizeWhenNoneIsGiven()
    {
        ShapeWrapper shape = _document.Shapes.AddPicture(_picture, false, true, 0, 0, -1, -1, _document.Range(5, 5));

        Assert.Equal(30, shape.Width, 3);
        Assert.Equal(15, shape.Height, 3);
    }

    [Fact]
    public void TheStackFollowsTheOrderTheShapesWereAddedInUntilTheyAreMoved()
    {
        ShapeWrapper a = Add();
        ShapeWrapper b = Add(10, 10);
        ShapeWrapper c = Add(5, 5);

        Assert.Equal(new[] { a, b, c }, _document.Shapes.ToArray());
        Assert.Equal(new[] { "Picture 1", "Picture 2", "Picture 3" }, _document.Shapes.Select(s => s.Name).ToArray());

        c.ZOrder(ZOrderCommand.SendToBack);
        Assert.Equal(new[] { c, a, b }, _document.Shapes.ToArray());
        Assert.Equal(1, c.ZOrderPosition);

        c.ZOrder(ZOrderCommand.BringForward);
        Assert.Equal(new[] { a, c, b }, _document.Shapes.ToArray());

        a.ZOrder(ZOrderCommand.BringToFront);
        Assert.Equal(new[] { c, b, a }, _document.Shapes.ToArray());

        a.ZOrder(ZOrderCommand.SendBackward);
        Assert.Equal(new[] { c, a, b }, _document.Shapes.ToArray());
        AssertValid();
    }

    [Fact]
    public void WrappingChangesWhatTheMarkupSaysAndReadsBack()
    {
        ShapeWrapper shape = Add();

        foreach (WrapType type in new[] { WrapType.Square, WrapType.Tight, WrapType.Through, WrapType.TopBottom, WrapType.Behind, WrapType.None })
        {
            shape.WrapFormat.Type = type;
            Assert.Equal(type, shape.WrapFormat.Type);
        }

        shape.WrapFormat.Type = WrapType.Square;
        shape.WrapFormat.Side = WrapSideType.Left;
        Assert.Equal(WrapSideType.Left, shape.WrapFormat.Side);
        shape.WrapFormat.Side = WrapSideType.Largest;
        Assert.Equal(WrapSideType.Largest, shape.WrapFormat.Side);
        shape.WrapFormat.DistanceRight = 4.5;
        Assert.Equal(4.5, shape.WrapFormat.DistanceRight, 3);

        Assert.Throws<NotSupportedException>(() => shape.WrapFormat.Type = WrapType.Inline);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.WrapFormat.Type = (WrapType)99);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.WrapFormat.DistanceTop = -1);
        AssertValid();
    }

    [Fact]
    public void BehindTextAndInFrontOfTextAreSetByZOrder()
    {
        ShapeWrapper shape = Add();

        shape.ZOrder(ZOrderCommand.SendBehindText);
        Assert.Equal(WrapType.Behind, shape.WrapFormat.Type);
        shape.ZOrder(ZOrderCommand.BringInFrontOfText);
        Assert.Equal(WrapType.None, shape.WrapFormat.Type);
    }

    [Fact]
    public void PositionSizeRotationAndLocksAreSetAndRead()
    {
        ShapeWrapper shape = Add();

        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
        shape.Left = -12.5;
        shape.Top = 50;
        shape.Rotation = 390;
        shape.LockAspectRatio = true;
        shape.Width = 120;
        shape.LockAnchor = true;
        shape.Visible = false;
        shape.Name = "Box";
        shape.AlternativeText = "A box";
        shape.Title = "Title";

        Assert.Equal(RelativeHorizontalPosition.Page, shape.RelativeHorizontalPosition);
        Assert.Equal(RelativeVerticalPosition.Margin, shape.RelativeVerticalPosition);
        Assert.Equal(-12.5, shape.Left, 3);
        Assert.Equal(50, shape.Top, 3);
        Assert.Equal(30, shape.Rotation, 3);
        Assert.Equal(120, shape.Width, 3);
        Assert.Equal(60, shape.Height, 3);
        Assert.True(shape.LockAnchor);
        Assert.False(shape.Visible);
        Assert.Equal("Box", shape.Name);
        Assert.Equal("A box", shape.AlternativeText);
        Assert.Equal("Title", shape.Title);

        shape.Rotation = 0;
        Assert.Equal(0, shape.Rotation, 3);
        Assert.Throws<ArgumentNullException>(() => shape.Name = null);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.Width = -2);
        AssertValid();
    }

    [Fact]
    public void DeleteTakesTheShapeAndItsPictureFileAway()
    {
        ShapeWrapper shape = Add();
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.ImageParts);

        shape.Delete();

        Assert.Equal(0, _document.Shapes.Count);
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.ImageParts);
        Assert.Equal("Hello world again", _document.Content.Text.TrimEnd('\r'));
        AssertValid();
    }

    [Fact]
    public void AShapeBecomesInlineAndBack()
    {
        ShapeWrapper shape = Add();

        InlineShapeWrapper inline = shape.ConvertToInlineShape();
        Assert.Equal(0, _document.Shapes.Count);
        Assert.Equal(1, _document.InlineShapes.Count);
        Assert.Equal(60, inline.Width, 3);

        ShapeWrapper again = inline.ConvertToShape();
        Assert.Equal(1, _document.Shapes.Count);
        Assert.Equal(0, _document.InlineShapes.Count);
        Assert.Equal(WrapType.None, again.WrapFormat.Type);
        Assert.Equal(60, again.Width, 3);
        AssertValid();
    }

    [Fact]
    public void ALinkedPictureIsALinkedShape()
    {
        ShapeWrapper shape = _document.Shapes.AddPicture(_picture, true, false, 0, 0, 30, 15, _document.Range(0, 0));

        Assert.Equal(ShapeType.LinkedPicture, shape.Type);
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.ImageParts);
        shape.Delete();
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.ExternalRelationships);
    }

    [Fact]
    public void AHeaderHasItsOwnShapes()
    {
        HeaderFooterWrapper header = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        header.Range.InsertAfter("Header");
        ShapeWrapper shape = header.Shapes.AddPicture(_picture, false, true, 0, 0, 30, 15, header.Range.Duplicate());

        Assert.Equal(1, header.Shapes.Count);
        Assert.Equal(0, _document.Shapes.Count);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.HeaderParts.First().ImageParts);
        shape.Delete();
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.HeaderParts.First().ImageParts);
        AssertValid();
    }

    [Fact]
    public void TheArgumentsAreChecked()
    {
        Assert.Throws<ArgumentNullException>(() => _document.Shapes.AddPicture(null));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.Shapes[0]);
    }
}
