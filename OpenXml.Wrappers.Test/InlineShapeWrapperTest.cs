using System;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Inline pictures. The markup is Word 16's (the inline-picture-* scenarios of MarkupConformanceTest); the values of the members are those
/// Word gave for the same calls.
/// </summary>
public class InlineShapeWrapperTest : IDisposable
{
    private readonly string _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + ".docx");
    private readonly string _picture = System.IO.Path.GetFullPath(TestFile.Picture);
    private readonly System.Collections.Generic.List<string> _files = new System.Collections.Generic.List<string>();
    private readonly WordprocessingDocumentWrapper _document;

    public InlineShapeWrapperTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.InsertAfter("Hello world again");
    }

    public void Dispose()
    {
        _document.Dispose();
        System.IO.File.Delete(_path);
        foreach (string file in _files)
        {
            System.IO.File.Delete(file);
        }
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        System.Collections.Generic.IList<string> errors = DocumentIntegrity.CheckAll(_document.WordprocessingDocument);
        Assert.True(errors.Count == 0, string.Join("; ", errors));
    }

    private string WriteFile(string extension, params byte[] bytes)
    {
        string path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), System.IO.Path.GetRandomFileName() + extension);
        System.IO.File.WriteAllBytes(path, bytes);
        _files.Add(path);
        return path;
    }

    [Fact]
    public void APictureTakesTheSizeOfItsFileAndSitsInTheText()
    {
        InlineShapeWrapper shape = _document.InlineShapes.AddPicture(_picture, false, true, _document.Range(5, 5));

        Assert.Equal(1, _document.InlineShapes.Count);
        Assert.Equal(shape, _document.InlineShapes[0]);
        Assert.Equal(30, shape.Width, 2);
        Assert.Equal(15, shape.Height, 2);
        Assert.Equal(InlineShapeType.Picture, shape.Type);
        Assert.False(shape.LockAspectRatio);
        Assert.Equal(string.Empty, shape.AlternativeText);
        Assert.Equal(100, shape.ScaleWidth, 1);
        Assert.Equal(5, shape.Range.Start);
        Assert.Equal(6, shape.Range.End);
        Assert.Single(_document.Range(0, 11).InlineShapes);
        Assert.Empty(_document.Range(6, 11).InlineShapes);
        AssertValid();
    }

    [Fact]
    public void ThePictureGoesBeforeTheTextOfTheRange()
    {
        _document.InlineShapes.AddPicture(_picture, false, true, _document.Range(6, 11));

        Assert.Equal(6, _document.InlineShapes[0].Range.Start);
        Assert.Equal("Hello world again", _document.Content.Text.TrimEnd('\r').Replace("\u0001", string.Empty).Replace("\u0002", string.Empty));
    }

    [Fact]
    public void SizeAndScaleFollowTheAspectRatioLock()
    {
        InlineShapeWrapper shape = _document.InlineShapes.AddPicture(_picture, false, true, _document.Range(5, 5));

        shape.Width = 60;
        Assert.Equal(60, shape.Width, 2);
        Assert.Equal(15, shape.Height, 2);
        Assert.Equal(200, shape.ScaleWidth, 1);
        Assert.Equal(100, shape.ScaleHeight, 1);

        shape.LockAspectRatio = true;
        Assert.True(shape.LockAspectRatio);
        shape.Height = 30;
        Assert.Equal(30, shape.Height, 2);
        Assert.Equal(120, shape.Width, 1);

        shape.ScaleHeight = 100;
        Assert.Equal(15, shape.Height, 2);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.Width = -1);
        AssertValid();
    }

    [Fact]
    public void TheDescriptionAndTitleAreKept()
    {
        InlineShapeWrapper shape = _document.InlineShapes.AddPicture(_picture, false, true, _document.Range(5, 5));
        shape.AlternativeText = "A blue box";
        shape.Title = "Box";

        Assert.Equal("A blue box", shape.AlternativeText);
        Assert.Equal("Box", shape.Title);

        shape.AlternativeText = null;
        Assert.Equal(string.Empty, shape.AlternativeText);
        AssertValid();
    }

    [Fact]
    public void ThePictureFileIsSharedAndLeavesWithItsLastShape()
    {
        InlineShapeWrapper first = _document.InlineShapes.AddPicture(_picture, false, true, _document.Range(5, 5));
        InlineShapeWrapper second = _document.InlineShapes.AddPicture(_picture, false, true, _document.Range(0, 0));
        Func<int> parts = () => _document.WordprocessingDocument.MainDocumentPart.ImageParts.Count();

        Assert.Equal(2, _document.InlineShapes.Count);
        Assert.Equal(1, parts());
        Assert.Equal("Picture 2", second.Element.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().First().Name.Value);

        first.Delete();
        Assert.Equal(1, _document.InlineShapes.Count);
        Assert.Equal(1, parts());
        AssertValid();

        second.Delete();
        Assert.Equal(0, _document.InlineShapes.Count);
        Assert.Equal(0, parts());
        Assert.Equal("Hello world again", _document.Content.Text.TrimEnd('\r'));
        AssertValid();
    }

    [Fact]
    public void ALinkedPictureIsAnExternalRelationship()
    {
        InlineShapeWrapper shape = _document.InlineShapes.AddPicture(_picture, true, false, _document.Range(5, 5));

        Assert.Equal(InlineShapeType.LinkedPicture, shape.Type);
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.ImageParts);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.ExternalRelationships);
        Assert.Equal(100, shape.ScaleWidth, 1);

        shape.Delete();
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.ExternalRelationships);
    }

    [Fact]
    public void AHeaderCanHoldAPicture()
    {
        HeaderFooterWrapper header = _document.Sections[0].Headers[HeaderFooterIndex.Primary];
        header.Range.InsertAfter("Header");
        InlineShapeWrapper shape = header.Range.InlineShapes.AddPicture(_picture, false, true, header.Range.Duplicate());

        Assert.Equal(1, header.Range.InlineShapes.Count);
        Assert.Equal(0, _document.InlineShapes.Count);
        Assert.Single(_document.WordprocessingDocument.MainDocumentPart.HeaderParts.First().ImageParts);
        shape.Delete();
        Assert.Empty(_document.WordprocessingDocument.MainDocumentPart.HeaderParts.First().ImageParts);
        AssertValid();
    }

    [Fact]
    public void JpegGifAndBmpFilesAreMeasured()
    {
        //A JPEG: SOI, JFIF header at 72 dpi, a baseline frame header of 144 x 72 pixels.
        string jpeg = WriteFile(".jpg", 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, (byte)'J', (byte)'F', (byte)'I', (byte)'F', 0x00, 0x01, 0x01, 0x01, 0x00, 0x48, 0x00, 0x48, 0x00, 0x00,
            0xFF, 0xC0, 0x00, 0x0B, 0x08, 0x00, 0x48, 0x00, 0x90, 0x01, 0x01, 0x11, 0x00);
        //A GIF of 10 x 5 pixels.
        string gif = WriteFile(".gif", (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 10, 0, 5, 0, 0, 0, 0);
        //A BMP of 20 x 10 pixels at 3780 pixels to the metre.
        byte[] bmp = new byte[54];
        bmp[0] = (byte)'B';
        bmp[1] = (byte)'M';
        bmp[18] = 20;
        bmp[22] = 10;
        bmp[38] = 0xC4;
        bmp[39] = 0x0E;
        bmp[42] = 0xC4;
        bmp[43] = 0x0E;
        string bitmap = WriteFile(".bmp", bmp);

        InlineShapeWrapper a = _document.InlineShapes.AddPicture(jpeg, false, true, _document.Range(0, 0));
        InlineShapeWrapper b = _document.InlineShapes.AddPicture(gif, false, true, _document.Range(0, 0));
        InlineShapeWrapper c = _document.InlineShapes.AddPicture(bitmap, false, true, _document.Range(0, 0));

        Assert.Equal(144, a.Width, 1);
        Assert.Equal(72, a.Height, 1);
        Assert.Equal(7.5, b.Width, 1);
        Assert.Equal(3.75, b.Height, 1);
        Assert.Equal(15, c.Width, 1);
        Assert.Equal(7.5, c.Height, 1);
    }

    [Fact]
    public void AFileThatIsNotAPictureIsRefused()
    {
        string text = WriteFile(".txt", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12);

        Assert.Throws<ArgumentException>(() => _document.InlineShapes.AddPicture(text, false, true, _document.Range(0, 0)));
        Assert.Throws<ArgumentNullException>(() => _document.InlineShapes.AddPicture(null));
        Assert.Throws<ArgumentOutOfRangeException>(() => _document.InlineShapes[0]);
        Assert.Equal(0, _document.InlineShapes.Count);
    }
}
