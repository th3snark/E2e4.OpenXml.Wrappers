using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using WpDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What an inline shape is, as <c>WdInlineShapeType</c> in Word (with the same values; only the kinds that can be told from the markup).
/// </summary>
public enum InlineShapeType
{
    /// <summary>A picture stored in the document.</summary>
    Picture = 3,

    /// <summary>A picture linked to a file.</summary>
    LinkedPicture = 4,

    /// <summary>A chart.</summary>
    Chart = 12,

    /// <summary>A locked canvas.</summary>
    LockedCanvas = 14,

    /// <summary>A SmartArt graphic.</summary>
    SmartArt = 15
}

/// <summary>
/// An inline shape, as <c>InlineShape</c> in Word: a picture (or other graphic) that sits in the text like a character.
/// </summary>
public class InlineShapeWrapper : IEquatable<InlineShapeWrapper>
{
    private const double EmuPerPoint = 12700;
    private const string PictureUri = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private const string ImageRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    private readonly Drawing _drawing;

    internal InlineShapeWrapper(Drawing drawing)
    {
        _drawing = drawing;
    }

    internal Drawing Element
    {
        get { return _drawing; }
    }

    private WpDrawing.Inline Inline
    {
        get { return _drawing.GetFirstChild<WpDrawing.Inline>(); }
    }

    private A.Blip Blip
    {
        get { return Inline.Descendants<A.Blip>().FirstOrDefault(); }
    }

    /// <summary>What the shape is, as <c>InlineShape.Type</c> in Word.</summary>
    public InlineShapeType Type
    {
        get
        {
            string uri = Inline.Descendants<A.GraphicData>().FirstOrDefault()?.Uri?.Value ?? string.Empty;
            if (uri.EndsWith("/chart", StringComparison.Ordinal))
            {
                return InlineShapeType.Chart;
            }
            if (uri.EndsWith("/diagram", StringComparison.Ordinal))
            {
                return InlineShapeType.SmartArt;
            }
            if (uri.EndsWith("/lockedCanvas", StringComparison.Ordinal))
            {
                return InlineShapeType.LockedCanvas;
            }
            return Blip?.Link?.Value != null ? InlineShapeType.LinkedPicture : InlineShapeType.Picture;
        }
    }

    private long Cx
    {
        get { return Inline.Extent?.Cx?.Value ?? 0; }
    }

    private long Cy
    {
        get { return Inline.Extent?.Cy?.Value ?? 0; }
    }

    private void SetSize(long cx, long cy)
    {
        Inline.Extent.Cx = cx;
        Inline.Extent.Cy = cy;
        foreach (A.Extents extents in Inline.Descendants<A.Extents>())
        {
            extents.Cx = cx;
            extents.Cy = cy;
        }
    }

    private static long ToEmu(double points)
    {
        if (points < 0 || double.IsNaN(points) || double.IsInfinity(points))
        {
            throw new ArgumentOutOfRangeException(nameof(points));
        }
        return (long)Math.Round(points * EmuPerPoint);
    }

    /// <summary>The width in points, as <c>InlineShape.Width</c> in Word. With <see cref="LockAspectRatio"/> the height follows.</summary>
    public double Width
    {
        get { return Cx / EmuPerPoint; }
        set
        {
            long cx = ToEmu(value);
            SetSize(cx, LockAspectRatio && Cx > 0 ? (long)Math.Round((double)Cy * cx / Cx) : Cy);
        }
    }

    /// <summary>The height in points, as <c>InlineShape.Height</c> in Word. With <see cref="LockAspectRatio"/> the width follows.</summary>
    public double Height
    {
        get { return Cy / EmuPerPoint; }
        set
        {
            long cy = ToEmu(value);
            SetSize(LockAspectRatio && Cy > 0 ? (long)Math.Round((double)Cx * cy / Cy) : Cx, cy);
        }
    }

    /// <summary>True if changing the width or the height changes the other to keep the shape's proportions, as <c>InlineShape.LockAspectRatio</c> in Word.</summary>
    public bool LockAspectRatio
    {
        get { return Inline.NonVisualGraphicFrameDrawingProperties?.GetFirstChild<A.GraphicFrameLocks>()?.NoChangeAspect?.Value == true; }
        set
        {
            WpDrawing.NonVisualGraphicFrameDrawingProperties properties = Inline.NonVisualGraphicFrameDrawingProperties;
            if (properties == null)
            {
                properties = new WpDrawing.NonVisualGraphicFrameDrawingProperties();
                Inline.DocProperties.InsertAfterSelf(properties);
            }
            //Word writes the element even when the lock is off.
            A.GraphicFrameLocks locks = properties.GetFirstChild<A.GraphicFrameLocks>();
            if (locks == null)
            {
                locks = new A.GraphicFrameLocks();
                properties.AppendChild(locks);
            }
            locks.NoChangeAspect = value ? new BooleanValue(true) : null;
        }
    }

    //The picture at its own size, or null when the file is not available (a linked picture whose file has gone).
    private PictureFile Original()
    {
        A.Blip blip = Blip;
        OpenXmlPartContainer container = _drawing.GetRoot() is OpenXmlPartRootElement root ? root.OpenXmlPart : null;
        if (blip == null || container == null)
        {
            return null;
        }
        try
        {
            if (blip.Embed?.Value != null)
            {
                using (Stream stream = container.GetPartById(blip.Embed.Value).GetStream())
                using (MemoryStream copy = new MemoryStream())
                {
                    stream.CopyTo(copy);
                    return PictureFile.Read(copy.ToArray());
                }
            }
            string link = blip.Link?.Value;
            Uri target = link == null ? null : container.ExternalRelationships.FirstOrDefault(r => r.Id == link)?.Uri;
            return target != null && target.IsFile && File.Exists(target.LocalPath) ? PictureFile.Read(File.ReadAllBytes(target.LocalPath)) : null;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    /// <summary>The width as a percentage of the picture's own width, as <c>InlineShape.ScaleWidth</c> in Word; 100 if the picture cannot be read.</summary>
    public double ScaleWidth
    {
        get
        {
            PictureFile original = Original();
            return original == null || original.WidthEmu == 0 ? 100 : 100.0 * Cx / original.WidthEmu;
        }
        set
        {
            PictureFile original = Original() ?? throw new InvalidOperationException("The picture is not available to measure.");
            SetSize((long)Math.Round(original.WidthEmu * Math.Max(0, value) / 100.0), Cy);
        }
    }

    /// <summary>The height as a percentage of the picture's own height, as <c>InlineShape.ScaleHeight</c> in Word; 100 if the picture cannot be read.</summary>
    public double ScaleHeight
    {
        get
        {
            PictureFile original = Original();
            return original == null || original.HeightEmu == 0 ? 100 : 100.0 * Cy / original.HeightEmu;
        }
        set
        {
            PictureFile original = Original() ?? throw new InvalidOperationException("The picture is not available to measure.");
            SetSize(Cx, (long)Math.Round(original.HeightEmu * Math.Max(0, value) / 100.0));
        }
    }

    //An attribute with no text is left out, not written empty.
    private static StringValue Text(string value)
    {
        return string.IsNullOrEmpty(value) ? null : new StringValue(value);
    }

    //Once the description or title is set, Word gives the picture itself the same name, description and title as the shape.
    private void SyncPictureName()
    {
        Pic.NonVisualDrawingProperties picture = Inline.Descendants<Pic.NonVisualDrawingProperties>().FirstOrDefault();
        if (picture == null)
        {
            return;
        }
        WpDrawing.DocProperties shape = Inline.DocProperties;
        picture.Name = shape.Name?.Value;
        picture.Description = Text(shape.Description?.Value);
        picture.Title = Text(shape.Title?.Value);
    }

    /// <summary>The text a screen reader says for the shape, as <c>InlineShape.AlternativeText</c> in Word.</summary>
    public string AlternativeText
    {
        get { return Inline.DocProperties?.Description?.Value ?? string.Empty; }
        set
        {
            Inline.DocProperties.Description = Text(value);
            SyncPictureName();
        }
    }

    /// <summary>The title of the shape, as <c>InlineShape.Title</c> in Word.</summary>
    public string Title
    {
        get { return Inline.DocProperties?.Title?.Value ?? string.Empty; }
        set
        {
            Inline.DocProperties.Title = Text(value);
            SyncPictureName();
        }
    }

    /// <summary>The place the shape takes in the text (one position), as <c>InlineShape.Range</c> in Word.</summary>
    public RangeWrapper Range
    {
        get
        {
            StoryMap map = StoryMap.Build(StoryMap.FindRoot(_drawing));
            Atom atom = map.Atoms.First(a => a.Kind == AtomKind.Object && (ReferenceEquals(a.Element, _drawing) || a.Element.Descendants<Drawing>().Contains(_drawing)));
            return RangeWrapper.FromMap(map, atom.Start, atom.End);
        }
    }

    /// <summary>
    /// Makes the shape a floating one, anchored where it is, in front of the text, as <c>InlineShape.ConvertToShape</c> in Word. Without
    /// layout its position is not known, so it is put at the anchor (Word keeps the place it had on the page).
    /// </summary>
    public ShapeWrapper ConvertToShape()
    {
        Drawing anchored = ShapeWrapper.NewAnchored(_drawing, 0, 0, ShapeWrapperCollection.NextRelativeHeight(_drawing), true);
        _drawing.InsertBeforeSelf(anchored);
        _drawing.Remove();
        return new ShapeWrapper(anchored);
    }

    /// <summary>
    /// Takes the shape out of the text, as <c>InlineShape.Delete</c> in Word. A picture file nothing uses any more goes out of the package.
    /// </summary>
    public void Delete()
    {
        A.Blip blip = Blip;
        string embed = blip?.Embed?.Value;
        string link = blip?.Link?.Value;
        OpenXmlElement root = _drawing.GetRoot();
        OpenXmlPartContainer container = root is OpenXmlPartRootElement partRoot ? partRoot.OpenXmlPart : null;

        Range.Delete();

        PictureFile.Release(root, container, embed, link);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as InlineShapeWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(InlineShapeWrapper other)
    {
        return !(other is null) && ReferenceEquals(_drawing, other._drawing);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _drawing.GetHashCode();
    }

    public static bool operator ==(InlineShapeWrapper left, InlineShapeWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(InlineShapeWrapper left, InlineShapeWrapper right)
    {
        return !(left == right);
    }

    internal static Drawing NewPicture(OpenXmlPartContainer container, MainDocumentPart main, PictureFile file, string fileName, bool link)
    {
        string id;
        if (link)
        {
            id = container.AddExternalRelationship(ImageRelationship, new Uri("file:///" + Path.GetFullPath(fileName))).Id;
        }
        else
        {
            ImagePart part = container.Parts.Select(p => p.OpenXmlPart).OfType<ImagePart>().FirstOrDefault(p => p.ContentType == file.ContentType && SameBytes(p, file.Bytes));
            if (part == null)
            {
                part = PictureFile.AddPart(container, file.ContentType);
                using (MemoryStream stream = new MemoryStream(file.Bytes))
                {
                    part.FeedData(stream);
                }
            }
            id = container.GetIdOfPart(part);
        }

        uint docPrId = main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<WpDrawing.DocProperties>())
            .Select(d => d.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
        A.Blip blip = new A.Blip();
        if (link)
        {
            blip.Link = id;
        }
        else
        {
            blip.Embed = id;
            blip.AppendChild(new A.BlipExtensionList(new A.BlipExtension(new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false })
            {
                Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
            }));
        }

        WpDrawing.Inline inline = new WpDrawing.Inline() { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };
        inline.AppendChild(new WpDrawing.Extent() { Cx = file.WidthEmu, Cy = file.HeightEmu });
        inline.AppendChild(new WpDrawing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
        inline.AppendChild(new WpDrawing.DocProperties() { Id = docPrId, Name = "Picture " + docPrId.ToString(System.Globalization.CultureInfo.InvariantCulture) });
        inline.AppendChild(new WpDrawing.NonVisualGraphicFrameDrawingProperties());
        Pic.Picture picture = new Pic.Picture(
            new Pic.NonVisualPictureProperties(
                new Pic.NonVisualDrawingProperties() { Id = (uint)(ListDefaults.NewRandom() & 0x7FFFFFFF), Name = string.Empty },
                new Pic.NonVisualPictureDrawingProperties()),
            new Pic.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
            new Pic.ShapeProperties(
                new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = file.WidthEmu, Cy = file.HeightEmu }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
        A.GraphicData graphicData = new A.GraphicData(picture) { Uri = PictureUri };
        inline.AppendChild(new A.Graphic(graphicData));
        return new Drawing(inline);
    }

    private static bool SameBytes(ImagePart part, byte[] bytes)
    {
        using (Stream stream = part.GetStream())
        {
            if (stream.Length != bytes.Length)
            {
                return false;
            }
            byte[] existing = new byte[bytes.Length];
            int read = 0;
            while (read < existing.Length)
            {
                int count = stream.Read(existing, read, existing.Length - read);
                if (count <= 0)
                {
                    return false;
                }
                read += count;
            }
            for (int i = 0; i < bytes.Length; i++)
            {
                if (existing[i] != bytes[i])
                {
                    return false;
                }
            }
            return true;
        }
    }
}

/// <summary>
/// The inline shapes of a range or a document, as <c>InlineShapes</c> in Word.
/// </summary>
public class InlineShapeWrapperCollection : IEnumerable<InlineShapeWrapper>
{
    private readonly RangeWrapper _range;

    internal InlineShapeWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the shapes are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<InlineShapeWrapper> GetItems()
    {
        OpenXmlElement root = StoryMap.FindRoot(_range.AnchorElement);
        StoryMap map = StoryMap.Build(root);
        _range.GetStoryPositions(out int start, out int end);
        List<InlineShapeWrapper> items = new List<InlineShapeWrapper>();
        foreach (Atom atom in map.Atoms)
        {
            if (atom.Kind != AtomKind.Object || atom.Start < start || atom.End > end)
            {
                continue;
            }
            Drawing drawing = atom.Element as Drawing ?? atom.Element.Descendants<Drawing>().FirstOrDefault();
            if (drawing?.GetFirstChild<WpDrawing.Inline>() != null)
            {
                items.Add(new InlineShapeWrapper(drawing));
            }
        }
        return items;
    }

    /// <summary>The number of inline shapes.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The shape at the index (0-based), as <c>InlineShapes.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no shape at the index.</exception>
    public InlineShapeWrapper this[int index]
    {
        get
        {
            List<InlineShapeWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Inserts a picture from a PNG, JPEG, GIF or BMP file, as <c>InlineShapes.AddPicture</c> in Word. The picture goes at the start of
    /// the range (the text of the range stays) and has the size its file gives it. A picture already in the document is shared.
    /// </summary>
    /// <param name="fileName">The picture file.</param>
    /// <param name="linkToFile">True to link to the file instead of storing the picture in the document.</param>
    /// <param name="saveWithDocument">Whether to store the picture in the document; a linked picture is only linked.</param>
    /// <param name="range">Where to put the picture; the start of the range this collection is for when null.</param>
    /// <exception cref="ArgumentException">The file is not a PNG, JPEG, GIF or BMP picture.</exception>
    /// <exception cref="NotSupportedException">The place is not in the main text, a header or a footer.</exception>
    public InlineShapeWrapper AddPicture(string fileName, bool linkToFile = false, bool saveWithDocument = true, RangeWrapper range = null)
    {
        Guard.ThrowIfNull(fileName);
        PictureFile file = PictureFile.Read(File.ReadAllBytes(fileName));
        RangeWrapper at = (range ?? _range).Duplicate();
        at.Collapse(CollapseDirection.Start);

        Drawing drawing = null;
        at.InsertNoteReference(paragraph =>
        {
            MainDocumentPart main = StyleHelper.GetMainDocumentPart(paragraph);
            OpenXmlPartContainer container = paragraph.GetRoot() is OpenXmlPartRootElement root ? root.OpenXmlPart : null;
            if (main == null || container == null)
            {
                throw new NotSupportedException("A picture can be inserted only in a document.");
            }
            drawing = InlineShapeWrapper.NewPicture(container, main, file, fileName, linkToFile);
            return new List<Run> { new Run(new RunProperties(new NoProof()), drawing) };
        });
        return new InlineShapeWrapper(drawing);
    }

    /// <inheritdoc/>
    public IEnumerator<InlineShapeWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
