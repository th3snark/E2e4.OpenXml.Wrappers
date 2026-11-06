using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using WpDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a shape is, as <c>MsoShapeType</c> in Office (with the same values; only the kinds this library makes).
/// </summary>
public enum ShapeType
{
    /// <summary>A shape of the gallery (see <see cref="AutoShapeType"/>).</summary>
    AutoShape = 1,

    /// <summary>A straight line.</summary>
    Line = 9,

    /// <summary>A linked picture.</summary>
    LinkedPicture = 11,

    /// <summary>A picture.</summary>
    Picture = 13,

    /// <summary>A text box.</summary>
    TextBox = 17
}

/// <summary>
/// What a shape is placed against horizontally, as <c>WdRelativeHorizontalPosition</c> in Word (with the same values).
/// </summary>
public enum RelativeHorizontalPosition
{
    /// <summary>The margin.</summary>
    Margin = 0,

    /// <summary>The edge of the page.</summary>
    Page = 1,

    /// <summary>The column.</summary>
    Column = 2,

    /// <summary>The character the shape is anchored at.</summary>
    Character = 3,

    /// <summary>The left margin area.</summary>
    LeftMarginArea = 4,

    /// <summary>The right margin area.</summary>
    RightMarginArea = 5,

    /// <summary>The inside margin area.</summary>
    InnerMarginArea = 6,

    /// <summary>The outside margin area.</summary>
    OuterMarginArea = 7
}

/// <summary>
/// What a shape is placed against vertically, as <c>WdRelativeVerticalPosition</c> in Word (with the same values).
/// </summary>
public enum RelativeVerticalPosition
{
    /// <summary>The margin.</summary>
    Margin = 0,

    /// <summary>The edge of the page.</summary>
    Page = 1,

    /// <summary>The paragraph the shape is anchored in.</summary>
    Paragraph = 2,

    /// <summary>The line the shape is anchored in.</summary>
    Line = 3,

    /// <summary>The top margin area.</summary>
    TopMarginArea = 4,

    /// <summary>The bottom margin area.</summary>
    BottomMarginArea = 5,

    /// <summary>The inside margin area.</summary>
    InnerMarginArea = 6,

    /// <summary>The outside margin area.</summary>
    OuterMarginArea = 7
}

/// <summary>
/// How text wraps around a shape, as <c>WdWrapType</c> in Word (with the same values).
/// </summary>
public enum WrapType
{
    /// <summary>Around the square the shape fits in.</summary>
    Square = 0,

    /// <summary>Around the outline of the shape.</summary>
    Tight = 1,

    /// <summary>Around the outline, and through open parts of the shape.</summary>
    Through = 2,

    /// <summary>The shape is in front of the text.</summary>
    None = 3,

    /// <summary>The text stops above the shape and goes on below it.</summary>
    TopBottom = 4,

    /// <summary>The shape is behind the text.</summary>
    Behind = 5,

    /// <summary>The shape is in the text like a character.</summary>
    Inline = 7
}

/// <summary>
/// Which sides of a shape text goes round, as <c>WdWrapSideType</c> in Word (with the same values).
/// </summary>
public enum WrapSideType
{
    /// <summary>Both sides.</summary>
    Both = 0,

    /// <summary>The left side only.</summary>
    Left = 1,

    /// <summary>The right side only.</summary>
    Right = 2,

    /// <summary>The wider side only.</summary>
    Largest = 3
}

/// <summary>
/// Where a shape goes in the stack, as <c>MsoZOrderCmd</c> in Office (with the same values).
/// </summary>
public enum ZOrderCommand
{
    /// <summary>On top of all the others.</summary>
    BringToFront = 0,

    /// <summary>Under all the others.</summary>
    SendToBack = 1,

    /// <summary>One place up.</summary>
    BringForward = 2,

    /// <summary>One place down.</summary>
    SendBackward = 3,

    /// <summary>In front of the text.</summary>
    BringInFrontOfText = 4,

    /// <summary>Behind the text.</summary>
    SendBehindText = 5
}

/// <summary>
/// How text wraps around a shape, as <c>WrapFormat</c> in Word.
/// </summary>
public class WrapFormatWrapper
{
    private const double EmuPerPoint = 12700;
    private readonly ShapeWrapper _shape;

    internal WrapFormatWrapper(ShapeWrapper shape)
    {
        _shape = shape;
    }

    private WpDrawing.Anchor Anchor
    {
        get { return _shape.AnchorElement; }
    }

    private static double Points(UInt32Value value)
    {
        return (value?.Value ?? 0U) / EmuPerPoint;
    }

    private static UInt32Value Emu(double points)
    {
        if (points < 0 || double.IsNaN(points) || double.IsInfinity(points))
        {
            throw new ArgumentOutOfRangeException(nameof(points));
        }
        return (uint)Math.Round(points * EmuPerPoint);
    }

    /// <summary>The distance in points between the shape and text above it, as <c>WrapFormat.DistanceTop</c> in Word.</summary>
    public double DistanceTop
    {
        get { return Points(Anchor.DistanceFromTop); }
        set { Anchor.DistanceFromTop = Emu(value); }
    }

    /// <summary>The distance in points between the shape and text below it, as <c>WrapFormat.DistanceBottom</c> in Word.</summary>
    public double DistanceBottom
    {
        get { return Points(Anchor.DistanceFromBottom); }
        set { Anchor.DistanceFromBottom = Emu(value); }
    }

    /// <summary>The distance in points between the shape and text on its left, as <c>WrapFormat.DistanceLeft</c> in Word.</summary>
    public double DistanceLeft
    {
        get { return Points(Anchor.DistanceFromLeft); }
        set { Anchor.DistanceFromLeft = Emu(value); }
    }

    /// <summary>The distance in points between the shape and text on its right, as <c>WrapFormat.DistanceRight</c> in Word.</summary>
    public double DistanceRight
    {
        get { return Points(Anchor.DistanceFromRight); }
        set { Anchor.DistanceFromRight = Emu(value); }
    }

    /// <summary>True if the shape may overlap other shapes, as <c>WrapFormat.AllowOverlap</c> in Word.</summary>
    public bool AllowOverlap
    {
        get { return Anchor.AllowOverlap?.Value != false; }
        set { Anchor.AllowOverlap = value; }
    }

    private OpenXmlElement WrapElement
    {
        get
        {
            return Anchor.ChildElements.FirstOrDefault(e => e is WpDrawing.WrapNone || e is WpDrawing.WrapSquare || e is WpDrawing.WrapTight
                || e is WpDrawing.WrapThrough || e is WpDrawing.WrapTopBottom);
        }
    }

    /// <summary>How text wraps around the shape, as <c>WrapFormat.Type</c> in Word.</summary>
    public WrapType Type
    {
        get
        {
            OpenXmlElement wrap = WrapElement;
            return wrap is WpDrawing.WrapSquare ? WrapType.Square
                : wrap is WpDrawing.WrapTight ? WrapType.Tight
                : wrap is WpDrawing.WrapThrough ? WrapType.Through
                : wrap is WpDrawing.WrapTopBottom ? WrapType.TopBottom
                : Anchor.BehindDoc?.Value == true ? WrapType.Behind : WrapType.None;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            if (value == WrapType.Inline)
            {
                throw new NotSupportedException("Use ConvertToInlineShape to put a shape in the text.");
            }
            WrapSideType side = Side;
            OpenXmlElement existing = WrapElement;
            OpenXmlElement wrap = NewWrap(value, side);
            existing.InsertBeforeSelf(wrap);
            existing.Remove();
            Anchor.BehindDoc = value == WrapType.Behind || value == WrapType.Tight;
        }
    }

    private static string WrapText(WrapSideType side)
    {
        return side == WrapSideType.Left ? "left" : side == WrapSideType.Right ? "right" : side == WrapSideType.Largest ? "largest" : "bothSides";
    }

    private static WpDrawing.WrapPolygon NewPolygon()
    {
        //Word's outline for a picture: its rectangle, a little short of the edges.
        WpDrawing.WrapPolygon polygon = new WpDrawing.WrapPolygon { Edited = false };
        polygon.AppendChild(new WpDrawing.StartPoint { X = 0L, Y = 0L });
        foreach ((long x, long y) in new[] { (0L, 20520L), (21060L, 20520L), (21060L, 0L), (0L, 0L) })
        {
            polygon.AppendChild(new WpDrawing.LineTo { X = x, Y = y });
        }
        return polygon;
    }

    private static OpenXmlElement NewWrap(WrapType type, WrapSideType side)
    {
        switch (type)
        {
            case WrapType.Square:
                return new WpDrawing.WrapSquare { WrapText = ToWrapText(side) };
            case WrapType.Tight:
                return new WpDrawing.WrapTight(NewPolygon()) { WrapText = ToWrapText(side) };
            case WrapType.Through:
                return new WpDrawing.WrapThrough(NewPolygon()) { WrapText = ToWrapText(side) };
            case WrapType.TopBottom:
                return new WpDrawing.WrapTopBottom();
            default:
                return new WpDrawing.WrapNone();
        }
    }

    private static WpDrawing.WrapTextValues ToWrapText(WrapSideType side)
    {
        return side == WrapSideType.Left ? WpDrawing.WrapTextValues.Left
            : side == WrapSideType.Right ? WpDrawing.WrapTextValues.Right
            : side == WrapSideType.Largest ? WpDrawing.WrapTextValues.Largest : WpDrawing.WrapTextValues.BothSides;
    }

    /// <summary>Which sides of the shape text goes round, as <c>WrapFormat.Side</c> in Word.</summary>
    public WrapSideType Side
    {
        get
        {
            OpenXmlElement wrap = WrapElement;
            string text = wrap is WpDrawing.WrapSquare square ? square.WrapText?.InnerText
                : wrap is WpDrawing.WrapTight tight ? tight.WrapText?.InnerText
                : wrap is WpDrawing.WrapThrough through ? through.WrapText?.InnerText : null;
            return text == "left" ? WrapSideType.Left : text == "right" ? WrapSideType.Right : text == "largest" ? WrapSideType.Largest : WrapSideType.Both;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            OpenXmlElement wrap = WrapElement;
            if (wrap is WpDrawing.WrapSquare square)
            {
                square.WrapText = ToWrapText(value);
            }
            else if (wrap is WpDrawing.WrapTight tight)
            {
                tight.WrapText = ToWrapText(value);
            }
            else if (wrap is WpDrawing.WrapThrough through)
            {
                through.WrapText = ToWrapText(value);
            }
        }
    }
}

/// <summary>
/// A floating shape, as <c>Shape</c> in Word: a picture that sits on the page, anchored to a paragraph, with text wrapping round it.
/// </summary>
public partial class ShapeWrapper : IEquatable<ShapeWrapper>
{
    private const double EmuPerPoint = 12700;
    private const long RotationUnit = 60000;
    private const long EffectExtentUnit = 19050;

    private readonly Drawing _drawing;

    internal ShapeWrapper(Drawing drawing)
    {
        _drawing = drawing;
    }

    internal Drawing Element
    {
        get { return _drawing; }
    }

    internal WpDrawing.Anchor AnchorElement
    {
        get { return _drawing.GetFirstChild<WpDrawing.Anchor>(); }
    }

    private WpDrawing.DocProperties DocProperties
    {
        get { return AnchorElement.GetFirstChild<WpDrawing.DocProperties>(); }
    }

    private A.Blip Blip
    {
        get { return AnchorElement.Descendants<A.Blip>().FirstOrDefault(); }
    }

    /// <summary>What the shape is, as <c>Shape.Type</c> in Word.</summary>
    public ShapeType Type
    {
        get { return Shape != null ? (IsTextBox ? ShapeType.TextBox : Shape.GetFirstChild<Wps.NonVisualConnectorProperties>() != null ? ShapeType.Line : ShapeType.AutoShape) : Blip?.Link?.Value != null ? ShapeType.LinkedPicture : ShapeType.Picture; }
    }

    /// <summary>The name of the shape, as <c>Shape.Name</c> in Word.</summary>
    public string Name
    {
        get { return DocProperties.Name?.Value ?? string.Empty; }
        set
        {
            Guard.ThrowIfNull(value);
            DocProperties.Name = value;
            SyncPicture();
        }
    }

    /// <summary>The identifier of the shape, as <c>Shape.ID</c> in Word.</summary>
    public int ID
    {
        get { return unchecked((int)(DocProperties.Id?.Value ?? 0U)); }
    }

    /// <summary>The text a screen reader says for the shape, as <c>Shape.AlternativeText</c> in Word.</summary>
    public string AlternativeText
    {
        get { return DocProperties.Description?.Value ?? string.Empty; }
        set
        {
            DocProperties.Description = Text(value);
            SyncPicture();
        }
    }

    /// <summary>The title of the shape, as <c>Shape.Title</c> in Word.</summary>
    public string Title
    {
        get { return DocProperties.Title?.Value ?? string.Empty; }
        set
        {
            DocProperties.Title = Text(value);
            SyncPicture();
        }
    }

    private static StringValue Text(string value)
    {
        return string.IsNullOrEmpty(value) ? null : new StringValue(value);
    }

    //Once the name, description or title is set, Word gives the picture itself the same ones.
    private void SyncPicture()
    {
        Pic.NonVisualDrawingProperties picture = AnchorElement.Descendants<Pic.NonVisualDrawingProperties>().FirstOrDefault();
        if (picture == null)
        {
            return;
        }
        picture.Name = DocProperties.Name?.Value;
        picture.Description = Text(DocProperties.Description?.Value);
        picture.Title = Text(DocProperties.Title?.Value);
        picture.Hidden = DocProperties.Hidden?.Value == true ? new BooleanValue(true) : null;
    }

    /// <summary>False if the shape is hidden, as <c>Shape.Visible</c> in Word.</summary>
    public bool Visible
    {
        get { return DocProperties.Hidden?.Value != true; }
        set
        {
            DocProperties.Hidden = value ? null : new BooleanValue(true);
            SyncPicture();
        }
    }

    /// <summary>True if the anchor cannot be moved, as <c>Shape.LockAnchor</c> in Word.</summary>
    public bool LockAnchor
    {
        get { return AnchorElement.Locked?.Value == true; }
        set { AnchorElement.Locked = value; }
    }

    /// <summary>True if changing the width or the height changes the other, as <c>Shape.LockAspectRatio</c> in Word.</summary>
    public bool LockAspectRatio
    {
        get { return AnchorElement.GetFirstChild<WpDrawing.NonVisualGraphicFrameDrawingProperties>()?.GetFirstChild<A.GraphicFrameLocks>()?.NoChangeAspect?.Value == true; }
        set
        {
            WpDrawing.NonVisualGraphicFrameDrawingProperties properties = AnchorElement.GetFirstChild<WpDrawing.NonVisualGraphicFrameDrawingProperties>();
            A.GraphicFrameLocks locks = properties.GetFirstChild<A.GraphicFrameLocks>();
            if (locks == null)
            {
                locks = new A.GraphicFrameLocks();
                properties.AppendChild(locks);
            }
            locks.NoChangeAspect = value ? new BooleanValue(true) : null;
        }
    }

    private long Cx
    {
        get { return AnchorElement.Extent?.Cx?.Value ?? 0; }
    }

    private long Cy
    {
        get { return AnchorElement.Extent?.Cy?.Value ?? 0; }
    }

    private void SetSize(long cx, long cy)
    {
        AnchorElement.Extent.Cx = cx;
        AnchorElement.Extent.Cy = cy;
        foreach (A.Extents extents in AnchorElement.Descendants<A.Extents>())
        {
            extents.Cx = cx;
            extents.Cy = cy;
        }
        SetEffectExtent();
    }

    private static long ToEmu(double points)
    {
        if (points < 0 || double.IsNaN(points) || double.IsInfinity(points))
        {
            throw new ArgumentOutOfRangeException(nameof(points));
        }
        return (long)Math.Round(points * EmuPerPoint);
    }

    /// <summary>The width in points, as <c>Shape.Width</c> in Word. With <see cref="LockAspectRatio"/> the height follows.</summary>
    public double Width
    {
        get { return Cx / EmuPerPoint; }
        set
        {
            long cx = ToEmu(value);
            SetSize(cx, LockAspectRatio && Cx > 0 ? (long)Math.Round((double)Cy * cx / Cx) : Cy);
        }
    }

    /// <summary>The height in points, as <c>Shape.Height</c> in Word. With <see cref="LockAspectRatio"/> the width follows.</summary>
    public double Height
    {
        get { return Cy / EmuPerPoint; }
        set
        {
            long cy = ToEmu(value);
            SetSize(LockAspectRatio && Cy > 0 ? (long)Math.Round((double)Cx * cy / Cy) : Cx, cy);
        }
    }

    /// <summary>The rotation in degrees (0 to 360), as <c>Shape.Rotation</c> in Word.</summary>
    public double Rotation
    {
        get { return (AnchorElement.Descendants<A.Transform2D>().FirstOrDefault()?.Rotation?.Value ?? 0) / (double)RotationUnit; }
        set
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            double degrees = value % 360;
            if (degrees < 0)
            {
                degrees += 360;
            }
            foreach (A.Transform2D transform in AnchorElement.Descendants<A.Transform2D>())
            {
                transform.Rotation = degrees == 0 ? null : new Int32Value((int)Math.Round(degrees * RotationUnit));
            }
            SetEffectExtent();
        }
    }

    //The room a rotated shape takes beyond its frame, in the steps of 1.5 points Word rounds it up to.
    private void SetEffectExtent()
    {
        WpDrawing.EffectExtent extent = AnchorElement.GetFirstChild<WpDrawing.EffectExtent>();
        double radians = Rotation * Math.PI / 180;
        if (extent == null)
        {
            return;
        }
        double width = Math.Abs(Cx * Math.Cos(radians)) + Math.Abs(Cy * Math.Sin(radians));
        double height = Math.Abs(Cx * Math.Sin(radians)) + Math.Abs(Cy * Math.Cos(radians));
        long across = Rotation == 0 ? 0 : RoundUp((width - Cx) / 2);
        long down = Rotation == 0 ? 0 : RoundUp((height - Cy) / 2);
        extent.LeftEdge = across;
        extent.RightEdge = 0L;
        extent.TopEdge = down;
        extent.BottomEdge = down;
    }

    private static long RoundUp(double emu)
    {
        return emu <= 0 ? 0 : (long)Math.Ceiling(emu / EffectExtentUnit) * EffectExtentUnit;
    }

    /// <summary>What <see cref="Left"/> is measured from, as <c>Shape.RelativeHorizontalPosition</c> in Word.</summary>
    public RelativeHorizontalPosition RelativeHorizontalPosition
    {
        get
        {
            string from = AnchorElement.HorizontalPosition?.RelativeFrom?.InnerText;
            return from == "margin" ? RelativeHorizontalPosition.Margin
                : from == "page" ? RelativeHorizontalPosition.Page
                : from == "character" ? RelativeHorizontalPosition.Character
                : from == "leftMargin" ? RelativeHorizontalPosition.LeftMarginArea
                : from == "rightMargin" ? RelativeHorizontalPosition.RightMarginArea
                : from == "insideMargin" ? RelativeHorizontalPosition.InnerMarginArea
                : from == "outsideMargin" ? RelativeHorizontalPosition.OuterMarginArea : RelativeHorizontalPosition.Column;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            AnchorElement.HorizontalPosition.RelativeFrom = value == RelativeHorizontalPosition.Margin ? WpDrawing.HorizontalRelativePositionValues.Margin
                : value == RelativeHorizontalPosition.Page ? WpDrawing.HorizontalRelativePositionValues.Page
                : value == RelativeHorizontalPosition.Character ? WpDrawing.HorizontalRelativePositionValues.Character
                : value == RelativeHorizontalPosition.LeftMarginArea ? WpDrawing.HorizontalRelativePositionValues.LeftMargin
                : value == RelativeHorizontalPosition.RightMarginArea ? WpDrawing.HorizontalRelativePositionValues.RightMargin
                : value == RelativeHorizontalPosition.InnerMarginArea ? WpDrawing.HorizontalRelativePositionValues.InsideMargin
                : value == RelativeHorizontalPosition.OuterMarginArea ? WpDrawing.HorizontalRelativePositionValues.OutsideMargin
                : WpDrawing.HorizontalRelativePositionValues.Column;
        }
    }

    /// <summary>What <see cref="Top"/> is measured from, as <c>Shape.RelativeVerticalPosition</c> in Word.</summary>
    public RelativeVerticalPosition RelativeVerticalPosition
    {
        get
        {
            string from = AnchorElement.VerticalPosition?.RelativeFrom?.InnerText;
            return from == "margin" ? RelativeVerticalPosition.Margin
                : from == "page" ? RelativeVerticalPosition.Page
                : from == "line" ? RelativeVerticalPosition.Line
                : from == "topMargin" ? RelativeVerticalPosition.TopMarginArea
                : from == "bottomMargin" ? RelativeVerticalPosition.BottomMarginArea
                : from == "insideMargin" ? RelativeVerticalPosition.InnerMarginArea
                : from == "outsideMargin" ? RelativeVerticalPosition.OuterMarginArea : RelativeVerticalPosition.Paragraph;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            AnchorElement.VerticalPosition.RelativeFrom = value == RelativeVerticalPosition.Margin ? WpDrawing.VerticalRelativePositionValues.Margin
                : value == RelativeVerticalPosition.Page ? WpDrawing.VerticalRelativePositionValues.Page
                : value == RelativeVerticalPosition.Line ? WpDrawing.VerticalRelativePositionValues.Line
                : value == RelativeVerticalPosition.TopMarginArea ? WpDrawing.VerticalRelativePositionValues.TopMargin
                : value == RelativeVerticalPosition.BottomMarginArea ? WpDrawing.VerticalRelativePositionValues.BottomMargin
                : value == RelativeVerticalPosition.InnerMarginArea ? WpDrawing.VerticalRelativePositionValues.InsideMargin
                : value == RelativeVerticalPosition.OuterMarginArea ? WpDrawing.VerticalRelativePositionValues.OutsideMargin
                : WpDrawing.VerticalRelativePositionValues.Paragraph;
        }
    }

    /// <summary>
    /// The distance in points from <see cref="RelativeHorizontalPosition"/> to the left edge, as <c>Shape.Left</c> in Word. Changing the
    /// thing it is measured from does not move the shape on the page here, as no layout is known.
    /// </summary>
    public double Left
    {
        get { return (AnchorElement.HorizontalPosition?.PositionOffset?.Text is string text && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0) / EmuPerPoint; }
        set
        {
            WpDrawing.HorizontalPosition position = AnchorElement.HorizontalPosition;
            position.RemoveAllChildren();
            position.AppendChild(new WpDrawing.PositionOffset(ToSigned(value).ToString(CultureInfo.InvariantCulture)));
        }
    }

    /// <summary>
    /// The distance in points from <see cref="RelativeVerticalPosition"/> to the top edge, as <c>Shape.Top</c> in Word.
    /// </summary>
    public double Top
    {
        get { return (AnchorElement.VerticalPosition?.PositionOffset?.Text is string text && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0) / EmuPerPoint; }
        set
        {
            WpDrawing.VerticalPosition position = AnchorElement.VerticalPosition;
            position.RemoveAllChildren();
            position.AppendChild(new WpDrawing.PositionOffset(ToSigned(value).ToString(CultureInfo.InvariantCulture)));
        }
    }

    private static long ToSigned(double points)
    {
        if (double.IsNaN(points) || double.IsInfinity(points))
        {
            throw new ArgumentOutOfRangeException(nameof(points));
        }
        return (long)Math.Round(points * EmuPerPoint);
    }

    /// <summary>How text wraps around the shape, as <c>Shape.WrapFormat</c> in Word.</summary>
    public WrapFormatWrapper WrapFormat
    {
        get { return new WrapFormatWrapper(this); }
    }

    private List<ShapeWrapper> Siblings()
    {
        return new ShapeWrapperCollection(new RangeWrapper(StoryMap.FindRoot(_drawing))).ToList();
    }

    /// <summary>The place of the shape in the stack of shapes, from 1 (the lowest), as <c>Shape.ZOrderPosition</c> in Word.</summary>
    public int ZOrderPosition
    {
        get { return Siblings().FindIndex(s => s == this) + 1; }
    }

    private uint Height_
    {
        get { return AnchorElement.RelativeHeight?.Value ?? 0U; }
    }

    /// <summary>Moves the shape in the stack, as <c>Shape.ZOrder</c> in Word.</summary>
    public void ZOrder(ZOrderCommand command)
    {
        if (!EnumCompat.IsDefined(command))
        {
            throw new ArgumentOutOfRangeException(nameof(command));
        }
        List<ShapeWrapper> all = Siblings();
        int index = all.FindIndex(s => s == this);
        switch (command)
        {
            case ZOrderCommand.BringToFront:
                AnchorElement.RelativeHeight = all.Max(s => s.Height_) + 1024U;
                break;
            case ZOrderCommand.SendToBack:
                AnchorElement.RelativeHeight = all.Min(s => s.Height_) - 1025U;
                break;
            case ZOrderCommand.BringForward:
                Swap(all, index, index + 1);
                break;
            case ZOrderCommand.SendBackward:
                Swap(all, index, index - 1);
                break;
            case ZOrderCommand.BringInFrontOfText:
                AnchorElement.BehindDoc = false;
                break;
            default:
                AnchorElement.BehindDoc = true;
                break;
        }
    }

    private void Swap(List<ShapeWrapper> all, int from, int to)
    {
        if (to < 0 || to >= all.Count)
        {
            return;
        }
        uint mine = all[from].Height_;
        AnchorElement.RelativeHeight = all[to].Height_;
        all[to].AnchorElement.RelativeHeight = mine;
    }

    /// <summary>The place in the text the shape is anchored at, as <c>Shape.Anchor</c> in Word.</summary>
    public RangeWrapper Anchor
    {
        get { return AnchorAtom().range; }
    }

    private (RangeWrapper range, Atom atom) AnchorAtom()
    {
        StoryMap map = StoryMap.Build(StoryMap.FindRoot(_drawing));
        Atom atom = map.Atoms.First(a => a.Kind == AtomKind.Object && (ReferenceEquals(a.Element, _drawing) || a.Element.Descendants<Drawing>().Contains(_drawing)));
        return (RangeWrapper.FromMap(map, atom.Start, atom.End), atom);
    }

    /// <summary>
    /// Takes the shape away, as <c>Shape.Delete</c> in Word. A picture file nothing uses any more goes out of the package.
    /// </summary>
    public void Delete()
    {
        A.Blip blip = Blip;
        string embed = blip?.Embed?.Value;
        string link = blip?.Link?.Value;
        OpenXmlElement root = _drawing.GetRoot();
        OpenXmlPartContainer container = root is OpenXmlPartRootElement partRoot ? partRoot.OpenXmlPart : null;

        Anchor.Delete();
        PictureFile.Release(root, container, embed, link);
    }

    /// <summary>
    /// Makes the shape an inline shape in the text where it is anchored, as <c>Shape.ConvertToInlineShape</c> in Word.
    /// </summary>
    public InlineShapeWrapper ConvertToInlineShape()
    {
        WpDrawing.Anchor anchor = AnchorElement;
        WpDrawing.Inline inline = new WpDrawing.Inline() { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U };
        inline.AppendChild((OpenXmlElement)anchor.Extent.CloneNode(true));
        inline.AppendChild(new WpDrawing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
        inline.AppendChild((OpenXmlElement)anchor.GetFirstChild<WpDrawing.DocProperties>().CloneNode(true));
        inline.AppendChild((OpenXmlElement)anchor.GetFirstChild<WpDrawing.NonVisualGraphicFrameDrawingProperties>().CloneNode(true));
        inline.AppendChild((OpenXmlElement)anchor.GetFirstChild<A.Graphic>().CloneNode(true));
        Drawing replacement = new Drawing(inline);
        _drawing.InsertBeforeSelf(replacement);
        _drawing.Remove();
        return new InlineShapeWrapper(replacement);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as ShapeWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(ShapeWrapper other)
    {
        return !(other is null) && ReferenceEquals(_drawing, other._drawing);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _drawing.GetHashCode();
    }

    public static bool operator ==(ShapeWrapper left, ShapeWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(ShapeWrapper left, ShapeWrapper right)
    {
        return !(left == right);
    }

    /// <summary>
    /// Makes the floating form of a picture from the inline one the picture code builds, as Word writes it.
    /// </summary>
    internal static Drawing NewAnchored(Drawing inlineDrawing, double left, double top, uint relativeHeight, bool toTwips = false)
    {
        WpDrawing.Inline inline = inlineDrawing.GetFirstChild<WpDrawing.Inline>();
        WpDrawing.Anchor anchor = new WpDrawing.Anchor()
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 114300U,
            DistanceFromRight = 114300U,
            SimplePos = false,
            RelativeHeight = relativeHeight,
            BehindDoc = false,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = true
        };
        anchor.AppendChild(new WpDrawing.SimplePosition() { X = 0L, Y = 0L });
        anchor.AppendChild(new WpDrawing.HorizontalPosition(new WpDrawing.PositionOffset(ToSigned(left).ToString(CultureInfo.InvariantCulture)))
        { RelativeFrom = WpDrawing.HorizontalRelativePositionValues.Column });
        anchor.AppendChild(new WpDrawing.VerticalPosition(new WpDrawing.PositionOffset(ToSigned(top).ToString(CultureInfo.InvariantCulture)))
        { RelativeFrom = WpDrawing.VerticalRelativePositionValues.Paragraph });
        anchor.AppendChild((OpenXmlElement)inline.Extent.CloneNode(true));
        anchor.AppendChild(new WpDrawing.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L });
        anchor.AppendChild(new WpDrawing.WrapNone());
        anchor.AppendChild((OpenXmlElement)inline.GetFirstChild<WpDrawing.DocProperties>().CloneNode(true));
        anchor.AppendChild((OpenXmlElement)inline.GetFirstChild<WpDrawing.NonVisualGraphicFrameDrawingProperties>().CloneNode(true));
        anchor.AppendChild((OpenXmlElement)inline.GetFirstChild<A.Graphic>().CloneNode(true));
        Wp14.RelativeWidth sizeWidth = new Wp14.RelativeWidth(new Wp14.PercentageWidth("0"));
        sizeWidth.SetAttribute(new OpenXmlAttribute(string.Empty, "relativeFrom", string.Empty, "margin"));
        Wp14.RelativeHeight sizeHeight = new Wp14.RelativeHeight(new Wp14.PercentageHeight("0"));
        sizeHeight.SetAttribute(new OpenXmlAttribute(string.Empty, "relativeFrom", string.Empty, "margin"));
        anchor.AppendChild(sizeWidth);
        anchor.AppendChild(sizeHeight);
        Drawing result = new Drawing(anchor);
        if (toTwips)
        {
            //Word gives the shape a size in whole twips.
            ShapeWrapper shape = new ShapeWrapper(result);
            shape.Width = Math.Round(shape.Width * 20) / 20;
            shape.Height = Math.Round(shape.Height * 20) / 20;
        }
        return result;
    }
}

/// <summary>
/// The floating shapes of a document or header, as <c>Shapes</c> in Word.
/// </summary>
public partial class ShapeWrapperCollection : IEnumerable<ShapeWrapper>
{
    private const uint FirstHeight = 251658240U;
    private readonly RangeWrapper _range;

    internal ShapeWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range whose story the shapes are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    //In the order of the stack, lowest first.
    private List<ShapeWrapper> GetItems()
    {
        OpenXmlElement root = StoryMap.FindRoot(_range.AnchorElement);
        return root.Descendants<Drawing>().Where(d => d.GetFirstChild<WpDrawing.Anchor>() != null).Select(d => new ShapeWrapper(d))
            .OrderBy(s => s.AnchorElement.RelativeHeight?.Value ?? 0U).ToList();
    }

    //The place above all the shapes of the story.
    internal static uint NextRelativeHeight(OpenXmlElement inStory)
    {
        List<ShapeWrapper> existing = new ShapeWrapperCollection(new RangeWrapper(StoryMap.FindRoot(inStory))).GetItems();
        return existing.Count == 0 ? FirstHeight : existing.Max(s => s.AnchorElement.RelativeHeight?.Value ?? 0U) + 1024U;
    }

    /// <summary>The number of shapes.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The shape at the index (0-based), lowest in the stack first, as <c>Shapes.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no shape at the index.</exception>
    public ShapeWrapper this[int index]
    {
        get
        {
            List<ShapeWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Inserts a picture as a floating shape, as <c>Shapes.AddPicture</c> in Word. The shape is anchored at the start of the paragraph the
    /// anchor is in, in front of the text, and gets the next place in the stack.
    /// </summary>
    /// <param name="fileName">The picture file (PNG, JPEG, GIF or BMP).</param>
    /// <param name="linkToFile">True to link to the file instead of storing the picture in the document.</param>
    /// <param name="saveWithDocument">Whether to store the picture in the document; a linked picture is only linked.</param>
    /// <param name="left">The distance in points from the column to the left edge.</param>
    /// <param name="top">The distance in points from the paragraph to the top edge.</param>
    /// <param name="width">The width in points; the picture's own when negative.</param>
    /// <param name="height">The height in points; the picture's own when negative.</param>
    /// <param name="anchor">Where to anchor the shape; the start of the range this collection is for when null.</param>
    /// <exception cref="ArgumentException">The file is not a PNG, JPEG, GIF or BMP picture.</exception>
    public ShapeWrapper AddPicture(string fileName, bool linkToFile = false, bool saveWithDocument = true, double left = 0, double top = 0,
        double width = -1, double height = -1, RangeWrapper anchor = null)
    {
        Guard.ThrowIfNull(fileName);
        PictureFile file = PictureFile.Read(File.ReadAllBytes(fileName));
        Paragraph paragraph = (anchor ?? _range).Duplicate().ParagraphOrFirst();
        RangeWrapper at = new RangeWrapper(paragraph);
        at.Collapse(CollapseDirection.Start);

        uint relativeHeight = NextRelativeHeight(_range.AnchorElement);
        Drawing drawing = null;
        at.InsertNoteReference(target =>
        {
            MainDocumentPart main = StyleHelper.GetMainDocumentPart(target);
            OpenXmlPartContainer container = target.GetRoot() is OpenXmlPartRootElement root ? root.OpenXmlPart : null;
            if (main == null || container == null)
            {
                throw new NotSupportedException("A picture can be inserted only in a document.");
            }
            Drawing inline = InlineShapeWrapper.NewPicture(container, main, file, fileName, linkToFile);
            drawing = ShapeWrapper.NewAnchored(inline, left, top, relativeHeight);
            ShapeWrapper shape = new ShapeWrapper(drawing);
            //The size of the file, in whole twips, where none is given.
            shape.Width = width >= 0 ? width : Math.Round(file.WidthEmu / 635.0) / 20;
            shape.Height = height >= 0 ? height : Math.Round(file.HeightEmu / 635.0) / 20;
            return new List<Run> { new Run(new RunProperties(new NoProof()), drawing) };
        });
        return new ShapeWrapper(drawing);
    }

    /// <inheritdoc/>
    public IEnumerator<ShapeWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

internal static class RangeAnchorExtensions
{
    internal static Paragraph ParagraphOrFirst(this RangeWrapper range)
    {
        return range.StartElement as Paragraph ?? range.StartElement?.Ancestors<Paragraph>().FirstOrDefault() ?? range.Paragraphs[0].Paragraph;
    }
}
