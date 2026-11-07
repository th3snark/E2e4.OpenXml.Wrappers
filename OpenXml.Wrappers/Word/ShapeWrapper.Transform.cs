using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using WpDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>Which way a shape is turned over, as <c>MsoFlipCmd</c> in Office (with the same values).</summary>
public enum FlipCommand
{
    /// <summary>From left to right.</summary>
    Horizontal = 0,

    /// <summary>From top to bottom.</summary>
    Vertical = 1
}

/// <summary>The point of a shape that stays where it is when the shape is scaled, as <c>MsoScaleFrom</c> in Office (with the same values).</summary>
public enum ScaleFrom
{
    /// <summary>The top left corner.</summary>
    TopLeft = 0,

    /// <summary>The middle.</summary>
    Middle = 1,

    /// <summary>The bottom right corner.</summary>
    BottomRight = 2
}

public partial class ShapeWrapper
{
    private const string DecorativeUri = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";
    private const string DecorativeNamespace = "http://schemas.microsoft.com/office/drawing/2017/decorative";

    /// <summary>True if the shape is turned over from left to right, as <c>Shape.HorizontalFlip</c> in Word.</summary>
    public bool HorizontalFlip
    {
        get { return AnchorElement.Descendants<A.Transform2D>().FirstOrDefault()?.HorizontalFlip?.Value == true; }
    }

    /// <summary>True if the shape is turned over from top to bottom, as <c>Shape.VerticalFlip</c> in Word.</summary>
    public bool VerticalFlip
    {
        get { return AnchorElement.Descendants<A.Transform2D>().FirstOrDefault()?.VerticalFlip?.Value == true; }
    }

    /// <summary>
    /// Turns the shape over, as <c>Shape.Flip</c> in Word. A shape that is rotated is also given the mirror image of its rotation, so that
    /// it ends up where the flip puts it.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="flipCmd"/> is not a value of <see cref="FlipCommand"/>.</exception>
    public void Flip(FlipCommand flipCmd)
    {
        ThrowIfChild();
        ThrowIfGroup();
        if (!EnumCompat.IsDefined(flipCmd))
        {
            throw new ArgumentOutOfRangeException(nameof(flipCmd));
        }
        double rotation = Rotation;
        foreach (A.Transform2D transform in AnchorElement.Descendants<A.Transform2D>())
        {
            if (flipCmd == FlipCommand.Horizontal)
            {
                transform.HorizontalFlip = transform.HorizontalFlip?.Value == true ? null : new BooleanValue(true);
            }
            else
            {
                transform.VerticalFlip = transform.VerticalFlip?.Value == true ? null : new BooleanValue(true);
            }
        }
        if (rotation != 0)
        {
            Rotation = 360 - rotation;
        }
    }

    /// <summary>Moves the shape to the right by a distance in points (to the left if it is negative), as <c>Shape.IncrementLeft</c> in Word.</summary>
    public void IncrementLeft(double increment)
    {
        Left += increment;
    }

    /// <summary>Moves the shape down by a distance in points (up if it is negative), as <c>Shape.IncrementTop</c> in Word.</summary>
    public void IncrementTop(double increment)
    {
        Top += increment;
    }

    /// <summary>Turns the shape by a number of degrees, clockwise if positive, as <c>Shape.IncrementRotation</c> in Word.</summary>
    public void IncrementRotation(double increment)
    {
        Rotation += increment;
    }

    /// <summary>Scales the width by a factor, as <c>Shape.ScaleWidth</c> in Word; the point <paramref name="fScale"/> names stays where it is.</summary>
    /// <param name="factor">The new width as a multiple of the width now; more than zero.</param>
    /// <param name="relativeToOriginalSize">Scaling relative to the original size is not supported.</param>
    /// <param name="fScale">The point that stays where it is.</param>
    /// <exception cref="NotSupportedException"><paramref name="relativeToOriginalSize"/> is true.</exception>
    public void ScaleWidth(double factor, bool relativeToOriginalSize = false, ScaleFrom fScale = ScaleFrom.TopLeft)
    {
        Scale(factor, relativeToOriginalSize, fScale, true);
    }

    /// <summary>Scales the height by a factor, as <c>Shape.ScaleHeight</c> in Word. See <see cref="ScaleWidth"/>.</summary>
    public void ScaleHeight(double factor, bool relativeToOriginalSize = false, ScaleFrom fScale = ScaleFrom.TopLeft)
    {
        Scale(factor, relativeToOriginalSize, fScale, false);
    }

    private void Scale(double factor, bool relativeToOriginalSize, ScaleFrom fScale, bool width)
    {
        if (!(factor > 0) || double.IsInfinity(factor))
        {
            throw new ArgumentOutOfRangeException(nameof(factor));
        }
        if (!EnumCompat.IsDefined(fScale))
        {
            throw new ArgumentOutOfRangeException(nameof(fScale));
        }
        if (relativeToOriginalSize)
        {
            throw new NotSupportedException("Scaling relative to the original size is not supported.");
        }
        long old = width ? Cx : Cy;
        long scaled = (long)Math.Round(old * factor);
        //Scaling one side leaves the other as it is, whatever the lock on the aspect ratio.
        SetSize(width ? scaled : Cx, width ? Cy : scaled);
        double shift = (scaled - old) / EmuPerPoint * (fScale == ScaleFrom.TopLeft ? 0 : fScale == ScaleFrom.Middle ? 0.5 : 1);
        if (shift != 0)
        {
            if (width)
            {
                Left -= shift;
            }
            else
            {
                Top -= shift;
            }
        }
    }

    /// <summary>
    /// Makes a copy of the shape, as <c>Shape.Duplicate</c> in Word: it has the same name and place, is anchored in front of the shape, and is
    /// in front of it.
    /// </summary>
    public ShapeWrapper Duplicate()
    {
        ThrowIfChild();
        Run run = _drawing.Ancestors<Run>().FirstOrDefault() ?? throw new InvalidOperationException("The shape is not in the text.");
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(run) ?? throw new NotSupportedException("A shape can be duplicated only in a document.");
        Run copy = (Run)run.CloneNode(true);
        Drawing drawing = copy.Descendants<Drawing>().First();
        WpDrawing.Anchor anchor = drawing.GetFirstChild<WpDrawing.Anchor>();
        uint id = main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<WpDrawing.DocProperties>())
            .Select(d => d.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
        anchor.GetFirstChild<WpDrawing.DocProperties>().Id = id;
        anchor.RelativeHeight = ShapeWrapperCollection.NextRelativeHeight(run);
        run.InsertBeforeSelf(copy);
        return new ShapeWrapper(drawing);
    }

    /// <summary>True if the shape is placed inside a table cell when its anchor is in one, as <c>Shape.LayoutInCell</c> in Word. The default is true.</summary>
    public bool LayoutInCell
    {
        get { return AnchorElement.LayoutInCell?.Value ?? true; }
        set { AnchorElement.LayoutInCell = value; }
    }

    /// <summary>True if the shape is only for decoration and has no alternative text, as <c>Shape.Decorative</c> in Word.</summary>
    public bool Decorative
    {
        get { return DocProperties.Descendants<A.NonVisualDrawingPropertiesExtension>().Any(e => e.Uri?.Value == DecorativeUri); }
        set
        {
            WpDrawing.DocProperties properties = DocProperties;
            foreach (A.NonVisualDrawingPropertiesExtension extension in properties.Descendants<A.NonVisualDrawingPropertiesExtension>().Where(e => e.Uri?.Value == DecorativeUri).ToList())
            {
                extension.Remove();
            }
            A.NonVisualDrawingPropertiesExtensionList list = properties.GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>();
            if (list != null && !list.HasChildren)
            {
                list.Remove();
                list = null;
            }
            if (!value)
            {
                return;
            }
            if (list == null)
            {
                list = new A.NonVisualDrawingPropertiesExtensionList();
                properties.AppendChild(list);
            }
            A.NonVisualDrawingPropertiesExtension decorative = new A.NonVisualDrawingPropertiesExtension() { Uri = DecorativeUri };
            OpenXmlUnknownElement mark = new OpenXmlUnknownElement("adec", "decorative", DecorativeNamespace);
            mark.SetAttribute(new OpenXmlAttribute("val", string.Empty, "1"));
            decorative.AppendChild(mark);
            list.AppendChild(decorative);
        }
    }
}
