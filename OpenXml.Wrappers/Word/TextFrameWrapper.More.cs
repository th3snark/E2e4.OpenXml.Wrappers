using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OpenXml.Wrappers.Word;

/// <summary>Where the text of a shape sits from top to bottom, as <c>MsoVerticalAnchor</c> in Office (with the same values).</summary>
public enum VerticalAnchor
{
    /// <summary>At the top.</summary>
    Top = 1,

    /// <summary>In the middle.</summary>
    Middle = 3,

    /// <summary>At the bottom.</summary>
    Bottom = 4
}

/// <summary>Whether the text of a shape is centred as a block, as <c>MsoHorizontalAnchor</c> in Office (with the same values).</summary>
public enum HorizontalAnchor
{
    /// <summary>The text starts at the left margin.</summary>
    None = 1,

    /// <summary>The text is centred as a block.</summary>
    Center = 2
}

public partial class TextFrameWrapper
{
    private const double EmuPerPoint = 12700;

    private Wps.TextBodyProperties Body
    {
        get { return _shape.Shape.GetFirstChild<Wps.TextBodyProperties>() ?? throw new InvalidOperationException("The shape has no text frame."); }
    }

    /// <summary>The shape the text belongs to, as <c>TextFrame.Parent</c> in Word.</summary>
    public ShapeWrapper Parent
    {
        get { return _shape; }
    }

    private double GetMargin(Func<Wps.TextBodyProperties, int?> read, double fallback)
    {
        int? emu = read(Body);
        return emu.HasValue ? emu.Value / EmuPerPoint : fallback;
    }

    private static int ToInset(double points)
    {
        if (points < 0 || double.IsNaN(points) || double.IsInfinity(points) || points * EmuPerPoint > int.MaxValue)
        {
            throw new ArgumentOutOfRangeException(nameof(points));
        }
        return (int)Math.Round(points * EmuPerPoint);
    }

    /// <summary>The distance in points from the left edge of the shape to its text, as <c>TextFrame.MarginLeft</c> in Word. The default is 7.2.</summary>
    public double MarginLeft
    {
        get { return GetMargin(b => b.LeftInset?.Value, 7.2); }
        set { Body.LeftInset = ToInset(value); }
    }

    /// <summary>The distance in points from the right edge of the shape to its text, as <c>TextFrame.MarginRight</c> in Word. The default is 7.2.</summary>
    public double MarginRight
    {
        get { return GetMargin(b => b.RightInset?.Value, 7.2); }
        set { Body.RightInset = ToInset(value); }
    }

    /// <summary>The distance in points from the top edge of the shape to its text, as <c>TextFrame.MarginTop</c> in Word. The default is 3.6.</summary>
    public double MarginTop
    {
        get { return GetMargin(b => b.TopInset?.Value, 3.6); }
        set { Body.TopInset = ToInset(value); }
    }

    /// <summary>The distance in points from the bottom edge of the shape to its text, as <c>TextFrame.MarginBottom</c> in Word. The default is 3.6.</summary>
    public double MarginBottom
    {
        get { return GetMargin(b => b.BottomInset?.Value, 3.6); }
        set { Body.BottomInset = ToInset(value); }
    }

    /// <summary>Where the text sits from top to bottom, as <c>TextFrame.VerticalAnchor</c> in Word.</summary>
    public VerticalAnchor VerticalAnchor
    {
        get
        {
            A.TextAnchoringTypeValues? anchor = Body.Anchor?.Value;
            return anchor == A.TextAnchoringTypeValues.Center ? Word.VerticalAnchor.Middle : anchor == A.TextAnchoringTypeValues.Bottom ? Word.VerticalAnchor.Bottom : Word.VerticalAnchor.Top;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Body.Anchor = value == Word.VerticalAnchor.Middle ? A.TextAnchoringTypeValues.Center : value == Word.VerticalAnchor.Bottom ? A.TextAnchoringTypeValues.Bottom : A.TextAnchoringTypeValues.Top;
        }
    }

    /// <summary>Whether the text is centred as a block, as <c>TextFrame.HorizontalAnchor</c> in Word.</summary>
    public HorizontalAnchor HorizontalAnchor
    {
        get { return Body.AnchorCenter?.Value == true ? Word.HorizontalAnchor.Center : Word.HorizontalAnchor.None; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Body.AnchorCenter = value == Word.HorizontalAnchor.Center;
        }
    }

    /// <summary>True if text wraps at the right edge of the shape, as <c>TextFrame.WordWrap</c> in Word. The default is true.</summary>
    public bool WordWrap
    {
        get { return Body.Wrap?.Value != A.TextWrappingValues.None; }
        set { Body.Wrap = value ? A.TextWrappingValues.Square : A.TextWrappingValues.None; }
    }

    /// <summary>True if the shape grows or shrinks to fit its text, as <c>TextFrame.AutoSize</c> in Word.</summary>
    public bool AutoSize
    {
        get { return Body.GetFirstChild<A.ShapeAutoFit>() != null; }
        set
        {
            Wps.TextBodyProperties body = Body;
            body.RemoveAllChildren<A.NoAutoFit>();
            body.RemoveAllChildren<A.NormalAutoFit>();
            body.RemoveAllChildren<A.ShapeAutoFit>();
            if (value)
            {
                body.AppendChild(new A.ShapeAutoFit());
            }
            else
            {
                body.AppendChild(new A.NoAutoFit());
            }
        }
    }

    /// <summary>The direction of the text, as <c>TextFrame.Orientation</c> in Word.</summary>
    public TextOrientation Orientation
    {
        get
        {
            A.TextVerticalValues? vertical = Body.Vertical?.Value;
            return vertical == A.TextVerticalValues.Vertical270 ? TextOrientation.Upward
                : vertical == A.TextVerticalValues.Vertical ? TextOrientation.Downward
                : vertical == A.TextVerticalValues.EastAsianVetical ? TextOrientation.VerticalFarEast
                : vertical == A.TextVerticalValues.MongolianVertical ? TextOrientation.Vertical
                : TextOrientation.Horizontal;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Body.Vertical = value == TextOrientation.Upward ? A.TextVerticalValues.Vertical270
                : value == TextOrientation.Downward ? A.TextVerticalValues.Vertical
                : value == TextOrientation.VerticalFarEast ? A.TextVerticalValues.EastAsianVetical
                : value == TextOrientation.Vertical ? A.TextVerticalValues.MongolianVertical
                : A.TextVerticalValues.Horizontal;
        }
    }

    /// <summary>Deletes the text of the shape, as <c>TextFrame.DeleteText</c> in Word.</summary>
    public void DeleteText()
    {
        TextBoxContent content = _shape.TextBoxContent;
        if (content == null)
        {
            return;
        }
        content.RemoveAllChildren();
        content.AppendChild(new Paragraph());
    }
}
