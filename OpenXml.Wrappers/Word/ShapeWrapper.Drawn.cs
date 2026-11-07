using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using WpDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The shapes of the gallery the library can draw, as <c>MsoAutoShapeType</c> in Office (with the same values).
/// </summary>
public enum AutoShapeType
{
    /// <summary>A rectangle.</summary>
    Rectangle = 1,

    /// <summary>A diamond.</summary>
    Diamond = 4,

    /// <summary>A rectangle with rounded corners.</summary>
    RoundedRectangle = 5,

    /// <summary>A triangle with two equal sides.</summary>
    IsoscelesTriangle = 7,

    /// <summary>An oval.</summary>
    Oval = 9,

    /// <summary>A hexagon.</summary>
    Hexagon = 10,

    /// <summary>An arrow pointing right.</summary>
    RightArrow = 33,

    /// <summary>A parallelogram.</summary>
    Parallelogram = 2,

    /// <summary>A trapezoid.</summary>
    Trapezoid = 3,

    /// <summary>An octagon.</summary>
    Octagon = 6,

    /// <summary>A triangle with a right angle.</summary>
    RightTriangle = 8,

    /// <summary>A cross.</summary>
    Cross = 11,

    /// <summary>A cylinder.</summary>
    Can = 13,

    /// <summary>A cube.</summary>
    Cube = 14,

    /// <summary>A smiley face.</summary>
    SmileyFace = 17,

    /// <summary>A ring.</summary>
    Donut = 18,

    /// <summary>A circle with a bar across it.</summary>
    NoSymbol = 19,

    /// <summary>A heart.</summary>
    Heart = 21,

    /// <summary>A lightning bolt.</summary>
    LightningBolt = 22,

    /// <summary>A sun.</summary>
    Sun = 23,

    /// <summary>A moon.</summary>
    Moon = 24,

    /// <summary>An arrow pointing left.</summary>
    LeftArrow = 34,

    /// <summary>An arrow pointing up.</summary>
    UpArrow = 35,

    /// <summary>An arrow pointing down.</summary>
    DownArrow = 36,

    /// <summary>An arrow pointing left and right.</summary>
    LeftRightArrow = 37,

    /// <summary>A pentagon that points right, like an arrow.</summary>
    Pentagon = 51,

    /// <summary>The flowchart symbol for a process.</summary>
    FlowchartProcess = 61,

    /// <summary>The flowchart symbol for a decision.</summary>
    FlowchartDecision = 63,

    /// <summary>A star with five points.</summary>
    Star5 = 92,

    /// <summary>A cloud.</summary>
    Cloud = 179
}

/// <summary>The direction of the text in a text box, as <c>MsoTextOrientation</c> in Office (with the same value; only the usual one).</summary>
public enum TextOrientation
{
    /// <summary>Left to right.</summary>
    Horizontal = 1,

    /// <summary>Upward, from the bottom of the shape to the top.</summary>
    Upward = 2,

    /// <summary>Downward, from the top of the shape to the bottom.</summary>
    Downward = 3,

    /// <summary>Vertical text for Asian languages.</summary>
    VerticalFarEast = 4,

    /// <summary>Vertical, with the letters upright.</summary>
    Vertical = 5
}

/// <summary>
/// A colour of a shape, as <c>ColorFormat</c> in Word, of which only the red, green and blue value is kept.
/// </summary>
public class ShapeColorWrapper
{
    private readonly Func<int> _get;
    private readonly Action<int> _set;

    internal ShapeColorWrapper(Func<int> get, Action<int> set)
    {
        _get = get;
        _set = set;
    }

    /// <summary>The colour as <see cref="ColorHelper.RGB"/> makes it, as <c>ColorFormat.RGB</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a colour from 0 to 16777215.</exception>
    public int RGB
    {
        get { return _get(); }
        set { _set(WordColor.FromValue(value).Value); }
    }
}

/// <summary>
/// The fill of a shape, as <c>FillFormat</c> in Word: no fill or a solid colour with some transparency.
/// </summary>
public class FillFormatWrapper
{
    private readonly ShapeWrapper _shape;

    internal FillFormatWrapper(ShapeWrapper shape)
    {
        _shape = shape;
    }

    /// <summary>True if the shape is filled, as <c>FillFormat.Visible</c> in Word.</summary>
    public bool Visible
    {
        get
        {
            Wps.ShapeProperties properties = _shape.ShapeProperties;
            if (properties.GetFirstChild<A.NoFill>() != null)
            {
                return false;
            }
            return properties.GetFirstChild<A.SolidFill>() != null || _shape.FillFromStyle;
        }
        set
        {
            Wps.ShapeProperties properties = _shape.ShapeProperties;
            if (value)
            {
                if (properties.GetFirstChild<A.NoFill>() != null || properties.GetFirstChild<A.SolidFill>() == null)
                {
                    SetSolid(ForeColor.RGB, Transparency);
                }
            }
            else
            {
                _shape.SetFill(new A.NoFill());
            }
        }
    }

    /// <summary>The colour of the fill, as <c>FillFormat.ForeColor</c> in Word. Setting it makes the fill visible.</summary>
    public ShapeColorWrapper ForeColor
    {
        get { return new ShapeColorWrapper(GetColor, value => SetSolid(value, Transparency)); }
    }

    private int GetColor()
    {
        A.RgbColorModelHex color = _shape.ShapeProperties.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>();
        if (color?.Val?.Value != null)
        {
            return new WordColor(color.Val.Value).Value;
        }
        return _shape.FillFromStyle ? ShapeWrapper.ThemeAccent1 : 0xFFFFFF;
    }

    /// <summary>How much of the background shows through the fill (0 to 1), as <c>FillFormat.Transparency</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not from 0 to 1.</exception>
    public double Transparency
    {
        get
        {
            A.Alpha alpha = _shape.ShapeProperties.GetFirstChild<A.SolidFill>()?.Descendants<A.Alpha>().FirstOrDefault();
            return alpha?.Val?.Value == null ? 0 : 1 - (alpha.Val.Value / 100000.0);
        }
        set
        {
            if (double.IsNaN(value) || value < 0 || value > 1)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            SetSolid(ForeColor.RGB, value);
        }
    }

    /// <summary>Fills the shape with one colour, as <c>FillFormat.Solid</c> in Word.</summary>
    public void Solid()
    {
        SetSolid(ForeColor.RGB, Transparency);
    }

    private void SetSolid(int color, double transparency)
    {
        A.RgbColorModelHex rgb = new A.RgbColorModelHex();
        //The value of Word is blue, green, red, where the document holds red, green, blue.
        rgb.Val = WordColor.FromValue(color).Hex;
        if (transparency > 0)
        {
            rgb.AppendChild(new A.Alpha() { Val = (int)Math.Round((1 - transparency) * 100000) });
        }
        _shape.SetFill(new A.SolidFill(rgb));
    }
}

/// <summary>
/// The outline of a shape, as <c>LineFormat</c> in Word.
/// </summary>
public class LineFormatWrapper
{
    private const double EmuPerPoint = 12700;
    private readonly ShapeWrapper _shape;

    internal LineFormatWrapper(ShapeWrapper shape)
    {
        _shape = shape;
    }

    private A.Outline Outline
    {
        get { return _shape.ShapeProperties.GetFirstChild<A.Outline>(); }
    }

    /// <summary>True if the outline is drawn, as <c>LineFormat.Visible</c> in Word.</summary>
    public bool Visible
    {
        get
        {
            A.Outline outline = Outline;
            if (outline?.GetFirstChild<A.NoFill>() != null)
            {
                return false;
            }
            return outline?.GetFirstChild<A.SolidFill>() != null || _shape.LineFromStyle;
        }
        set
        {
            A.Outline outline = _shape.GetOrAddOutline();
            if (!value)
            {
                outline.RemoveAllChildren();
                outline.AppendChild(new A.NoFill());
            }
            else if (outline.GetFirstChild<A.NoFill>() != null || outline.GetFirstChild<A.SolidFill>() == null)
            {
                SetColor(ForeColor.RGB);
            }
        }
    }

    /// <summary>The colour of the outline, as <c>LineFormat.ForeColor</c> in Word. Setting it makes the outline visible.</summary>
    public ShapeColorWrapper ForeColor
    {
        get { return new ShapeColorWrapper(GetColor, SetColor); }
    }

    private int GetColor()
    {
        A.RgbColorModelHex color = Outline?.GetFirstChild<A.SolidFill>()?.GetFirstChild<A.RgbColorModelHex>();
        if (color?.Val?.Value != null)
        {
            return new WordColor(color.Val.Value).Value;
        }
        return _shape.LineFromStyle ? ShapeWrapper.ThemeLine : 0;
    }

    private void SetColor(int color)
    {
        A.Outline outline = _shape.GetOrAddOutline();
        foreach (OpenXmlElement fill in outline.ChildElements.Where(c => c is A.NoFill || c is A.SolidFill).ToList())
        {
            fill.Remove();
        }
        outline.InsertAt(new A.SolidFill(new A.RgbColorModelHex() { Val = WordColor.FromValue(color).Hex }), 0);
    }

    /// <summary>The width of the outline in points, as <c>LineFormat.Weight</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative.</exception>
    public double Weight
    {
        get { return Outline?.Width?.Value != null ? Outline.Width.Value / EmuPerPoint : _shape.LineFromStyle ? 1.5 : 0.75; }
        set
        {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _shape.GetOrAddOutline().Width = (int)Math.Round(value * EmuPerPoint);
        }
    }
}

/// <summary>
/// The text inside a shape, as <c>TextFrame</c> in Word.
/// </summary>
public partial class TextFrameWrapper
{
    private readonly ShapeWrapper _shape;

    internal TextFrameWrapper(ShapeWrapper shape)
    {
        _shape = shape;
    }

    /// <summary>The text of the shape, as <c>TextFrame.TextRange</c> in Word.</summary>
    public RangeWrapper TextRange
    {
        get { return new RangeWrapper(_shape.GetOrAddTextBoxContent()); }
    }

    /// <summary>True if the shape has text, as <c>TextFrame.HasText</c> in Word.</summary>
    public bool HasText
    {
        get
        {
            TextBoxContent content = _shape.TextBoxContent;
            return content != null && content.Descendants<Text>().Any(t => t.Text.Length > 0);
        }
    }
}

public partial class ShapeWrapper
{
    //The colours of the default style of Word 365 (the Office theme: accent 1 and its darker shade), which shapes without a colour of their own show.
    internal const int ThemeAccent1 = 8544277;
    internal const int ThemeLine = 3351556;

    private const string ShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    internal Wps.WordprocessingShape Shape
    {
        get { return AnchorElement?.Descendants<Wps.WordprocessingShape>().FirstOrDefault(); }
    }

    internal Wps.ShapeProperties ShapeProperties
    {
        get { return (Shape ?? throw new InvalidOperationException("The shape is a picture.")).GetFirstChild<Wps.ShapeProperties>(); }
    }

    internal bool FillFromStyle
    {
        get { return Shape?.GetFirstChild<Wps.ShapeStyle>()?.FillReference?.Index?.Value > 0; }
    }

    internal bool LineFromStyle
    {
        get { return Shape?.GetFirstChild<Wps.ShapeStyle>()?.LineReference?.Index?.Value > 0 || (IsTextBox && ShapeProperties.GetFirstChild<A.Outline>() == null); }
    }

    private bool IsTextBox
    {
        get { return Shape?.GetFirstChild<Wps.NonVisualDrawingShapeProperties>()?.TextBox?.Value == true; }
    }

    internal TextBoxContent TextBoxContent
    {
        get { return Shape?.Descendants<TextBoxContent>().FirstOrDefault(); }
    }

    /// <summary>Which of the shapes of the gallery this is, as <c>Shape.AutoShapeType</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The shape is a picture, or a line.</exception>
    public AutoShapeType AutoShapeType
    {
        get
        {
            string preset = ShapeProperties.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText;
            AutoShapeType found = Presets.Where(p => p.Preset == preset).Select(p => p.Type).FirstOrDefault();
            return found == 0 ? throw new InvalidOperationException("The shape is not one of the shapes of the gallery.") : found;
        }
    }

    /// <summary>The fill of the shape, as <c>Shape.Fill</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The shape is a picture.</exception>
    public FillFormatWrapper Fill
    {
        get
        {
            _ = ShapeProperties;
            return new FillFormatWrapper(this);
        }
    }

    /// <summary>The outline of the shape, as <c>Shape.Line</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The shape is a picture.</exception>
    public LineFormatWrapper Line
    {
        get
        {
            _ = ShapeProperties;
            return new LineFormatWrapper(this);
        }
    }

    /// <summary>The text of the shape, as <c>Shape.TextFrame</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The shape is a picture or a line.</exception>
    public TextFrameWrapper TextFrame
    {
        get
        {
            if (Shape == null || Shape.GetFirstChild<Wps.NonVisualConnectorProperties>() != null)
            {
                throw new InvalidOperationException("The shape has no text.");
            }
            return new TextFrameWrapper(this);
        }
    }

    internal TextBoxContent GetOrAddTextBoxContent()
    {
        TextBoxContent content = TextBoxContent;
        if (content == null)
        {
            //Text in a shape of the gallery is centred, as Word starts it; in a text box it is not.
            Paragraph first = IsTextBox ? new Paragraph() : new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
            Wps.TextBoxInfo2 box = new Wps.TextBoxInfo2(new TextBoxContent(first));
            Shape.GetFirstChild<Wps.TextBodyProperties>().InsertBeforeSelf(box);
            content = box.GetFirstChild<TextBoxContent>();
        }
        return content;
    }

    //Puts a fill in the place it has in the properties of the shape: after the geometry and before the outline.
    internal void SetFill(OpenXmlElement fill)
    {
        Wps.ShapeProperties properties = ShapeProperties;
        foreach (OpenXmlElement old in properties.ChildElements.Where(c => c is A.NoFill || c is A.SolidFill || c is A.GradientFill || c is A.BlipFill
            || c is A.PatternFill || c is A.GroupFill).ToList())
        {
            old.Remove();
        }
        OpenXmlElement geometry = properties.ChildElements.LastOrDefault(c => c is A.PresetGeometry || c is A.CustomGeometry || c is A.Transform2D);
        geometry.InsertAfterSelf(fill);
    }

    internal A.Outline GetOrAddOutline()
    {
        Wps.ShapeProperties properties = ShapeProperties;
        A.Outline outline = properties.GetFirstChild<A.Outline>();
        if (outline == null)
        {
            outline = new A.Outline();
            OpenXmlElement before = properties.ChildElements.LastOrDefault(c => c is A.PresetGeometry || c is A.CustomGeometry || c is A.Transform2D
                || c is A.NoFill || c is A.SolidFill || c is A.GradientFill || c is A.BlipFill || c is A.PatternFill || c is A.GroupFill);
            before.InsertAfterSelf(outline);
        }
        return outline;
    }

    private static readonly (AutoShapeType Type, string Preset, string Name)[] Presets =
    {
        (AutoShapeType.Rectangle, "rect", "Rectangle"),
        (AutoShapeType.Diamond, "diamond", "Diamond"),
        (AutoShapeType.RoundedRectangle, "roundRect", "Rectangle: Rounded Corners"),
        (AutoShapeType.IsoscelesTriangle, "triangle", "Isosceles Triangle"),
        (AutoShapeType.Oval, "ellipse", "Oval"),
        (AutoShapeType.Hexagon, "hexagon", "Hexagon"),
        (AutoShapeType.RightArrow, "rightArrow", "Arrow: Right"),
        (AutoShapeType.Parallelogram, "parallelogram", "Parallelogram"),
        (AutoShapeType.Trapezoid, "trapezoid", "Trapezoid"),
        (AutoShapeType.Octagon, "octagon", "Octagon"),
        (AutoShapeType.RightTriangle, "rtTriangle", "Right Triangle"),
        (AutoShapeType.Cross, "plus", "Cross"),
        (AutoShapeType.Can, "can", "Cylinder"),
        (AutoShapeType.Cube, "cube", "Cube"),
        (AutoShapeType.SmileyFace, "smileyFace", "Smiley Face"),
        (AutoShapeType.Donut, "donut", "Circle: Hollow"),
        (AutoShapeType.NoSymbol, "noSmoking", "\"Not Allowed\" Symbol"),
        (AutoShapeType.Heart, "heart", "Heart"),
        (AutoShapeType.LightningBolt, "lightningBolt", "Lightning Bolt"),
        (AutoShapeType.Sun, "sun", "Sun"),
        (AutoShapeType.Moon, "moon", "Moon"),
        (AutoShapeType.LeftArrow, "leftArrow", "Arrow: Left"),
        (AutoShapeType.UpArrow, "upArrow", "Arrow: Up"),
        (AutoShapeType.DownArrow, "downArrow", "Arrow: Down"),
        (AutoShapeType.LeftRightArrow, "leftRightArrow", "Arrow: Left-Right"),
        (AutoShapeType.Pentagon, "homePlate", "Arrow: Pentagon"),
        (AutoShapeType.FlowchartProcess, "flowChartProcess", "Flowchart: Process"),
        (AutoShapeType.FlowchartDecision, "flowChartDecision", "Flowchart: Decision"),
        (AutoShapeType.Star5, "star5", "Star: 5 Points"),
        (AutoShapeType.Cloud, "cloud", "Cloud")
    };

    internal static (AutoShapeType Type, string Preset, string Name) PresetOf(AutoShapeType type)
    {
        if (!Presets.Any(p => p.Type == type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }
        return Presets.First(p => p.Type == type);
    }

    private const string Namespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
        + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    private const string BodyProperties = "<wps:bodyPr rot=\"0\" spcFirstLastPara=\"0\" vertOverflow=\"overflow\" horzOverflow=\"overflow\" vert=\"horz\" wrap=\"square\" "
        + "lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\" numCol=\"1\" spcCol=\"0\" rtlCol=\"0\" fromWordArt=\"0\" anchor=\"{0}\" anchorCtr=\"0\" "
        + "forceAA=\"0\" compatLnSpc=\"1\"><a:prstTxWarp prst=\"textNoShape\"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr>";

    private const string FilledStyle = "<wps:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"15000\"/></a:schemeClr></a:lnRef>"
        + "<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef>"
        + "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></wps:style>";

    private const string LineStyle = "<wps:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>"
        + "<a:effectRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef></wps:style>";

    /// <summary>The floating form of a drawn shape, in front of the text, as Word 16 writes it (without the old copy for earlier versions).</summary>
    internal static Drawing NewDrawn(string kind, string preset, string name, uint id, long leftEmu, long topEmu, long cx, long cy, uint relativeHeight)
    {
        string shape;
        if (kind == "textbox")
        {
            shape = "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{0}\" cy=\"{1}\"/></a:xfrm>"
                + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/><a:ln w=\"6350\"><a:solidFill><a:prstClr val=\"black\"/></a:solidFill></a:ln></wps:spPr>"
                + "<wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>" + BodyProperties.Replace("{0}", "t") + "</wps:wsp>";
        }
        else if (kind == "line")
        {
            shape = "<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{0}\" cy=\"{1}\"/></a:xfrm>"
                + "<a:prstGeom prst=\"line\"><a:avLst/></a:prstGeom></wps:spPr>" + LineStyle + "<wps:bodyPr/></wps:wsp>";
        }
        else
        {
            shape = "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{0}\" cy=\"{1}\"/></a:xfrm>"
                + "<a:prstGeom prst=\"" + preset + "\"><a:avLst/></a:prstGeom></wps:spPr>" + FilledStyle + BodyProperties.Replace("{0}", "ctr") + "</wps:wsp>";
        }
        shape = string.Format(CultureInfo.InvariantCulture, shape, cx, cy);

        string xml = string.Format(
            CultureInfo.InvariantCulture,
            "<w:drawing {0}><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"{1}\" behindDoc=\"0\" locked=\"0\" "
            + "layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>{2}</wp:posOffset></wp:positionH>"
            + "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>{3}</wp:posOffset></wp:positionV><wp:extent cx=\"{4}\" cy=\"{5}\"/>"
            + "<wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapNone/><wp:docPr id=\"{6}\" name=\"{7}\"/><wp:cNvGraphicFramePr/>"
            + "<a:graphic><a:graphicData uri=\"{8}\">{9}</a:graphicData></a:graphic></wp:anchor></w:drawing>",
            Namespaces,
            relativeHeight,
            leftEmu,
            topEmu,
            cx,
            cy,
            id,
            System.Security.SecurityElement.Escape(name),
            ShapeUri,
            shape);
        return new Drawing(xml);
    }
}

public partial class ShapeWrapperCollection
{
    /// <summary>
    /// Adds a text box, as <c>Shapes.AddTextbox</c> in Word, in front of the text, anchored at the start of the paragraph of the anchor. It
    /// has an outline of half a point and no fill, and one empty paragraph; put text in <see cref="ShapeWrapper.TextFrame"/>.
    /// </summary>
    /// <param name="orientation">The direction of the text.</param>
    /// <param name="left">The distance in points from the left edge of the page.</param>
    /// <param name="top">The distance in points from the top edge of the page.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <param name="anchor">Where to anchor the shape; the start of the range this collection is for when null.</param>
    /// <remarks>Word measures the position from the page and stores it from the column and the paragraph. The library takes the paragraph to be
    /// at the top margin and the column at the left margin, which they are for the first paragraph of a page with one column.</remarks>
    public ShapeWrapper AddTextbox(TextOrientation orientation, double left, double top, double width, double height, RangeWrapper anchor = null)
    {
        if (!EnumCompat.IsDefined(orientation))
        {
            throw new ArgumentOutOfRangeException(nameof(orientation));
        }
        ShapeWrapper box = AddDrawn("textbox", "rect", "Text Box", left, top, width, height, anchor);
        if (orientation != TextOrientation.Horizontal)
        {
            box.TextFrame.Orientation = orientation;
        }
        return box;
    }

    /// <summary>
    /// Adds a label, as <c>Shapes.AddLabel</c> in Word: the same text box as <see cref="AddTextbox"/> makes. See there for the position.
    /// </summary>
    public ShapeWrapper AddLabel(TextOrientation orientation, double left, double top, double width, double height, RangeWrapper anchor = null)
    {
        return AddTextbox(orientation, left, top, width, height, anchor);
    }

    /// <summary>
    /// Adds a shape of the gallery, as <c>Shapes.AddShape</c> in Word, in front of the text, anchored at the start of the paragraph of the
    /// anchor. See <see cref="AddTextbox"/> for the position.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The type is not one the library can draw (see <see cref="AutoShapeType"/>).</exception>
    public ShapeWrapper AddShape(AutoShapeType type, double left, double top, double width, double height, RangeWrapper anchor = null)
    {
        (AutoShapeType Type, string Preset, string Name) preset = ShapeWrapper.PresetOf(type);
        return AddDrawn("shape", preset.Preset, preset.Name, left, top, width, height, anchor);
    }

    /// <summary>
    /// Adds a straight line from one point to another, as <c>Shapes.AddLine</c> in Word, anchored at the start of the paragraph of the anchor.
    /// See <see cref="AddTextbox"/> for the position. A line that goes up or to the left is a line that is turned over, as in Word.
    /// </summary>
    public ShapeWrapper AddLine(double beginX, double beginY, double endX, double endY, RangeWrapper anchor = null)
    {
        foreach (double value in new[] { beginX, beginY, endX, endY })
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new ArgumentOutOfRangeException(nameof(beginX));
            }
        }
        ShapeWrapper line = AddDrawn("line", "line", "Straight Connector", Math.Min(beginX, endX), Math.Min(beginY, endY), Math.Abs(endX - beginX), Math.Abs(endY - beginY), anchor);
        if (endX < beginX)
        {
            line.Flip(FlipCommand.Horizontal);
        }
        if (endY < beginY)
        {
            line.Flip(FlipCommand.Vertical);
        }
        return line;
    }

    private ShapeWrapper AddDrawn(string kind, string preset, string name, double left, double top, double width, double height, RangeWrapper anchor)
    {
        foreach (double value in new[] { left, top, width, height })
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new ArgumentOutOfRangeException(nameof(left));
            }
        }
        if (width < 0 || height < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width));
        }

        Paragraph paragraph = (anchor ?? _range).Duplicate().ParagraphOrFirst();
        RangeWrapper at = new RangeWrapper(paragraph);
        at.Collapse(CollapseDirection.Start);

        double leftMargin = 72;
        double topMargin = 72;
        PageSetupWrapper setup = at.Sections.Count > 0 ? at.Sections[0].PageSetup : null;
        if (setup != null)
        {
            leftMargin = setup.LeftMargin;
            topMargin = setup.TopMargin;
        }

        uint relativeHeight = NextRelativeHeight(_range.AnchorElement);
        Drawing drawing = null;
        at.InsertNoteReference(target =>
        {
            MainDocumentPart main = StyleHelper.GetMainDocumentPart(target);
            if (main == null)
            {
                throw new NotSupportedException("A shape can be inserted only in a document.");
            }
            uint id = main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<WpDrawing.DocProperties>())
                .Select(d => d.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
            drawing = ShapeWrapper.NewDrawn(kind, preset, name + " " + id.ToString(CultureInfo.InvariantCulture), id,
                ToEmu(left - leftMargin), ToEmu(top - topMargin), ToEmu(width), ToEmu(height), relativeHeight);
            return new List<Run> { new Run(new RunProperties(new NoProof()), drawing) };
        });
        return new ShapeWrapper(drawing);
    }

    private static long ToEmu(double points)
    {
        return (long)Math.Round(points * 12700);
    }
}
