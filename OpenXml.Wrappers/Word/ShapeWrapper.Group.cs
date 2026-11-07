using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using WpDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The shapes inside a group, as <c>GroupItems</c> of a <c>Shape</c> in Word. They are shapes of the gallery, text boxes and lines; each has a
/// name, an identifier and a size (in the scale of the group), but no place of its own: a group item is not anchored, wrapped or stacked.
/// </summary>
public class GroupShapesCollection : IEnumerable<ShapeWrapper>
{
    private readonly ShapeWrapper _group;

    internal GroupShapesCollection(ShapeWrapper group)
    {
        _group = group;
    }

    private List<ShapeWrapper> GetItems()
    {
        return _group.GroupChildren().Select(c => new ShapeWrapper(_group.Element, c)).ToList();
    }

    /// <summary>The group.</summary>
    public ShapeWrapper Parent
    {
        get { return _group; }
    }

    /// <summary>The number of shapes in the group.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The shape at the index (0-based), as <c>GroupShapes.Item</c> in Word.</summary>
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

public partial class ShapeWrapper
{
    private readonly Wps.WordprocessingShape _child;

    //A shape inside a group: it has no anchor of its own, only the properties Word gives it inside the group.
    internal ShapeWrapper(Drawing group, Wps.WordprocessingShape child)
        : this(group)
    {
        _child = child;
    }

    internal bool IsGroup
    {
        get { return _child == null && _drawing.GetFirstChild<WpDrawing.Anchor>()?.Descendants<Wpg.WordprocessingGroup>().Any() == true; }
    }

    internal Wpg.WordprocessingGroup GroupElement
    {
        get { return AnchorElement.Descendants<Wpg.WordprocessingGroup>().FirstOrDefault(); }
    }

    internal IEnumerable<Wps.WordprocessingShape> GroupChildren()
    {
        return GroupElement?.Elements<Wps.WordprocessingShape>() ?? Enumerable.Empty<Wps.WordprocessingShape>();
    }

    private void ThrowIfChild()
    {
        if (_child != null)
        {
            throw new NotSupportedException("A shape in a group has no place, wrap or stacking of its own.");
        }
    }

    private void ThrowIfGroup()
    {
        if (IsGroup)
        {
            throw new NotSupportedException("A group cannot be turned, flipped or have text; do it to the shapes in it.");
        }
    }

    /// <summary>True if the shape is in a group, as <c>Shape.Child</c> in Word.</summary>
    public bool Child
    {
        get { return _child != null; }
    }

    /// <summary>The group the shape is in, as <c>Shape.ParentGroup</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The shape is not in a group.</exception>
    public ShapeWrapper ParentGroup
    {
        get { return _child == null ? throw new InvalidOperationException("The shape is not in a group.") : new ShapeWrapper(_drawing); }
    }

    /// <summary>The shapes in the group, as <c>Shape.GroupItems</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The shape is not a group.</exception>
    public GroupShapesCollection GroupItems
    {
        get { return IsGroup ? new GroupShapesCollection(this) : throw new InvalidOperationException("The shape is not a group."); }
    }

    //The scale of the shapes in the group: how many times larger than they are drawn in the group they appear.
    private (double X, double Y) GroupScale()
    {
        A.TransformGroup transform = GroupElement.Descendants<A.TransformGroup>().FirstOrDefault();
        long chx = transform?.ChildExtents?.Cx?.Value ?? 0;
        long chy = transform?.ChildExtents?.Cy?.Value ?? 0;
        return (chx == 0 ? 1 : (double)Cx / chx, chy == 0 ? 1 : (double)Cy / chy);
    }

    /// <summary>
    /// Takes the shapes out of the group, as <c>Shape.Ungroup</c> in Word, and returns them in the order they were in. They keep their names, take
    /// the places they had on the page, and stack in the place of the group, the last in the group on top.
    /// </summary>
    /// <exception cref="InvalidOperationException">The shape is not a group.</exception>
    public IList<ShapeWrapper> Ungroup()
    {
        if (!IsGroup)
        {
            throw new InvalidOperationException("The shape is not a group.");
        }
        Run run = _drawing.Ancestors<Run>().First();
        WpDrawing.Anchor anchor = AnchorElement;
        List<Wps.WordprocessingShape> children = GroupChildren().ToList();
        long left = ReadOffset(anchor.HorizontalPosition);
        long top = ReadOffset(anchor.VerticalPosition);
        (double X, double Y) scale = GroupScale();
        uint topHeight = anchor.RelativeHeight?.Value ?? 0U;
        uint next = NextDocPrId(run);

        List<ShapeWrapper> result = new List<ShapeWrapper>();
        OpenXmlElement before = run;
        for (int i = 0; i < children.Count; i++)
        {
            Wps.WordprocessingShape child = children[i];
            Wps.NonVisualDrawingProperties properties = child.GetFirstChild<Wps.NonVisualDrawingProperties>();
            A.Transform2D transform = child.GetFirstChild<Wps.ShapeProperties>().GetFirstChild<A.Transform2D>();
            long x = transform.Offset?.X?.Value ?? 0;
            long y = transform.Offset?.Y?.Value ?? 0;
            long cx = (long)Math.Round((transform.Extents?.Cx?.Value ?? 0) * scale.X);
            long cy = (long)Math.Round((transform.Extents?.Cy?.Value ?? 0) * scale.Y);

            Wps.WordprocessingShape shape = (Wps.WordprocessingShape)child.CloneNode(true);
            shape.GetFirstChild<Wps.NonVisualDrawingProperties>()?.Remove();
            A.Transform2D clonedTransform = shape.GetFirstChild<Wps.ShapeProperties>().GetFirstChild<A.Transform2D>();
            clonedTransform.Offset = new A.Offset() { X = 0L, Y = 0L };
            clonedTransform.Extents = new A.Extents() { Cx = cx, Cy = cy };

            string name = properties?.Name?.Value ?? "Shape";
            uint height = topHeight - 1024U * (uint)(children.Count - 1 - i);
            Drawing drawing = NewDrawn("shape", "rect", name, next++, left + (long)Math.Round(x * scale.X), top + (long)Math.Round(y * scale.Y), cx, cy, height);
            WpDrawing.Anchor placed = drawing.GetFirstChild<WpDrawing.Anchor>();
            placed.HorizontalPosition.RelativeFrom = anchor.HorizontalPosition.RelativeFrom;
            placed.VerticalPosition.RelativeFrom = anchor.VerticalPosition.RelativeFrom;
            A.GraphicData data = placed.Descendants<A.GraphicData>().First();
            data.RemoveAllChildren();
            data.AppendChild(shape);

            Run copy = new Run(new RunProperties(new NoProof()), drawing);
            before.InsertAfterSelf(copy);
            before = copy;
            result.Add(new ShapeWrapper(drawing));
        }
        run.Remove();
        return result;
    }

    private static long ReadOffset(WpDrawing.HorizontalPosition position)
    {
        return position?.PositionOffset?.Text is string text && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0;
    }

    private static long ReadOffset(WpDrawing.VerticalPosition position)
    {
        return position?.PositionOffset?.Text is string text && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0;
    }

    internal static uint NextDocPrId(OpenXmlElement context)
    {
        DocumentFormat.OpenXml.Packaging.MainDocumentPart main = StyleHelper.GetMainDocumentPart(context);
        return main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<WpDrawing.DocProperties>())
            .Select(d => d.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;
    }
}

public partial class ShapeWrapperCollection
{
    private const long GroupLeftMargin = 2540;

    /// <summary>
    /// Puts shapes in a group, as <c>ShapeRange.Group</c> in Word, and returns the group. The group takes the place of the shapes, keeps their
    /// order, stacks where the topmost of them was, and is named like a new shape (Group 3). The shapes must be shapes of the gallery, text
    /// boxes or lines anchored in the same paragraph and placed from the same thing. Word starts the group 0.2 point to the left of the
    /// leftmost shape, which moves with it.
    /// </summary>
    /// <param name="shapes">The shapes to group, at least two.</param>
    /// <exception cref="ArgumentException">There are fewer than two shapes.</exception>
    /// <exception cref="NotSupportedException">A shape is a picture or a group, or the shapes are not placed alike.</exception>
    public ShapeWrapper Group(IEnumerable<ShapeWrapper> shapes)
    {
        Guard.ThrowIfNull(shapes);
        List<ShapeWrapper> list = shapes.ToList();
        if (list.Count < 2)
        {
            throw new ArgumentException("A group needs at least two shapes.", nameof(shapes));
        }
        if (list.Any(s => s.Child || s.IsGroup || s.Shape == null))
        {
            throw new NotSupportedException("Only shapes of the gallery, text boxes and lines can be grouped.");
        }
        list = list.OrderBy(s => s.AnchorElement.RelativeHeight?.Value ?? 0U).ToList();
        WpDrawing.Anchor first = list[0].AnchorElement;
        foreach (ShapeWrapper shape in list)
        {
            WpDrawing.Anchor each = shape.AnchorElement;
            if (each.HorizontalPosition?.RelativeFrom?.Value != first.HorizontalPosition?.RelativeFrom?.Value
                || each.VerticalPosition?.RelativeFrom?.Value != first.VerticalPosition?.RelativeFrom?.Value
                || !ReferenceEquals(shape.Element.Ancestors<Paragraph>().FirstOrDefault(), list[0].Element.Ancestors<Paragraph>().FirstOrDefault()))
            {
                throw new NotSupportedException("The shapes must be anchored in the same paragraph and placed from the same thing.");
            }
        }

        long[] lefts = list.Select(s => ReadOffsetH(s.AnchorElement)).ToArray();
        long[] tops = list.Select(s => ReadOffsetV(s.AnchorElement)).ToArray();
        long[] widths = list.Select(s => s.AnchorElement.Extent?.Cx?.Value ?? 0).ToArray();
        long[] heights = list.Select(s => s.AnchorElement.Extent?.Cy?.Value ?? 0).ToArray();
        long groupLeft = lefts.Min() - GroupLeftMargin;
        long groupTop = tops.Min();
        long cx = Enumerable.Range(0, list.Count).Max(i => lefts[i] + widths[i]) - groupLeft;
        long cy = Enumerable.Range(0, list.Count).Max(i => tops[i] + heights[i]) - groupTop;

        Run firstRun = list.Select(s => s.Element.Ancestors<Run>().First()).OrderBy(r => r.Parent == null ? 0 : r.Parent.ChildElements.ToList().IndexOf(r)).First();
        uint id = ShapeWrapper.NextDocPrId(firstRun);
        uint relativeHeight = list[list.Count - 1].AnchorElement.RelativeHeight?.Value ?? 0U;
        Random random = new Random();

        Drawing drawing = ShapeWrapper.NewDrawn("shape", "rect", "Group " + id.ToString(CultureInfo.InvariantCulture), id, groupLeft, groupTop, cx, cy, relativeHeight);
        WpDrawing.Anchor anchor = drawing.GetFirstChild<WpDrawing.Anchor>();
        anchor.HorizontalPosition.RelativeFrom = first.HorizontalPosition.RelativeFrom;
        anchor.VerticalPosition.RelativeFrom = first.VerticalPosition.RelativeFrom;
        A.GraphicData data = anchor.Descendants<A.GraphicData>().First();
        data.Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
        data.RemoveAllChildren();
        Wpg.WordprocessingGroup group = new Wpg.WordprocessingGroup(
            new Wpg.NonVisualGroupDrawingShapeProperties(),
            new Wpg.GroupShapeProperties(new A.TransformGroup(
                new A.Offset() { X = 0L, Y = 0L },
                new A.Extents() { Cx = cx, Cy = cy },
                new A.ChildOffset() { X = 0L, Y = 0L },
                new A.ChildExtents() { Cx = cx, Cy = cy })));
        for (int i = 0; i < list.Count; i++)
        {
            Wps.WordprocessingShape child = (Wps.WordprocessingShape)list[i].Shape.CloneNode(true);
            WpDrawing.DocProperties properties = list[i].DocumentProperties;
            Wps.NonVisualDrawingProperties named = new Wps.NonVisualDrawingProperties() { Id = (uint)random.Next(1, int.MaxValue), Name = properties.Name?.Value };
            if (properties.Description != null)
            {
                named.Description = properties.Description.Value;
            }
            if (properties.Title != null)
            {
                named.Title = properties.Title.Value;
            }
            child.PrependChild(named);
            A.Transform2D transform = child.GetFirstChild<Wps.ShapeProperties>().GetFirstChild<A.Transform2D>();
            //The leftmost shapes start at the left of the group, which is 0.2 point left of where they were: they move with it.
            transform.Offset = new A.Offset() { X = lefts[i] == lefts.Min() ? 0L : lefts[i] - groupLeft, Y = tops[i] - groupTop };
            transform.Extents = new A.Extents() { Cx = widths[i], Cy = heights[i] };
            group.AppendChild(child);
        }
        data.AppendChild(group);

        Run run = new Run(new RunProperties(new NoProof()), drawing);
        firstRun.InsertBeforeSelf(run);
        foreach (ShapeWrapper shape in list)
        {
            shape.Element.Ancestors<Run>().First().Remove();
        }
        return new ShapeWrapper(drawing);
    }

    private static long ReadOffsetH(WpDrawing.Anchor anchor)
    {
        return anchor.HorizontalPosition?.PositionOffset?.Text is string text && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0;
    }

    private static long ReadOffsetV(WpDrawing.Anchor anchor)
    {
        return anchor.VerticalPosition?.PositionOffset?.Text is string text && long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long emu) ? emu : 0;
    }
}
