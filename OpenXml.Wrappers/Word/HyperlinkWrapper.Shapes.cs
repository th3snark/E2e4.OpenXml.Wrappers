using System;

using DocumentFormat.OpenXml.Drawing;
using WpDrawing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class HyperlinkWrapperCollection
{
    /// <summary>
    /// Makes a shape or picture a hyperlink, as <c>Hyperlinks.Add</c> in Word with a shape as the anchor: the address, with the place added to it
    /// after a <c>#</c>, and the text shown when the pointer rests on the shape. A link the shape has already is replaced.
    /// </summary>
    /// <param name="anchor">The shape.</param>
    /// <param name="address">The address to go to.</param>
    /// <param name="subAddress">The place in the address to go to.</param>
    /// <param name="screenTip">The text to show when the pointer rests on the shape.</param>
    /// <exception cref="ArgumentException">There is no address.</exception>
    public HyperlinkWrapper Add(ShapeWrapper anchor, string address, string subAddress = null, string screenTip = null)
    {
        Guard.ThrowIfNull(anchor);
        if (string.IsNullOrEmpty(address))
        {
            throw new ArgumentException("A shape needs an address to go to.", nameof(address));
        }
        HyperlinkWrapper link = new HyperlinkWrapper(anchor.DocumentProperties);
        link.Address = address;
        link.SubAddress = subAddress;
        link.ScreenTip = screenTip;
        return link;
    }
}

public partial class ShapeWrapper
{
    internal WpDrawing.DocProperties DocumentProperties
    {
        get { return DocProperties; }
    }

    /// <summary>
    /// The hyperlink of the shape, as <c>Shape.Hyperlink</c> in Word; null if it has none. Make one with
    /// <see cref="HyperlinkWrapperCollection.Add(ShapeWrapper, string, string, string)"/>.
    /// </summary>
    public HyperlinkWrapper Hyperlink
    {
        get { return DocProperties.GetFirstChild<HyperlinkOnClick>() == null ? null : new HyperlinkWrapper(DocProperties); }
    }
}
