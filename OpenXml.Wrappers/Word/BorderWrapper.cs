using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// One border of a table or table cell, as <c>Border</c> in Word.
/// </summary>
/// <remarks>
/// Reading a border that has not been set gives <see cref="BorderLineStyle.None"/>, a width of 0 and the colour "auto". A border takes
/// its real appearance from the table style when it is not set directly.
/// </remarks>
public class BorderWrapper
{
    private const double MinimumWidth = 0.25;
    private const double MaximumWidth = 12;

    private readonly BordersWrapper _borders;
    private readonly BorderPosition _position;

    internal BorderWrapper(BordersWrapper borders, BorderPosition position)
    {
        _borders = borders;
        _position = position;
    }

    /// <summary>
    /// The line style. Setting <see cref="BorderLineStyle.None"/> writes an explicit "no border"; use <see cref="Clear"/> to remove the
    /// direct setting so that the style of the table applies.
    /// </summary>
    /// <exception cref="ArgumentException">The style is <see cref="BorderLineStyle.Other"/>.</exception>
    public BorderLineStyle LineStyle
    {
        get
        {
            BorderType border = _borders.Find(_position);
            if (border?.Val == null || !border.Val.HasValue)
            {
                return BorderLineStyle.None;
            }
            BorderValues value = border.Val.Value;
            return value == BorderValues.Nil || value == BorderValues.None ? BorderLineStyle.None : Map(value);
        }
        set
        {
            if (value == BorderLineStyle.Other)
            {
                throw new ArgumentException("The style Other cannot be set.", nameof(value));
            }
            bool created = _borders.Find(_position) == null;
            BorderType border = _borders.GetOrAdd(_position);
            border.Val = ToValue(value);
            if (created && value != BorderLineStyle.None)
            {
                //The defaults of Word: half a point, automatic colour.
                border.Size = 4;
                border.Color = "auto";
                border.Space = 0;
            }
        }
    }

    /// <summary>
    /// The width of the line in points, from 0.25 to 12 (rounded to eighths of a point). A border that has not been set has a width of 0.
    /// </summary>
    public double LineWidth
    {
        get
        {
            BorderType border = _borders.Find(_position);
            return border?.Size == null || !border.Size.HasValue ? 0 : border.Size.Value / 8.0;
        }
        set
        {
            if (Double.IsNaN(value) || value < MinimumWidth || value > MaximumWidth)
            {
                throw new ArgumentOutOfRangeException(nameof(value), String.Format(CultureInfo.InvariantCulture, "The width must be between {0} and {1} points.", MinimumWidth, MaximumWidth));
            }
            BorderType border = _borders.Find(_position);
            if (border == null)
            {
                //Setting a width on a border without a style gives it the default style, as Word does.
                border = _borders.GetOrAdd(_position);
                border.Val = BorderValues.Single;
                border.Color = "auto";
                border.Space = 0;
            }
            border.Size = (uint)Math.Round(value * 8, MidpointRounding.AwayFromZero);
        }
    }

    /// <summary>
    /// The colour as six hex digits, or "auto".
    /// </summary>
    /// <exception cref="ArgumentException">The colour is neither six hex digits nor "auto".</exception>
    public string Color
    {
        get { return _borders.Find(_position)?.Color?.Value ?? "auto"; }
        set
        {
            if (value == null || !(String.Equals(value, "auto", StringComparison.OrdinalIgnoreCase) || (value.Length == 6 && value.All(Uri.IsHexDigit))))
            {
                throw new ArgumentException("The colour must be six hex digits or \"auto\".", nameof(value));
            }
            BorderType border = _borders.Find(_position);
            if (border == null)
            {
                border = _borders.GetOrAdd(_position);
                border.Val = BorderValues.Single;
                border.Size = 4;
                border.Space = 0;
            }
            border.Color = value;
        }
    }

    /// <summary>
    /// Removes the direct border setting.
    /// </summary>
    public void Clear()
    {
        _borders.Find(_position)?.Remove();
    }

    private static BorderLineStyle Map(BorderValues value)
    {
        foreach (BorderLineStyle style in Enum.GetValues<BorderLineStyle>())
        {
            if (style != BorderLineStyle.None && style != BorderLineStyle.Other && ToValue(style) == value)
            {
                return style;
            }
        }
        return BorderLineStyle.Other;
    }

    private static BorderValues ToValue(BorderLineStyle style)
    {
        return style switch
        {
            BorderLineStyle.SingleLine => BorderValues.Single,
            BorderLineStyle.Dotted => BorderValues.Dotted,
            BorderLineStyle.Dashed => BorderValues.Dashed,
            BorderLineStyle.DotDash => BorderValues.DotDash,
            BorderLineStyle.DoubleLine => BorderValues.Double,
            BorderLineStyle.Triple => BorderValues.Triple,
            BorderLineStyle.Wave => BorderValues.Wave,
            BorderLineStyle.Inset => BorderValues.Inset,
            BorderLineStyle.Outset => BorderValues.Outset,
            _ => BorderValues.Nil
        };
    }
}
