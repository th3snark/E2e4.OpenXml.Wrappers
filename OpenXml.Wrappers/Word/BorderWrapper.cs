using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// One border of a table, table row or table cell, as <c>Border</c> in Word.
/// </summary>
/// <remarks>
/// <para>Reading a border that has not been set gives <see cref="BorderLineStyle.None"/>, a width of 0 and the colour "auto". A border takes
/// its real appearance from the table style when it is not set directly.</para>
/// <para>A border of a row is the border of each of the cells it covers, so it can have mixed values: then <see cref="LineStyle"/> reads
/// <see cref="BorderLineStyle.Other"/>, <see cref="LineWidth"/> reads 0 and <see cref="Color"/> reads null. Setting a value sets it on
/// all of the cells.</para>
/// </remarks>
public class BorderWrapper
{

    private readonly IList<(BordersWrapper Borders, BorderPosition Position)> _targets;

    internal BorderWrapper(BordersWrapper borders, BorderPosition position)
        : this(new[] { (borders, position) })
    {
    }

    internal BorderWrapper(IList<(BordersWrapper Borders, BorderPosition Position)> targets)
    {
        _targets = targets;
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
            List<BorderLineStyle> styles = _targets.Select(t => GetLineStyle(t.Borders, t.Position)).Distinct().ToList();
            if (styles.Count == 0)
            {
                return BorderLineStyle.None;
            }
            return styles.Count == 1 ? styles[0] : BorderLineStyle.Other;
        }
        set
        {
            if (value == BorderLineStyle.Other)
            {
                throw new ArgumentException("The style Other cannot be set.", nameof(value));
            }
            foreach ((BordersWrapper borders, BorderPosition position) in _targets)
            {
                bool created = borders.Find(position) == null;
                BorderType border = borders.GetOrAdd(position);
                border.Val = ToValue(value);
                if (created && value != BorderLineStyle.None)
                {
                    //The defaults of Word: half a point, automatic colour.
                    border.Size = 4;
                    border.Color = "auto";
                    border.Space = borders.DefaultSpace;
                }
            }
        }
    }

    /// <summary>
    /// The width of the line (<c>Border.LineWidth</c> in Word): a number of points from 0.25 to 12 (rounded to eighths of a point) or a
    /// <see cref="BorderLineWidth"/>. A border that has not been set, or borders that differ, read as a width of 0.
    /// </summary>
    public LineWidth LineWidth
    {
        get
        {
            List<LineWidth> widths = _targets.Select(t => GetLineWidth(t.Borders, t.Position)).Distinct().ToList();
            return widths.Count == 1 ? widths[0] : default(LineWidth);
        }
        set
        {
            if (value.Points == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The width must be at least 0.25 points.");
            }
            foreach ((BordersWrapper borders, BorderPosition position) in _targets)
            {
                BorderType border = borders.Find(position);
                if (border == null)
                {
                    //Setting a width on a border without a style gives it the default style, as Word does.
                    border = borders.GetOrAdd(position);
                    border.Val = BorderValues.Single;
                    border.Color = "auto";
                    border.Space = borders.DefaultSpace;
                }
                border.Size = value.EighthPoints;
            }
        }
    }

    /// <summary>
    /// The colour (<c>Border.Color</c> in Word, which takes <c>RGB</c> values and <c>WdColor</c> constants); null when the borders differ.
    /// </summary>
    /// <exception cref="ArgumentNullException">The colour is null.</exception>
    public WordColor? Color
    {
        get
        {
            List<WordColor> colors = _targets.Select(t => WordColor.FromStored(t.Borders.FindEffective(t.Position)?.Color?.Value)).Distinct().ToList();
            if (colors.Count == 0)
            {
                return WordColor.Automatic;
            }
            return colors.Count == 1 ? colors[0] : (WordColor?)null;
        }
        set
        {
            if (!value.HasValue)
            {
                throw new ArgumentNullException(nameof(value));
            }
            string hex = value.Value.Hex;
            foreach ((BordersWrapper borders, BorderPosition position) in _targets)
            {
                BorderType border = borders.Find(position);
                if (border == null)
                {
                    border = borders.GetOrAdd(position);
                    border.Val = BorderValues.Single;
                    border.Size = 4;
                    border.Space = borders.DefaultSpace;
                }
                border.Color = hex;
            }
        }
    }

    /// <summary>
    /// The colour as one of the colours of Word's colour index (<c>Border.ColorIndex</c> in Word), the nearest to the colour of the border; null
    /// when the borders differ.
    /// </summary>
    public ColorIndex? ColorIndex
    {
        get { return Color?.ToColorIndex(); }
        set
        {
            if (!value.HasValue)
            {
                throw new ArgumentNullException(nameof(value));
            }
            Color = WordColor.FromColorIndex(value.Value);
        }
    }

    /// <summary>
    /// Whether the border has a line (<c>Border.Visible</c> in Word). Setting it to true gives a border without a line a single half-point line.
    /// </summary>
    public bool Visible
    {
        get { return LineStyle != BorderLineStyle.None; }
        set
        {
            if (!value)
            {
                LineStyle = BorderLineStyle.None;
            }
            else if (LineStyle == BorderLineStyle.None)
            {
                LineStyle = BorderLineStyle.SingleLine;
            }
        }
    }

    /// <summary>
    /// Removes the direct border setting.
    /// </summary>
    public void Clear()
    {
        foreach ((BordersWrapper borders, BorderPosition position) in _targets)
        {
            borders.Find(position)?.Remove();
        }
    }

    private static BorderLineStyle GetLineStyle(BordersWrapper borders, BorderPosition position)
    {
        BorderType border = borders.FindEffective(position);
        if (border?.Val == null || !border.Val.HasValue)
        {
            return BorderLineStyle.None;
        }
        BorderValues value = border.Val.Value;
        return value == BorderValues.Nil || value == BorderValues.None ? BorderLineStyle.None : Map(value);
    }

    private static LineWidth GetLineWidth(BordersWrapper borders, BorderPosition position)
    {
        BorderType border = borders.FindEffective(position);
        return border?.Size == null || !border.Size.HasValue ? default(LineWidth) : LineWidth.FromEighthPoints(border.Size.Value);
    }

    private static BorderLineStyle Map(BorderValues value)
    {
        foreach (KeyValuePair<BorderLineStyle, BorderValues> pair in Styles)
        {
            if (pair.Value == value)
            {
                return pair.Key;
            }
        }
        return BorderLineStyle.Other;
    }

    private static readonly List<KeyValuePair<BorderLineStyle, BorderValues>> Styles = new List<KeyValuePair<BorderLineStyle, BorderValues>>
    {
        Pair(BorderLineStyle.SingleLine, BorderValues.Single),
        Pair(BorderLineStyle.Dotted, BorderValues.Dotted),
        Pair(BorderLineStyle.DashSmallGap, BorderValues.DashSmallGap),
        Pair(BorderLineStyle.DashLargeGap, BorderValues.Dashed),
        Pair(BorderLineStyle.DashDot, BorderValues.DotDash),
        Pair(BorderLineStyle.DashDotDot, BorderValues.DotDotDash),
        Pair(BorderLineStyle.DoubleLine, BorderValues.Double),
        Pair(BorderLineStyle.Triple, BorderValues.Triple),
        Pair(BorderLineStyle.ThinThickSmallGap, BorderValues.ThinThickSmallGap),
        Pair(BorderLineStyle.ThickThinSmallGap, BorderValues.ThickThinSmallGap),
        Pair(BorderLineStyle.ThinThickThinSmallGap, BorderValues.ThinThickThinSmallGap),
        Pair(BorderLineStyle.ThinThickMediumGap, BorderValues.ThinThickMediumGap),
        Pair(BorderLineStyle.ThickThinMediumGap, BorderValues.ThickThinMediumGap),
        Pair(BorderLineStyle.ThinThickThinMediumGap, BorderValues.ThinThickThinMediumGap),
        Pair(BorderLineStyle.ThinThickLargeGap, BorderValues.ThinThickLargeGap),
        Pair(BorderLineStyle.ThickThinLargeGap, BorderValues.ThickThinLargeGap),
        Pair(BorderLineStyle.ThinThickThinLargeGap, BorderValues.ThinThickThinLargeGap),
        Pair(BorderLineStyle.SingleWavy, BorderValues.Wave),
        Pair(BorderLineStyle.DoubleWavy, BorderValues.DoubleWave),
        Pair(BorderLineStyle.DashDotStroked, BorderValues.DashDotStroked),
        Pair(BorderLineStyle.Emboss3D, BorderValues.ThreeDEmboss),
        Pair(BorderLineStyle.Engrave3D, BorderValues.ThreeDEngrave),
        Pair(BorderLineStyle.Outset, BorderValues.Outset),
        Pair(BorderLineStyle.Inset, BorderValues.Inset)
    };

    private static KeyValuePair<BorderLineStyle, BorderValues> Pair(BorderLineStyle style, BorderValues value)
    {
        return new KeyValuePair<BorderLineStyle, BorderValues>(style, value);
    }

    private static BorderValues ToValue(BorderLineStyle style)
    {
        foreach (KeyValuePair<BorderLineStyle, BorderValues> pair in Styles)
        {
            if (pair.Key == style)
            {
                return pair.Value;
            }
        }
        return BorderValues.Nil;
    }
}
