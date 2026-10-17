using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The alignment of a paragraph. The values are those of Word's WdParagraphAlignment.
/// </summary>
public enum ParagraphAlignment
{
    Left = 0,
    Center = 1,
    Right = 2,
    Justify = 3,
}

/// <summary>
/// The rule used for the spacing between the lines of a paragraph. The values are those of Word's WdLineSpacing.
/// </summary>
public enum LineSpacingRule
{
    /// <summary>Single line spacing.</summary>
    SingleSpacing = 0,

    /// <summary>One and a half lines.</summary>
    OnePointFive = 1,

    /// <summary>Double line spacing.</summary>
    DoubleSpacing = 2,

    /// <summary>At least the line spacing, which is in points.</summary>
    AtLeast = 3,

    /// <summary>Exactly the line spacing, which is in points.</summary>
    Exactly = 4,

    /// <summary>A number of lines. <see cref="ParagraphFormatWrapper.LineSpacing"/> is that number times 12 points.</summary>
    Multiple = 5,
}

/// <summary>
/// The paragraph formatting of a range, as <c>Range.ParagraphFormat</c> in Word.
/// </summary>
/// <remarks>
/// <para>Each property reads and sets the direct formatting of every paragraph that the range touches, and is null when they differ.
/// Formatting that comes from styles is not taken into account, so a property that has not been set reads as 0, false or
/// <see cref="ParagraphAlignment.Left"/>. Setting a property to null removes the direct formatting so that the style applies again.</para>
/// <para>As in Word, measurements are in points. (<see cref="ParagraphWrapper.SpaceBefore"/> and <see cref="ParagraphWrapper.SpaceAfter"/> are
/// the raw twentieths of a point.)</para>
/// </remarks>
public class ParagraphFormatWrapper
{
    private const double MaximumPoints = 1584;

    private readonly RangeWrapper _range;

    public ParagraphFormatWrapper(RangeWrapper range)
    {
        ArgumentNullException.ThrowIfNull(range);
        _range = range;
    }

    /// <summary>The alignment of the paragraphs.</summary>
    public ParagraphAlignment? Alignment
    {
        get
        {
            return GetValue<ParagraphAlignment?>(properties =>
            {
                JustificationValues? value = properties?.Justification?.Val?.Value;
                if (value == JustificationValues.Center)
                {
                    return ParagraphAlignment.Center;
                }
                if (value == JustificationValues.Right || value == JustificationValues.End)
                {
                    return ParagraphAlignment.Right;
                }
                if (value == JustificationValues.Both || value == JustificationValues.Distribute)
                {
                    return ParagraphAlignment.Justify;
                }
                return ParagraphAlignment.Left;
            });
        }
        set
        {
            if (value.HasValue && !Enum.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            Apply(properties =>
            {
                properties.RemoveAllChildren<Justification>();
                if (value.HasValue)
                {
                    JustificationValues justification = value.Value == ParagraphAlignment.Center ? JustificationValues.Center
                        : value.Value == ParagraphAlignment.Right ? JustificationValues.Right
                        : value.Value == ParagraphAlignment.Justify ? JustificationValues.Both
                        : JustificationValues.Left;
                    properties.Justification = new Justification() { Val = justification };
                }
            });
        }
    }

    /// <summary>The distance in points between the left margin and the left edge of the paragraphs.</summary>
    public double? LeftIndent
    {
        get { return GetValue<double?>(properties => ToPoints(properties?.Indentation?.Left?.Value ?? properties?.Indentation?.Start?.Value) ?? 0); }
        set
        {
            Points(value, -MaximumPoints, nameof(value));
            Apply(properties =>
            {
                Indentation indentation = GetOrAddIndentation(properties);
                indentation.Start = null;
                if (value.HasValue)
                {
                    indentation.Left = ToTwips(value.Value);
                }
                else
                {
                    indentation.Left = null;
                }
                RemoveIfEmpty(indentation);
            });
        }
    }

    /// <summary>The distance in points between the right margin and the right edge of the paragraphs.</summary>
    public double? RightIndent
    {
        get { return GetValue<double?>(properties => ToPoints(properties?.Indentation?.Right?.Value ?? properties?.Indentation?.End?.Value) ?? 0); }
        set
        {
            Points(value, -MaximumPoints, nameof(value));
            Apply(properties =>
            {
                Indentation indentation = GetOrAddIndentation(properties);
                indentation.End = null;
                if (value.HasValue)
                {
                    indentation.Right = ToTwips(value.Value);
                }
                else
                {
                    indentation.Right = null;
                }
                RemoveIfEmpty(indentation);
            });
        }
    }

    /// <summary>
    /// The indent in points of the first line relative to the left indent. A negative value is a hanging indent, where the first line
    /// starts to the left of the other lines.
    /// </summary>
    public double? FirstLineIndent
    {
        get
        {
            return GetValue<double?>(properties =>
            {
                Indentation indentation = properties?.Indentation;
                double hanging = ToPoints(indentation?.Hanging?.Value) ?? 0;
                return hanging != 0 ? -hanging : ToPoints(indentation?.FirstLine?.Value) ?? 0;
            });
        }
        set
        {
            Points(value, -MaximumPoints, nameof(value));
            Apply(properties =>
            {
                Indentation indentation = GetOrAddIndentation(properties);
                indentation.FirstLine = null;
                indentation.Hanging = null;
                if (value.HasValue)
                {
                    //A hanging indent is stored as a positive amount in its own attribute.
                    if (value.Value < 0)
                    {
                        indentation.Hanging = ToTwips(-value.Value);
                    }
                    else
                    {
                        indentation.FirstLine = ToTwips(value.Value);
                    }
                }
                RemoveIfEmpty(indentation);
            });
        }
    }

    /// <summary>The space in points before the paragraphs.</summary>
    public double? SpaceBefore
    {
        get { return GetValue<double?>(properties => ToPoints(properties?.SpacingBetweenLines?.Before?.Value) ?? 0); }
        set
        {
            Points(value, 0, nameof(value));
            Apply(properties =>
            {
                SpacingBetweenLines spacing = GetOrAddSpacing(properties);
                spacing.BeforeLines = null;
                spacing.BeforeAutoSpacing = null;
                if (value.HasValue)
                {
                    spacing.Before = ToTwips(value.Value);
                }
                else
                {
                    spacing.Before = null;
                }
                RemoveIfEmpty(spacing);
            });
        }
    }

    /// <summary>The space in points after the paragraphs.</summary>
    public double? SpaceAfter
    {
        get { return GetValue<double?>(properties => ToPoints(properties?.SpacingBetweenLines?.After?.Value) ?? 0); }
        set
        {
            Points(value, 0, nameof(value));
            Apply(properties =>
            {
                SpacingBetweenLines spacing = GetOrAddSpacing(properties);
                spacing.AfterLines = null;
                spacing.AfterAutoSpacing = null;
                if (value.HasValue)
                {
                    spacing.After = ToTwips(value.Value);
                }
                else
                {
                    spacing.After = null;
                }
                RemoveIfEmpty(spacing);
            });
        }
    }

    /// <summary>
    /// The line spacing in points. For <see cref="LineSpacingRule.Multiple"/> and the single, one and a half and double rules this is the number of
    /// lines times 12. Setting it makes the rule <see cref="LineSpacingRule.Multiple"/> unless the rule is already
    /// <see cref="LineSpacingRule.AtLeast"/> or <see cref="LineSpacingRule.Exactly"/>.
    /// </summary>
    public double? LineSpacing
    {
        get { return GetValue<double?>(properties => ToPoints(properties?.SpacingBetweenLines?.Line?.Value) ?? 12); }
        set
        {
            Points(value, 0, nameof(value));
            if (value.HasValue && value.Value == 0)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The line spacing must be greater than 0.");
            }

            Apply(properties =>
            {
                SpacingBetweenLines spacing = GetOrAddSpacing(properties);
                if (!value.HasValue)
                {
                    spacing.Line = null;
                    spacing.LineRule = null;
                }
                else
                {
                    spacing.Line = ToTwips(value.Value);
                    if (spacing.LineRule == null || !spacing.LineRule.HasValue || spacing.LineRule.Value == LineSpacingRuleValues.Auto)
                    {
                        spacing.LineRule = LineSpacingRuleValues.Auto;
                    }
                }
                RemoveIfEmpty(spacing);
            });
        }
    }

    /// <summary>The rule used for the line spacing.</summary>
    public LineSpacingRule? LineSpacingRule
    {
        get
        {
            return GetValue<LineSpacingRule?>(properties =>
            {
                SpacingBetweenLines spacing = properties?.SpacingBetweenLines;
                LineSpacingRuleValues? rule = spacing?.LineRule?.Value;
                if (rule == LineSpacingRuleValues.AtLeast)
                {
                    return Word.LineSpacingRule.AtLeast;
                }
                if (rule == LineSpacingRuleValues.Exact)
                {
                    return Word.LineSpacingRule.Exactly;
                }

                double line = ToPoints(spacing?.Line?.Value) ?? 12;
                return line == 12 ? Word.LineSpacingRule.SingleSpacing
                    : line == 18 ? Word.LineSpacingRule.OnePointFive
                    : line == 24 ? Word.LineSpacingRule.DoubleSpacing
                    : Word.LineSpacingRule.Multiple;
            });
        }
        set
        {
            if (value.HasValue && !Enum.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            Apply(properties =>
            {
                SpacingBetweenLines spacing = GetOrAddSpacing(properties);
                if (!value.HasValue)
                {
                    spacing.Line = null;
                    spacing.LineRule = null;
                }
                else if (value.Value == Word.LineSpacingRule.AtLeast || value.Value == Word.LineSpacingRule.Exactly)
                {
                    spacing.LineRule = value.Value == Word.LineSpacingRule.AtLeast ? LineSpacingRuleValues.AtLeast : LineSpacingRuleValues.Exact;
                    spacing.Line = spacing.Line ?? ToTwips(12);
                }
                else
                {
                    //The number of lines is kept when the rule stays a multiple, otherwise it is the number of lines of the rule.
                    bool wasAuto = spacing.LineRule == null || !spacing.LineRule.HasValue || spacing.LineRule.Value == LineSpacingRuleValues.Auto;
                    spacing.LineRule = LineSpacingRuleValues.Auto;
                    spacing.Line = value.Value == Word.LineSpacingRule.SingleSpacing ? "240"
                        : value.Value == Word.LineSpacingRule.OnePointFive ? "360"
                        : value.Value == Word.LineSpacingRule.DoubleSpacing ? "480"
                        : wasAuto && spacing.Line != null ? spacing.Line.Value : "240";
                }
                RemoveIfEmpty(spacing);
            });
        }
    }

    /// <summary>True to keep the paragraphs on the same page as the paragraph that follows.</summary>
    public bool? KeepWithNext
    {
        get { return GetOnOff(properties => properties?.KeepNext, e => e.Val); }
        set { SetOnOff<KeepNext>(value, flag => new KeepNext() { Val = flag }, properties => properties.KeepNext = null, (properties, element) => properties.KeepNext = element); }
    }

    /// <summary>True to keep all the lines of the paragraphs on the same page.</summary>
    public bool? KeepTogether
    {
        get { return GetOnOff(properties => properties?.KeepLines, e => e.Val); }
        set { SetOnOff<KeepLines>(value, flag => new KeepLines() { Val = flag }, properties => properties.KeepLines = null, (properties, element) => properties.KeepLines = element); }
    }

    /// <summary>True to start the paragraphs on a new page.</summary>
    public bool? PageBreakBefore
    {
        get { return GetOnOff(properties => properties?.PageBreakBefore, e => e.Val); }
        set { SetOnOff<PageBreakBefore>(value, flag => new PageBreakBefore() { Val = flag }, properties => properties.PageBreakBefore = null, (properties, element) => properties.PageBreakBefore = element); }
    }

    /// <summary>True to stop the first or last line of the paragraphs being left alone on a page.</summary>
    public bool? WidowControl
    {
        get { return GetOnOff(properties => properties?.WidowControl, e => e.Val); }
        set { SetOnOff<WidowControl>(value, flag => new WidowControl() { Val = flag }, properties => properties.WidowControl = null, (properties, element) => properties.WidowControl = element); }
    }

    private bool? GetOnOff<T>(Func<ParagraphProperties, T> find, Func<T, OnOffValue> getValue) where T : OpenXmlElement
    {
        return GetValue<bool?>(properties =>
        {
            T element = find(properties);
            if (element == null)
            {
                return false;
            }

            OnOffValue value = getValue(element);
            return value == null || !value.HasValue || value.Value;
        });
    }

    private void SetOnOff<T>(bool? value, Func<bool, T> create, Action<ParagraphProperties> remove, Action<ParagraphProperties, T> set) where T : OpenXmlElement
    {
        Apply(properties =>
        {
            remove(properties);
            if (value.HasValue)
            {
                set(properties, create(value.Value));
            }
        });
    }

    private static void Points(double? value, double minimum, string parameterName)
    {
        if (value.HasValue && (Double.IsNaN(value.Value) || value.Value < minimum || value.Value > MaximumPoints))
        {
            throw new ArgumentOutOfRangeException(parameterName, String.Format(CultureInfo.InvariantCulture, "The value must be between {0} and {1} points.", minimum, MaximumPoints));
        }
    }

    private static string ToTwips(double points)
    {
        return Math.Round(points * 20, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
    }

    private static double? ToPoints(string twips)
    {
        if (twips != null && Double.TryParse(twips, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
        {
            return value / 20;
        }
        return null;
    }

    private static Indentation GetOrAddIndentation(ParagraphProperties properties)
    {
        if (properties.Indentation == null)
        {
            properties.Indentation = new Indentation();
        }
        return properties.Indentation;
    }

    private static SpacingBetweenLines GetOrAddSpacing(ParagraphProperties properties)
    {
        if (properties.SpacingBetweenLines == null)
        {
            properties.SpacingBetweenLines = new SpacingBetweenLines();
        }
        return properties.SpacingBetweenLines;
    }

    private static void RemoveIfEmpty(OpenXmlElement element)
    {
        if (!element.HasAttributes)
        {
            element.Remove();
        }
    }

    private IList<Paragraph> GetParagraphs()
    {
        return _range.Paragraphs.Select(p => p.Paragraph).ToList();
    }

    /// <summary>
    /// Reads a value from the paragraph properties of every paragraph the range touches and returns it if they all agree, otherwise null.
    /// </summary>
    private T GetValue<T>(Func<ParagraphProperties, T> read)
    {
        IList<Paragraph> paragraphs = GetParagraphs();
        if (paragraphs.Count == 0)
        {
            return default(T);
        }

        T first = read(paragraphs[0].ParagraphProperties);
        return paragraphs.Skip(1).All(p => EqualityComparer<T>.Default.Equals(read(p.ParagraphProperties), first)) ? first : default(T);
    }

    private void Apply(Action<ParagraphProperties> change)
    {
        foreach (Paragraph paragraph in GetParagraphs())
        {
            ParagraphProperties properties = paragraph.GetOrAddParagraphProperties();
            change(properties);
            if (!properties.HasChildren && !properties.HasAttributes)
            {
                properties.Remove();
            }
        }
    }
}
