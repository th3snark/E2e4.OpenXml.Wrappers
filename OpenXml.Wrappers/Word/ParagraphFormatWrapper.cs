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
    Distribute = 4
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
/// <para>Each property reads the formatting that applies to every paragraph that the range touches, as Word does (the document defaults, the
/// table style, the paragraph style and the direct formatting), and is null when they differ. Setting a property sets the direct
/// formatting; setting it to null removes that so that the style applies again. Use <see cref="Direct"/> to read only the direct formatting.
/// The conditional formatting of table styles and the indents of numbering levels are not taken into account.</para>
/// <para>As in Word, measurements are in points. (<see cref="ParagraphWrapper.SpaceBefore"/> and <see cref="ParagraphWrapper.SpaceAfter"/> are
/// the raw twentieths of a point.)</para>
/// </remarks>
public class ParagraphFormatWrapper
{
    private const double MaximumPoints = 1584;

    private readonly RangeWrapper _range;
    private readonly Style _style;
    private readonly bool _direct;

    public ParagraphFormatWrapper(RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        _range = range;
    }

    //The paragraph format of a style: the paragraph properties of the style with those it inherits from the styles it is based on and the
    //document defaults.
    internal ParagraphFormatWrapper(Style style)
    {
        Guard.ThrowIfNull(style);
        _style = style;
    }

    private ParagraphFormatWrapper(RangeWrapper range, Style style, bool direct)
    {
        _range = range;
        _style = style;
        _direct = direct;
    }

    /// <summary>
    /// The same paragraph format, but reading only the formatting that has been set directly on the paragraphs (or in the style itself), not
    /// what comes from styles. A property that has not been set reads as 0, false or <see cref="ParagraphAlignment.Left"/>.
    /// </summary>
    public ParagraphFormatWrapper Direct
    {
        get { return _direct ? this : new ParagraphFormatWrapper(_range, _style, true); }
    }

    //The properties of a style are a different type to those of a paragraph, with the same children. They are edited as a copy of those of
    //a paragraph and written back.
    private ParagraphProperties CopyStyleProperties()
    {
        ParagraphProperties copy = new ParagraphProperties();
        if (_style.StyleParagraphProperties != null)
        {
            foreach (OpenXmlElement child in _style.StyleParagraphProperties.ChildElements)
            {
                copy.AppendChild(child.CloneNode(true));
            }
        }
        return copy;
    }

    private void StoreStyleProperties(ParagraphProperties copy)
    {
        _style.StyleParagraphProperties?.Remove();
        if (copy.HasChildren)
        {
            StyleParagraphProperties properties = new StyleParagraphProperties();
            foreach (OpenXmlElement child in copy.ChildElements.ToList())
            {
                properties.AppendChild(child.CloneNode(true));
            }
            _style.StyleParagraphProperties = properties;
        }
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
                if (value == JustificationValues.Both)
                {
                    return ParagraphAlignment.Justify;
                }
                if (value == JustificationValues.Distribute)
                {
                    return ParagraphAlignment.Distribute;
                }
                return ParagraphAlignment.Left;
            });
        }
        set
        {
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
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
                        : value.Value == ParagraphAlignment.Distribute ? JustificationValues.Distribute
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
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
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

    /// <summary>True to leave out the paragraphs when lines are numbered (<c>ParagraphFormat.LineNumbering</c> suppression in Word).</summary>
    public bool? NoLineNumber
    {
        get { return GetOnOff(properties => properties?.SuppressLineNumbers, e => e.Val); }
        set { SetOnOff<SuppressLineNumbers>(value, flag => new SuppressLineNumbers() { Val = flag }, properties => properties.SuppressLineNumbers = null, (properties, element) => properties.SuppressLineNumbers = element); }
    }

    /// <summary>True to swap the left and right indents on facing pages.</summary>
    public bool? MirrorIndents
    {
        get { return GetOnOff(properties => properties?.MirrorIndents, e => e.Val); }
        set { SetOnOff<MirrorIndents>(value, flag => new MirrorIndents() { Val = flag }, properties => properties.MirrorIndents = null, (properties, element) => properties.MirrorIndents = element); }
    }

    /// <summary>
    /// True to hyphenate the paragraphs when the document is set to hyphenate automatically; false switches it off for them. True by default.
    /// </summary>
    public bool? Hyphenation
    {
        get
        {
            return GetValue<bool?>(properties =>
            {
                SuppressAutoHyphens suppress = properties?.SuppressAutoHyphens;
                return suppress == null || (suppress.Val != null && suppress.Val.HasValue && !suppress.Val.Value);
            });
        }
        set
        {
            SetOnOff<SuppressAutoHyphens>(value.HasValue ? !value : (bool?)null, flag => new SuppressAutoHyphens() { Val = flag }, properties => properties.SuppressAutoHyphens = null, (properties, element) => properties.SuppressAutoHyphens = element);
        }
    }

    /// <summary>True to put the space before the paragraphs automatically, as Word does for web pages (<c>ParagraphFormat.SpaceBeforeAuto</c>).</summary>
    public bool? SpaceBeforeAuto
    {
        get { return GetValue<bool?>(properties => IsOn(properties?.SpacingBetweenLines?.BeforeAutoSpacing)); }
        set
        {
            Apply(properties =>
            {
                if (value.HasValue)
                {
                    GetOrAddSpacing(properties).BeforeAutoSpacing = value.Value;
                }
                else if (properties.SpacingBetweenLines != null)
                {
                    properties.SpacingBetweenLines.BeforeAutoSpacing = null;
                    RemoveIfEmpty(properties.SpacingBetweenLines);
                }
            });
        }
    }

    /// <summary>True to put the space after the paragraphs automatically (<c>ParagraphFormat.SpaceAfterAuto</c>).</summary>
    public bool? SpaceAfterAuto
    {
        get { return GetValue<bool?>(properties => IsOn(properties?.SpacingBetweenLines?.AfterAutoSpacing)); }
        set
        {
            Apply(properties =>
            {
                if (value.HasValue)
                {
                    GetOrAddSpacing(properties).AfterAutoSpacing = value.Value;
                }
                else if (properties.SpacingBetweenLines != null)
                {
                    properties.SpacingBetweenLines.AfterAutoSpacing = null;
                    RemoveIfEmpty(properties.SpacingBetweenLines);
                }
            });
        }
    }

    /// <summary>The outline level of the paragraphs, for headings and the outline view.</summary>
    public ParagraphOutlineLevel? OutlineLevel
    {
        get
        {
            return GetValue<ParagraphOutlineLevel?>(properties =>
            {
                int? level = properties?.OutlineLevel?.Val?.Value;
                return level.HasValue && level.Value >= 0 && level.Value <= 8 ? (ParagraphOutlineLevel)(level.Value + 1) : ParagraphOutlineLevel.BodyText;
            });
        }
        set
        {
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            Apply(properties =>
            {
                properties.OutlineLevel = null;
                if (value.HasValue)
                {
                    properties.OutlineLevel = new OutlineLevel() { Val = value.Value == ParagraphOutlineLevel.BodyText ? 9 : (int)value.Value - 1 };
                }
            });
        }
    }

    /// <summary>The tab stops of the paragraphs, as <c>ParagraphFormat.TabStops</c> in Word.</summary>
    public TabStopWrapperCollection TabStops
    {
        get { return new TabStopWrapperCollection(this); }
    }

    /// <summary>
    /// Removes the direct paragraph formatting so that the styles apply (<c>ParagraphFormat.Reset</c> in Word). The style, the section break
    /// and the formatting of the paragraph mark are kept.
    /// </summary>
    public void Reset()
    {
        Apply(properties =>
        {
            foreach (OpenXmlElement child in properties.ChildElements.Where(e => !(e is ParagraphStyleId || e is SectionProperties || e is ParagraphMarkRunProperties || e is ParagraphPropertiesChange)).ToList())
            {
                child.Remove();
            }
        });
    }

    //Whether a tab stop at the position comes from a style or the document defaults rather than being set on the paragraphs.
    internal bool IsInherited(long twips)
    {
        if (_style != null)
        {
            ParagraphProperties inherited = new FormattingResolver(_style).ResolveStyleParagraph(_style);
            return inherited.Tabs != null && inherited.Tabs.Elements<TabStop>().Any(t => t.Position != null && t.Position.HasValue && t.Position.Value == twips && !TabStopWrapper.IsClear(t));
        }

        IList<Paragraph> paragraphs = GetParagraphs();
        if (paragraphs.Count == 0)
        {
            return false;
        }
        FormattingResolver resolver = new FormattingResolver(paragraphs[0]);
        return paragraphs.Any(p =>
        {
            ParagraphProperties inherited = resolver.ResolveParagraph(p, false);
            return inherited.Tabs != null && inherited.Tabs.Elements<TabStop>().Any(t => t.Position != null && t.Position.HasValue && t.Position.Value == twips && !TabStopWrapper.IsClear(t));
        });
    }

    private static bool IsOn(OnOffValue value)
    {
        return value != null && value.HasValue && value.Value;
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
    internal T GetValue<T>(Func<ParagraphProperties, T> read)
    {
        return GetValue(read, EqualityComparer<T>.Default);
    }

    internal T GetValue<T>(Func<ParagraphProperties, T> read, IEqualityComparer<T> comparer)
    {
        if (_style != null)
        {
            return read(_direct ? CopyStyleProperties() : new FormattingResolver(_style).ResolveStyleParagraph(_style));
        }

        IList<Paragraph> paragraphs = GetParagraphs();
        if (paragraphs.Count == 0)
        {
            return default(T);
        }

        FormattingResolver resolver = _direct ? null : new FormattingResolver(paragraphs[0]);
        Func<Paragraph, ParagraphProperties> properties = p => _direct ? p.ParagraphProperties : resolver.ResolveParagraph(p);
        T first = read(properties(paragraphs[0]));
        return paragraphs.Skip(1).All(p => comparer.Equals(read(properties(p)), first)) ? first : default(T);
    }

    internal void Apply(Action<ParagraphProperties> change)
    {
        if (_style != null)
        {
            ParagraphProperties copy = CopyStyleProperties();
            change(copy);
            StoreStyleProperties(copy);
            return;
        }

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
