using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The character formatting of a range, as <c>Range.Font</c> in Word.
/// </summary>
/// <remarks>
/// <para>Each property reads the formatting that applies to the runs of the range, as Word does: the document defaults, the table style, the
/// paragraph style, the character style and the direct formatting. Setting a property sets the direct formatting. For a range of whole
/// paragraphs (a paragraph, table cell, row, table or the document) it also covers the formatting of the paragraph marks. Use
/// <see cref="Direct"/> to read only the direct formatting. The conditional formatting of table styles (such as the first row), numbering and
/// the tints and shades of theme colours are not taken into account.</para>
/// <para>A property is null when the runs of the range do not all have the same value. Setting a property to null removes the direct
/// formatting so that the style applies again. The characters of a range that starts or ends in the middle of a run are formatted by
/// splitting the run.</para>
/// </remarks>
public class FontWrapper
{

    private readonly RangeWrapper _range;
    private readonly Style _style;
    private readonly bool _direct;

    public FontWrapper(RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        _range = range;
    }

    //The font of a style: the run properties of the style with those it inherits from the styles it is based on and the document defaults.
    internal FontWrapper(Style style)
    {
        Guard.ThrowIfNull(style);
        _style = style;
    }

    private FontWrapper(RangeWrapper range, Style style, bool direct)
    {
        _range = range;
        _style = style;
        _direct = direct;
    }

    /// <summary>
    /// The same font, but reading only the formatting that has been set directly on the runs (or in the style itself), not what comes from
    /// styles. A property that has not been set reads as null, except for the on and off properties such as <see cref="Bold"/>, which read false.
    /// </summary>
    public FontWrapper Direct
    {
        get { return _direct ? this : new FontWrapper(_range, _style, true); }
    }

    /// <summary>Bold text.</summary>
    public bool? Bold
    {
        get { return GetOnOff<Bold>(e => e.Val); }
        set { SetOnOff("b", "bCs", value, flag => new Bold() { Val = flag }, flag => new BoldComplexScript() { Val = flag }); }
    }

    /// <summary>Italic text.</summary>
    public bool? Italic
    {
        get { return GetOnOff<Italic>(e => e.Val); }
        set { SetOnOff("i", "iCs", value, flag => new Italic() { Val = flag }, flag => new ItalicComplexScript() { Val = flag }); }
    }

    /// <summary>Struck through text.</summary>
    public bool? StrikeThrough
    {
        get { return GetOnOff<Strike>(e => e.Val); }
        set { SetOnOff("strike", null, value, flag => new Strike() { Val = flag }, null); }
    }

    /// <summary>The underline of the text, as <c>Font.Underline</c> in Word. <see cref="UnderlineStyle.None"/> when there is none.</summary>
    public UnderlineStyle? Underline
    {
        get
        {
            return GetValue<UnderlineStyle?>(properties =>
            {
                Underline underline = properties?.Elements<Underline>().FirstOrDefault();
                if (underline == null)
                {
                    return UnderlineStyle.None;
                }
                return underline.Val == null || !underline.Val.HasValue ? UnderlineStyle.SingleLine : ToUnderlineStyle(underline.Val.Value);
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
                //Keep the colour of the underline when only its style changes.
                string color = properties.Elements<Underline>().FirstOrDefault()?.Color?.Value;
                properties.RemoveRunProperty("u");
                if (value.HasValue)
                {
                    Underline underline = new Underline() { Val = ToUnderlineValues(value.Value) };
                    if (color != null && value.Value != UnderlineStyle.None)
                    {
                        underline.Color = color;
                    }
                    properties.SetRunProperty(underline);
                }
            });
        }
    }

    /// <summary>
    /// The colour of the underline (as <c>Font.UnderlineColor</c> in Word); null when the text differs. Setting it on text
    /// that is not underlined gives a single underline.
    /// </summary>
    public WordColor? UnderlineColor
    {
        get
        {
            return GetValue<WordColor?>(properties =>
            {
                Underline underline = properties?.Elements<Underline>().FirstOrDefault();
                string stored = underline?.Color?.Value;
                return stored == null ? (_direct ? (WordColor?)null : WordColor.Automatic) : WordColor.FromStored(stored);
            });
        }
        set
        {
            Apply(properties =>
            {
                Underline underline = properties.Elements<Underline>().FirstOrDefault();
                if (underline == null && value == null)
                {
                    return;
                }
                if (underline == null)
                {
                    underline = new Underline() { Val = UnderlineValues.Single };
                    properties.SetRunProperty(underline);
                }
                underline.Color = value.HasValue ? new DocumentFormat.OpenXml.StringValue(value.Value.Hex) : null;
            });
        }
    }

    /// <summary>Superscript text. Setting it to true makes <see cref="Subscript"/> false.</summary>
    public bool? Superscript
    {
        get { return GetVerticalAlignment(VerticalPositionValues.Superscript); }
        set { SetVerticalAlignment(value, VerticalPositionValues.Superscript); }
    }

    /// <summary>Subscript text. Setting it to true makes <see cref="Superscript"/> false.</summary>
    public bool? Subscript
    {
        get { return GetVerticalAlignment(VerticalPositionValues.Subscript); }
        set { SetVerticalAlignment(value, VerticalPositionValues.Subscript); }
    }

    /// <summary>Text in capital letters.</summary>
    public bool? AllCaps
    {
        get { return GetOnOff<Caps>(e => e.Val); }
        set { SetOnOff("caps", null, value, flag => new Caps() { Val = flag }, null); }
    }

    /// <summary>Small capitals.</summary>
    public bool? SmallCaps
    {
        get { return GetOnOff<SmallCaps>(e => e.Val); }
        set { SetOnOff("smallCaps", null, value, flag => new SmallCaps() { Val = flag }, null); }
    }

    /// <summary>Hidden text.</summary>
    public bool? Hidden
    {
        get { return GetOnOff<Vanish>(e => e.Val); }
        set { SetOnOff("vanish", null, value, flag => new Vanish() { Val = flag }, null); }
    }

    /// <summary>Text with a double line through it. Setting it to true makes <see cref="StrikeThrough"/> false.</summary>
    public bool? DoubleStrikeThrough
    {
        get { return GetOnOff<DoubleStrike>(e => e.Val); }
        set
        {
            if (value == true)
            {
                Apply(properties => properties.RemoveRunProperty("strike"));
            }
            SetOnOff("dstrike", null, value, flag => new DoubleStrike() { Val = flag }, null);
        }
    }

    /// <summary>Embossed text.</summary>
    public bool? Emboss
    {
        get { return GetOnOff<Emboss>(e => e.Val); }
        set { SetOnOff("emboss", null, value, flag => new Emboss() { Val = flag }, null); }
    }

    /// <summary>Engraved text.</summary>
    public bool? Engrave
    {
        get { return GetOnOff<Imprint>(e => e.Val); }
        set { SetOnOff("imprint", null, value, flag => new Imprint() { Val = flag }, null); }
    }

    /// <summary>Text drawn as an outline.</summary>
    public bool? Outline
    {
        get { return GetOnOff<Outline>(e => e.Val); }
        set { SetOnOff("outline", null, value, flag => new Outline() { Val = flag }, null); }
    }

    /// <summary>Text with a shadow.</summary>
    public bool? Shadow
    {
        get { return GetOnOff<Shadow>(e => e.Val); }
        set { SetOnOff("shadow", null, value, flag => new Shadow() { Val = flag }, null); }
    }

    /// <summary>The space between the characters in points; negative to condense them (<c>Font.Spacing</c> in Word).</summary>
    public double? Spacing
    {
        get
        {
            return GetValue<double?>(properties =>
            {
                long? twips = properties?.Elements<Spacing>().FirstOrDefault()?.Val?.Value;
                return twips.HasValue ? twips.Value / 20.0 : 0;
            });
        }
        set
        {
            if (value.HasValue && (Double.IsNaN(value.Value) || Math.Abs(value.Value) > 1584))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The spacing must be between -1584 and 1584 points.");
            }

            Apply(properties =>
            {
                properties.RemoveRunProperty("spacing");
                if (value.HasValue && value.Value != 0)
                {
                    properties.SetRunProperty(new Spacing() { Val = (int)Math.Round(value.Value * 20, MidpointRounding.AwayFromZero) });
                }
            });
        }
    }

    /// <summary>The width of the characters as a percentage of their normal width, 100 by default (<c>Font.Scaling</c> in Word).</summary>
    public int? Scaling
    {
        get
        {
            return GetValue<int?>(properties =>
            {
                long? scale = properties?.Elements<CharacterScale>().FirstOrDefault()?.Val?.Value;
                return (int)(scale ?? 100);
            });
        }
        set
        {
            if (value.HasValue && (value.Value < 1 || value.Value > 600))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The scaling must be between 1 and 600 percent.");
            }

            Apply(properties =>
            {
                properties.RemoveRunProperty("w");
                if (value.HasValue && value.Value != 100)
                {
                    properties.SetRunProperty(new CharacterScale() { Val = value.Value });
                }
            });
        }
    }

    /// <summary>The distance in points that the text is raised (positive) or lowered (negative) from the baseline (<c>Font.Position</c> in Word).</summary>
    public double? Position
    {
        get
        {
            return GetValue<double?>(properties =>
            {
                string value = properties?.Elements<Position>().FirstOrDefault()?.Val?.Value;
                return Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int halfPoints) ? halfPoints / 2.0 : 0;
            });
        }
        set
        {
            if (value.HasValue && (Double.IsNaN(value.Value) || Math.Abs(value.Value) > 1584))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The position must be between -1584 and 1584 points.");
            }

            Apply(properties =>
            {
                properties.RemoveRunProperty("position");
                if (value.HasValue && value.Value != 0)
                {
                    properties.SetRunProperty(new Position() { Val = ((int)Math.Round(value.Value * 2, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture) });
                }
            });
        }
    }

    /// <summary>The smallest font size in points that is kerned, 0 when the characters are not kerned (<c>Font.Kerning</c> in Word).</summary>
    public double? Kerning
    {
        get
        {
            return GetValue<double?>(properties =>
            {
                uint? halfPoints = properties?.Elements<Kern>().FirstOrDefault()?.Val?.Value;
                return halfPoints.HasValue ? halfPoints.Value / 2.0 : 0;
            });
        }
        set
        {
            if (value.HasValue && (Double.IsNaN(value.Value) || value.Value < 0 || value.Value > 1638))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The kerning must be between 0 and 1638 points.");
            }

            Apply(properties =>
            {
                properties.RemoveRunProperty("kern");
                if (value.HasValue && value.Value != 0)
                {
                    properties.SetRunProperty(new Kern() { Val = (uint)Math.Round(value.Value * 2, MidpointRounding.AwayFromZero) });
                }
            });
        }
    }

    /// <summary>
    /// Removes the direct character formatting, so that the styles apply (<c>Font.Reset</c> in Word). A character style that has been applied
    /// is kept.
    /// </summary>
    public void Reset()
    {
        Apply(properties =>
        {
            foreach (OpenXmlElement child in properties.ChildElements.Where(e => !(e is RunStyle)).ToList())
            {
                child.Remove();
            }
        });
    }

    private bool? GetVerticalAlignment(VerticalPositionValues position)
    {
        return GetValue<bool?>(properties =>
        {
            VerticalTextAlignment alignment = properties?.Elements<VerticalTextAlignment>().FirstOrDefault();
            return alignment?.Val != null && alignment.Val.HasValue && alignment.Val.Value == position;
        });
    }

    private void SetVerticalAlignment(bool? value, VerticalPositionValues position)
    {
        Apply(properties =>
        {
            VerticalTextAlignment existing = properties.Elements<VerticalTextAlignment>().FirstOrDefault();
            bool isThis = existing?.Val != null && existing.Val.HasValue && existing.Val.Value == position;
            if (value == true)
            {
                properties.SetRunProperty(new VerticalTextAlignment() { Val = position });
            }
            else if (isThis)
            {
                //Setting false (or null) only removes this kind of alignment; the opposite one stays.
                properties.RemoveRunProperty("vertAlign");
            }
        });
    }

    private static UnderlineStyle ToUnderlineStyle(UnderlineValues value)
    {
        foreach (KeyValuePair<UnderlineStyle, UnderlineValues> pair in UnderlineMap)
        {
            if (pair.Value == value)
            {
                return pair.Key;
            }
        }
        return UnderlineStyle.SingleLine;
    }

    private static UnderlineValues ToUnderlineValues(UnderlineStyle style)
    {
        foreach (KeyValuePair<UnderlineStyle, UnderlineValues> pair in UnderlineMap)
        {
            if (pair.Key == style)
            {
                return pair.Value;
            }
        }
        return UnderlineValues.Single;
    }

    //The Open XML SDK 3.x values are structs, so they are looked up in a list rather than a switch.
    private static readonly List<KeyValuePair<UnderlineStyle, UnderlineValues>> UnderlineMap = new List<KeyValuePair<UnderlineStyle, UnderlineValues>>
    {
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.None, UnderlineValues.None),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.SingleLine, UnderlineValues.Single),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.Words, UnderlineValues.Words),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DoubleLine, UnderlineValues.Double),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.Dotted, UnderlineValues.Dotted),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.Thick, UnderlineValues.Thick),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.Dash, UnderlineValues.Dash),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DotDash, UnderlineValues.DotDash),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DotDotDash, UnderlineValues.DotDotDash),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.Wavy, UnderlineValues.Wave),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DottedHeavy, UnderlineValues.DottedHeavy),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DashHeavy, UnderlineValues.DashedHeavy),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DotDashHeavy, UnderlineValues.DashDotHeavy),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DotDotDashHeavy, UnderlineValues.DashDotDotHeavy),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.WavyHeavy, UnderlineValues.WavyHeavy),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DashLong, UnderlineValues.DashLong),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.WavyDouble, UnderlineValues.WavyDouble),
        new KeyValuePair<UnderlineStyle, UnderlineValues>(UnderlineStyle.DashLongHeavy, UnderlineValues.DashLongHeavy)
    };

    /// <summary>The size of the characters in points, which may be a half point such as 10.5.</summary>
    public double? Size
    {
        get
        {
            return GetValue<double?>(properties =>
            {
                FontSize fontSize = properties?.Elements<FontSize>().FirstOrDefault();
                if (fontSize?.Val == null)
                {
                    //Word uses 10 points when nothing says otherwise.
                    return _direct ? (double?)null : 10;
                }
                return Convert.ToDouble(fontSize.Val.Value, CultureInfo.InvariantCulture) / 2;
            });
        }
        set
        {
            if (value.HasValue && (value.Value <= 0 || value.Value > 1638))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The font size must be greater than 0 and no more than 1638 points.");
            }

            Apply(properties =>
            {
                properties.RemoveRunProperty("sz");
                properties.RemoveRunProperty("szCs");
                if (value.HasValue)
                {
                    //The size is stored in half points.
                    string halfPoints = Math.Round(value.Value * 2, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
                    properties.SetRunProperty(new FontSize() { Val = halfPoints });
                    properties.SetRunProperty(new FontSizeComplexScript() { Val = halfPoints });
                }
            });
        }
    }

    /// <summary>The name of the font.</summary>
    public string Name
    {
        get
        {
            return GetValue((properties, resolver) =>
            {
                RunFonts fonts = properties?.Elements<RunFonts>().FirstOrDefault();
                if (_direct)
                {
                    return fonts?.Ascii?.Value;
                }

                string themeName = null;
                ThemeFontValues? theme = fonts?.AsciiTheme?.Value;
                if (theme.HasValue)
                {
                    themeName = resolver.GetThemeFontName(theme.Value);
                }
                //Word uses Times New Roman when nothing says otherwise.
                return themeName ?? fonts?.Ascii?.Value ?? "Times New Roman";
            });
        }
        set
        {
            Apply(properties =>
            {
                properties.RemoveRunProperty("rFonts");
                if (value != null)
                {
                    properties.SetRunProperty(new RunFonts() { Ascii = value, HighAnsi = value, EastAsia = value, ComplexScript = value });
                }
            });
        }
    }

    /// <summary>
    /// The colour of the text (<c>Font.Color</c> in Word, which takes <c>RGB</c> values and <c>WdColor</c> constants): assign a six digit
    /// hexadecimal string such as <c>"FF0000"</c>, a value from <see cref="ColorHelper.RGB"/> or a <see cref="NamedColor"/>. Null when the text
    /// differs; setting null removes the direct colour.
    /// </summary>
    public WordColor? Color
    {
        get
        {
            return GetValue<WordColor?>((properties, resolver) =>
            {
                Color color = properties?.Elements<Color>().FirstOrDefault();
                string value = color?.Val?.Value;
                if (_direct)
                {
                    return value == null ? (WordColor?)null : WordColor.FromStored(value);
                }

                if (value == null || String.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    ThemeColorValues? theme = color?.ThemeColor?.Value;
                    value = theme.HasValue ? resolver.GetThemeColorHex(theme.Value) : null;
                }
                return WordColor.FromStored(value);
            });
        }
        set
        {
            Apply(properties =>
            {
                properties.RemoveRunProperty("color");
                if (value.HasValue)
                {
                    properties.SetRunProperty(new Color() { Val = value.Value.Hex });
                }
            });
        }
    }

    /// <summary>
    /// The colour of the text as one of the sixteen colours of Word's colour index (<c>Font.ColorIndex</c> in Word), the nearest to the
    /// colour of the text; null when the text differs. Setting it gives the text that colour, and <see cref="ColorIndex.Auto"/> the automatic colour.
    /// </summary>
    public ColorIndex? ColorIndex
    {
        get { return Color?.ToColorIndex(); }
        set { Color = value.HasValue ? WordColor.FromColorIndex(value.Value) : (WordColor?)null; }
    }

    private bool? GetOnOff<T>(Func<T, OnOffValue> getValue) where T : OpenXmlElement
    {
        return GetValue<bool?>(properties =>
        {
            T element = properties?.ChildElements.OfType<T>().FirstOrDefault();
            if (element == null)
            {
                return false;
            }

            OnOffValue value = getValue(element);
            return value == null || !value.HasValue || value.Value;
        });
    }

    private void SetOnOff(string localName, string complexScriptLocalName, bool? value, Func<bool, OpenXmlElement> create, Func<bool, OpenXmlElement> createComplexScript)
    {
        Apply(properties =>
        {
            properties.RemoveRunProperty(localName);
            if (complexScriptLocalName != null)
            {
                properties.RemoveRunProperty(complexScriptLocalName);
            }

            if (value.HasValue)
            {
                properties.SetRunProperty(create(value.Value));
                if (createComplexScript != null)
                {
                    properties.SetRunProperty(createComplexScript(value.Value));
                }
            }
        });
    }

    /// <summary>
    /// Reads a value from the run properties of every run of the range (or from the paragraph marks if the range has no runs) and returns it if
    /// they all agree, otherwise null.
    /// </summary>
    private T GetValue<T>(Func<OpenXmlCompositeElement, T> read)
    {
        return GetValue((properties, resolver) => read(properties));
    }

    private T GetValue<T>(Func<OpenXmlCompositeElement, FormattingResolver, T> read)
    {
        if (_style != null)
        {
            if (_direct)
            {
                return read(_style.StyleRunProperties, null);
            }

            FormattingResolver styleResolver = new FormattingResolver(_style);
            return read(styleResolver.ResolveStyleRun(_style), styleResolver);
        }

        IList<Run> runs = _range.GetRuns().ToList();
        IList<Paragraph> marks = runs.Count == 0 ? _range.GetParagraphs().ToList() : new List<Paragraph>();
        if (runs.Count == 0 && marks.Count == 0)
        {
            return default(T);
        }

        FormattingResolver resolver = _direct ? null : new FormattingResolver(runs.Count > 0 ? (OpenXmlElement)runs[0] : marks[0]);
        IList<OpenXmlCompositeElement> targets = runs.Count > 0
            ? runs.Select(r => _direct ? (OpenXmlCompositeElement)r.RunProperties : resolver.ResolveRun(r)).ToList()
            : marks.Select(p => _direct ? (OpenXmlCompositeElement)p.ParagraphProperties?.ParagraphMarkRunProperties : resolver.ResolveParagraphMark(p)).ToList();

        T first = read(targets[0], resolver);
        return targets.Skip(1).All(t => EqualityComparer<T>.Default.Equals(read(t, resolver), first)) ? first : default(T);
    }

    /// <summary>
    /// Applies a change to the run properties of every run of the range and, where the range covers whole paragraphs, to the paragraph marks.
    /// </summary>
    private void Apply(Action<OpenXmlCompositeElement> change)
    {
        if (_style != null)
        {
            StyleRunProperties styleProperties = _style.StyleRunProperties;
            if (styleProperties == null)
            {
                styleProperties = new StyleRunProperties();
                _style.StyleRunProperties = styleProperties;
            }
            change(styleProperties);
            if (!styleProperties.HasChildren)
            {
                styleProperties.Remove();
            }
            return;
        }

        foreach (Run run in _range.GetRuns())
        {
            RunProperties runProperties = run.GetOrAddRunProperties();
            change(runProperties);
            if (!runProperties.HasChildren)
            {
                runProperties.Remove();
            }
        }

        foreach (Paragraph paragraph in _range.GetParagraphs())
        {
            ParagraphMarkRunProperties markProperties = paragraph.GetOrAddParagraphMarkRunProperties();
            change(markProperties);
            if (!markProperties.HasChildren)
            {
                markProperties.Remove();
            }
        }
    }
}
