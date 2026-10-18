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
    private static readonly Regex ColorPattern = new Regex("^([0-9A-Fa-f]{6}|auto)$", RegexOptions.CultureInvariant);

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

    /// <summary>Underlined text. Setting true gives a single underline and any underline style reads as true.</summary>
    public bool? Underline
    {
        get
        {
            return GetValue<bool?>(properties =>
            {
                Underline underline = properties?.Elements<Underline>().FirstOrDefault();
                return underline != null && underline.Val != null && underline.Val.HasValue && underline.Val.Value != UnderlineValues.None;
            });
        }
        set
        {
            Apply(properties =>
            {
                properties.RemoveRunProperty("u");
                if (value.HasValue)
                {
                    properties.SetRunProperty(new Underline() { Val = value.Value ? UnderlineValues.Single : UnderlineValues.None });
                }
            });
        }
    }

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
    /// The colour of the text as a six digit hexadecimal red, green and blue value such as <c>FF0000</c> for red, or <c>auto</c> for the
    /// automatic colour.
    /// </summary>
    public string Color
    {
        get
        {
            return GetValue((properties, resolver) =>
            {
                Color color = properties?.Elements<Color>().FirstOrDefault();
                string value = color?.Val?.Value;
                if (_direct)
                {
                    return value;
                }

                if (value == null || String.Equals(value, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    ThemeColorValues? theme = color?.ThemeColor?.Value;
                    value = theme.HasValue ? resolver.GetThemeColorHex(theme.Value) : null;
                }
                return value ?? "auto";
            });
        }
        set
        {
            if (value != null && !ColorPattern.IsMatch(value))
            {
                throw new ArgumentException("The colour must be six hexadecimal digits (such as FF0000) or 'auto'.", nameof(value));
            }

            Apply(properties =>
            {
                properties.RemoveRunProperty("color");
                if (value != null)
                {
                    properties.SetRunProperty(new Color() { Val = value });
                }
            });
        }
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
