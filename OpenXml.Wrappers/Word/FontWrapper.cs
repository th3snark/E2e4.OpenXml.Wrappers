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
/// <para>Each property reads and sets the direct formatting of the runs of the range. For a range of whole paragraphs (a paragraph, table cell,
/// row, table or the document) it also covers the formatting of the paragraph marks. Formatting that comes from styles is not taken into account.</para>
/// <para>A property is null when the runs of the range do not all have the same value. Boolean properties are false where no formatting
/// has been set. Setting a property to null removes the direct formatting so that the style applies again.</para>
/// <para>A range is made of whole elements, so unlike in Word the formatting cannot be applied to part of a run. Use <see cref="RangeWrapper.Find"/>
/// to change the text you want to format first, or a bookmark that surrounds whole runs.</para>
/// </remarks>
public class FontWrapper
{
    private static readonly Regex ColorPattern = new Regex("^([0-9A-Fa-f]{6}|auto)$", RegexOptions.CultureInvariant);

    private readonly RangeWrapper _range;
    private readonly Style _style;

    public FontWrapper(RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        _range = range;
    }

    //The font of a style: the run properties of the style itself, without anything it inherits from the style it is based on.
    internal FontWrapper(Style style)
    {
        Guard.ThrowIfNull(style);
        _style = style;
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
    public bool? Strikethrough
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
                return fontSize?.Val == null ? (double?)null : Convert.ToDouble(fontSize.Val.Value, CultureInfo.InvariantCulture) / 2;
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
        get { return GetValue(properties => properties?.Elements<RunFonts>().FirstOrDefault()?.Ascii?.Value); }
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
        get { return GetValue(properties => properties?.Elements<Color>().FirstOrDefault()?.Val?.Value); }
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
        if (_style != null)
        {
            return read(_style.StyleRunProperties);
        }

        IList<OpenXmlCompositeElement> targets = _range.GetRuns().Select(r => (OpenXmlCompositeElement)r.RunProperties).ToList();
        if (targets.Count == 0)
        {
            targets = _range.GetParagraphs().Select(p => (OpenXmlCompositeElement)p.ParagraphProperties?.ParagraphMarkRunProperties).ToList();
        }
        if (targets.Count == 0)
        {
            return default(T);
        }

        T first = read(targets[0]);
        return targets.Skip(1).All(t => EqualityComparer<T>.Default.Equals(read(t), first)) ? first : default(T);
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
