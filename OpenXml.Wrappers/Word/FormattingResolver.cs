using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Works out the formatting that applies to runs and paragraphs once the styles are taken into account, as Word reports it: the document
/// defaults, the table style, the paragraph style, the character style and then the direct formatting, each over the one before.
/// </summary>
/// <remarks>
/// A resolver caches what it finds in the styles, so make one for each read rather than keeping it. It does not take into account the
/// conditional formatting of table styles (first row and so on), numbering levels, or tints and shades of theme colours.
/// </remarks>
internal sealed class FormattingResolver
{
    //The properties that switch on and off. Where styles of different kinds set them they toggle, so a character style that is bold in a
    //paragraph style that is bold gives text that is not.
    private static readonly string[] ToggleNames = { "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "vanish" };

    private static readonly string[] IgnoredNames = { "rStyle", "rPrChange", "pPrChange", "ins", "del", "moveFrom", "moveTo", "pStyle", "rPr", "sectPr", "numPr" };

    private readonly MainDocumentPart _mainDocumentPart;
    private readonly Styles _styles;
    private Dictionary<string, Style> _stylesById;
    private readonly Dictionary<string, RunProperties> _runChains = new Dictionary<string, RunProperties>();
    private readonly Dictionary<string, ParagraphProperties> _paragraphChains = new Dictionary<string, ParagraphProperties>();
    private readonly Dictionary<Style, RunProperties> _styleRuns = new Dictionary<Style, RunProperties>();
    private readonly Dictionary<Style, ParagraphProperties> _styleParagraphs = new Dictionary<Style, ParagraphProperties>();
    private RunProperties _defaultRun;
    private ParagraphProperties _defaultParagraph;

    public FormattingResolver(OpenXmlElement context)
    {
        Guard.ThrowIfNull(context);
        _mainDocumentPart = StyleHelper.GetMainDocumentPart(context);
        _styles = _mainDocumentPart?.StyleDefinitionsPart?.Styles;
    }

    #region Runs

    /// <summary>The formatting of a run.</summary>
    public RunProperties ResolveRun(Run run)
    {
        return ResolveRun(run.RunProperties, run.Ancestors<Paragraph>().FirstOrDefault(), run.RunProperties?.RunStyle?.Val?.Value);
    }

    /// <summary>The formatting of the paragraph mark of a paragraph.</summary>
    public RunProperties ResolveParagraphMark(Paragraph paragraph)
    {
        return ResolveRun(paragraph.ParagraphProperties?.ParagraphMarkRunProperties, paragraph, null);
    }

    /// <summary>The character formatting of a style: the document defaults and the style with the styles it is based on.</summary>
    public RunProperties ResolveStyleRun(Style style)
    {
        if (!_styleRuns.TryGetValue(style, out RunProperties result))
        {
            result = new RunProperties();
            Overlay(result, GetDefaultRun());
            Overlay(result, CollapseRunChain(style.StyleId?.Value));
            _styleRuns[style] = result;
        }
        return result;
    }

    private RunProperties ResolveRun(OpenXmlElement direct, Paragraph paragraph, string characterStyleId)
    {
        RunProperties defaults = GetDefaultRun();
        RunProperties result = new RunProperties();
        Overlay(result, defaults);

        List<RunProperties> levels = new List<RunProperties>();
        string tableStyleId = GetTableStyleId(paragraph);
        if (tableStyleId != null)
        {
            levels.Add(CollapseRunChain(tableStyleId));
        }
        levels.Add(CollapseRunChain(GetParagraphStyleId(paragraph)));
        if (characterStyleId != null)
        {
            levels.Add(CollapseRunChain(characterStyleId));
        }

        foreach (RunProperties level in levels)
        {
            Overlay(result, level, skipToggles: true);
        }

        foreach (string name in ToggleNames)
        {
            bool value = false;
            OpenXmlElement last = null;
            foreach (RunProperties level in levels)
            {
                OpenXmlElement element = FindByName(level, name);
                if (element != null)
                {
                    last = element;
                    if (IsOn(element))
                    {
                        value = !value;
                    }
                }
            }
            if (last != null)
            {
                Remove(result, name);
                OpenXmlElement toggled = last.CloneNode(true);
                if (toggled is OnOffType onOff)
                {
                    onOff.Val = value;
                }
                result.AppendChild(toggled);
            }
        }

        if (direct != null)
        {
            Overlay(result, direct);
        }
        return result;
    }

    private RunProperties CollapseRunChain(string styleId)
    {
        if (styleId == null)
        {
            return new RunProperties();
        }
        if (!_runChains.TryGetValue(styleId, out RunProperties result))
        {
            result = new RunProperties();
            foreach (Style style in GetChain(styleId))
            {
                Overlay(result, style.StyleRunProperties);
            }
            _runChains[styleId] = result;
        }
        return result;
    }

    private RunProperties GetDefaultRun()
    {
        if (_defaultRun == null)
        {
            _defaultRun = new RunProperties();
            Overlay(_defaultRun, _styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle);
        }
        return _defaultRun;
    }

    #endregion

    #region Paragraphs

    /// <summary>The formatting of a paragraph.</summary>
    public ParagraphProperties ResolveParagraph(Paragraph paragraph)
    {
        ParagraphProperties result = new ParagraphProperties();
        Overlay(result, GetDefaultParagraph());
        string tableStyleId = GetTableStyleId(paragraph);
        if (tableStyleId != null)
        {
            Overlay(result, CollapseParagraphChain(tableStyleId));
        }
        Overlay(result, CollapseParagraphChain(GetParagraphStyleId(paragraph)));
        Overlay(result, paragraph.ParagraphProperties);
        return result;
    }

    /// <summary>The paragraph formatting of a style: the document defaults and the style with the styles it is based on.</summary>
    public ParagraphProperties ResolveStyleParagraph(Style style)
    {
        if (!_styleParagraphs.TryGetValue(style, out ParagraphProperties result))
        {
            result = new ParagraphProperties();
            Overlay(result, GetDefaultParagraph());
            Overlay(result, CollapseParagraphChain(style.StyleId?.Value));
            _styleParagraphs[style] = result;
        }
        return result;
    }

    private ParagraphProperties CollapseParagraphChain(string styleId)
    {
        if (styleId == null)
        {
            return new ParagraphProperties();
        }
        if (!_paragraphChains.TryGetValue(styleId, out ParagraphProperties result))
        {
            result = new ParagraphProperties();
            foreach (Style style in GetChain(styleId))
            {
                Overlay(result, style.StyleParagraphProperties);
            }
            _paragraphChains[styleId] = result;
        }
        return result;
    }

    private ParagraphProperties GetDefaultParagraph()
    {
        if (_defaultParagraph == null)
        {
            _defaultParagraph = new ParagraphProperties();
            Overlay(_defaultParagraph, _styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle);
        }
        return _defaultParagraph;
    }

    #endregion

    #region Themes

    /// <summary>The name of the font that a theme font setting such as <c>minorHAnsi</c> refers to, or null.</summary>
    public string GetThemeFontName(ThemeFontValues themeFont)
    {
        DocumentFormat.OpenXml.Drawing.FontScheme scheme = _mainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
        if (scheme == null)
        {
            return null;
        }

        bool major = themeFont == ThemeFontValues.MajorAscii || themeFont == ThemeFontValues.MajorHighAnsi
            || themeFont == ThemeFontValues.MajorEastAsia || themeFont == ThemeFontValues.MajorBidi;
        string typeface = major ? scheme.MajorFont?.LatinFont?.Typeface?.Value : scheme.MinorFont?.LatinFont?.Typeface?.Value;
        return String.IsNullOrEmpty(typeface) ? null : typeface;
    }

    /// <summary>The red, green and blue value of a theme colour, or null. Tints and shades are not applied.</summary>
    public string GetThemeColorHex(ThemeColorValues themeColor)
    {
        DocumentFormat.OpenXml.Drawing.ColorScheme scheme = _mainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        if (scheme == null)
        {
            return null;
        }

        DocumentFormat.OpenXml.Drawing.Color2Type color = null;
        if (themeColor == ThemeColorValues.Dark1 || themeColor == ThemeColorValues.Text1)
        {
            color = scheme.Dark1Color;
        }
        else if (themeColor == ThemeColorValues.Light1 || themeColor == ThemeColorValues.Background1)
        {
            color = scheme.Light1Color;
        }
        else if (themeColor == ThemeColorValues.Dark2 || themeColor == ThemeColorValues.Text2)
        {
            color = scheme.Dark2Color;
        }
        else if (themeColor == ThemeColorValues.Light2 || themeColor == ThemeColorValues.Background2)
        {
            color = scheme.Light2Color;
        }
        else if (themeColor == ThemeColorValues.Accent1)
        {
            color = scheme.Accent1Color;
        }
        else if (themeColor == ThemeColorValues.Accent2)
        {
            color = scheme.Accent2Color;
        }
        else if (themeColor == ThemeColorValues.Accent3)
        {
            color = scheme.Accent3Color;
        }
        else if (themeColor == ThemeColorValues.Accent4)
        {
            color = scheme.Accent4Color;
        }
        else if (themeColor == ThemeColorValues.Accent5)
        {
            color = scheme.Accent5Color;
        }
        else if (themeColor == ThemeColorValues.Accent6)
        {
            color = scheme.Accent6Color;
        }
        else if (themeColor == ThemeColorValues.Hyperlink)
        {
            color = scheme.Hyperlink;
        }
        else if (themeColor == ThemeColorValues.FollowedHyperlink)
        {
            color = scheme.FollowedHyperlinkColor;
        }

        return color?.RgbColorModelHex?.Val?.Value ?? color?.SystemColor?.LastColor?.Value;
    }

    #endregion

    #region Styles

    private string GetParagraphStyleId(Paragraph paragraph)
    {
        return paragraph?.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? GetDefaultStyleId(StyleValues.Paragraph);
    }

    private static string GetTableStyleId(Paragraph paragraph)
    {
        Table table = paragraph?.Ancestors<Table>().FirstOrDefault();
        return table?.GetFirstChild<TableProperties>()?.TableStyle?.Val?.Value;
    }

    private string GetDefaultStyleId(StyleValues type)
    {
        return _styles?.Elements<Style>().FirstOrDefault(s => s.Type != null && s.Type.HasValue && s.Type.Value == type && s.Default != null && s.Default.HasValue && s.Default.Value)?.StyleId?.Value;
    }

    private Style FindStyle(string styleId)
    {
        if (_stylesById == null)
        {
            _stylesById = new Dictionary<string, Style>();
            if (_styles != null)
            {
                foreach (Style style in _styles.Elements<Style>())
                {
                    string id = style.StyleId?.Value;
                    if (id != null && !_stylesById.ContainsKey(id))
                    {
                        _stylesById[id] = style;
                    }
                }
            }
        }
        return styleId != null && _stylesById.TryGetValue(styleId, out Style found) ? found : null;
    }

    //The style and the styles it is based on, the most basic first.
    private List<Style> GetChain(string styleId)
    {
        List<Style> chain = new List<Style>();
        HashSet<string> seen = new HashSet<string>();
        string current = styleId;
        while (current != null && seen.Add(current))
        {
            Style style = FindStyle(current);
            if (style == null)
            {
                break;
            }
            chain.Add(style);
            current = style.BasedOn?.Val?.Value;
        }
        chain.Reverse();
        return chain;
    }

    #endregion

    #region Merging

    private static bool IsOn(OpenXmlElement element)
    {
        return !(element is OnOffType onOff) || onOff.Val == null || !onOff.Val.HasValue || onOff.Val.Value;
    }

    private static OpenXmlElement FindByName(OpenXmlElement properties, string localName)
    {
        return properties.ChildElements.FirstOrDefault(e => e.LocalName == localName);
    }

    private static void Remove(OpenXmlElement properties, string localName)
    {
        foreach (OpenXmlElement existing in properties.ChildElements.Where(e => e.LocalName == localName).ToList())
        {
            existing.Remove();
        }
    }

    /// <summary>Puts the properties of the source over those of the target: a property replaces the one of the same kind.</summary>
    private static void Overlay(OpenXmlCompositeElement target, OpenXmlElement source, bool skipToggles = false)
    {
        if (source == null)
        {
            return;
        }

        foreach (OpenXmlElement child in source.ChildElements)
        {
            string name = child.LocalName;
            if (child is OpenXmlUnknownElement || Array.IndexOf(IgnoredNames, name) >= 0 || (skipToggles && Array.IndexOf(ToggleNames, name) >= 0))
            {
                continue;
            }

            OpenXmlElement existing = FindByName(target, name);
            if (existing != null && (child is RunFonts || child is Indentation || child is SpacingBetweenLines))
            {
                MergeAttributes(existing, child);
            }
            else
            {
                if (existing != null)
                {
                    existing.Remove();
                }
                target.AppendChild(child.CloneNode(true));
            }
        }
    }

    //Properties such as spacing and indentation are made of attributes, and each one of them is inherited on its own.
    private static void MergeAttributes(OpenXmlElement target, OpenXmlElement source)
    {
        if (source is Indentation indentation && target is Indentation merged)
        {
            if (indentation.Left != null || indentation.Start != null)
            {
                merged.Left = null;
                merged.Start = null;
            }
            if (indentation.Right != null || indentation.End != null)
            {
                merged.Right = null;
                merged.End = null;
            }
            if (indentation.FirstLine != null || indentation.Hanging != null)
            {
                merged.FirstLine = null;
                merged.Hanging = null;
            }
        }
        else if (source is RunFonts fonts && target is RunFonts mergedFonts)
        {
            //A font that is named replaces a theme font, which would otherwise win.
            if (fonts.Ascii != null && fonts.AsciiTheme == null)
            {
                mergedFonts.AsciiTheme = null;
            }
            if (fonts.HighAnsi != null && fonts.HighAnsiTheme == null)
            {
                mergedFonts.HighAnsiTheme = null;
            }
            if (fonts.EastAsia != null && fonts.EastAsiaTheme == null)
            {
                mergedFonts.EastAsiaTheme = null;
            }
            if (fonts.ComplexScript != null && fonts.ComplexScriptTheme == null)
            {
                mergedFonts.ComplexScriptTheme = null;
            }
        }

        foreach (OpenXmlAttribute attribute in source.GetAttributes())
        {
            target.SetAttribute(attribute);
        }
    }

    #endregion
}
