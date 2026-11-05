using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The formatting of a find or of its replacement: a font, a paragraph format, a style and highlighting. They are kept as the properties of a
/// run and a paragraph that are not in any document, written as the font and paragraph format of a range write them except that a value is
/// kept even when the text has it anyway from its styles, because it stands for what to look for or to apply.
/// </summary>
internal sealed class FindFormat
{
    private static readonly string[] OnOffRunProperties =
    {
        "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
        "vanish", "webHidden", "rtl", "cs", "specVanish", "oMath",
    };

    private static readonly string[] OnOffParagraphProperties =
    {
        "keepNext", "keepLines", "pageBreakBefore", "widowControl", "suppressLineNumbers", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "contextualSpacing", "mirrorIndents",
        "suppressOverlap",
    };

    //The elements whose attributes are several properties, so that applying one leaves the others as they are.
    private static readonly string[] Merged = { "rFonts", "lang", "spacing", "ind" };

    private static readonly string[] ParagraphPropertyOrder =
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
        "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    };

    private Paragraph _holder;
    private Run _run;
    private string _styleId;
    private StyleValues _styleType;

    public FindFormat()
    {
        Clear();
    }

    public FontWrapper Font { get; private set; }

    public ParagraphFormatWrapper ParagraphFormat { get; private set; }

    /// <summary>The name of the style, or null for any.</summary>
    public string Style { get; set; }

    /// <summary>True for highlighted text, false for text that is not, and null for either.</summary>
    public bool? Highlight { get; set; }

    /// <summary>The character properties that are set: the criteria for a run.</summary>
    public RunProperties RunProperties
    {
        get { return _run.RunProperties; }
    }

    /// <summary>The paragraph properties that are set.</summary>
    public IList<OpenXmlElement> ParagraphProperties
    {
        get
        {
            return _holder.ParagraphProperties == null
                ? new List<OpenXmlElement>()
                : _holder.ParagraphProperties.ChildElements.Where(e => e.LocalName != "rPr").ToList();
        }
    }

    public bool HasRunFormatting
    {
        get { return (RunProperties != null && RunProperties.HasChildren) || Highlight.HasValue; }
    }

    public bool HasParagraphFormatting
    {
        get { return ParagraphProperties.Count > 0; }
    }

    public bool HasFormatting
    {
        get { return HasRunFormatting || HasParagraphFormatting || Style != null; }
    }

    public void Clear()
    {
        _run = new Run(new Text("x"));
        _holder = new Paragraph(_run);
        RangeWrapper range = new RangeWrapper(new Body(_holder));
        Font = FontWrapper.ForCriteria(range);
        ParagraphFormat = ParagraphFormatWrapper.ForCriteria(range);
        Style = null;
        Highlight = null;
    }

    internal static bool IsOnOff(string localName)
    {
        return Array.IndexOf(OnOffRunProperties, localName) >= 0 || Array.IndexOf(OnOffParagraphProperties, localName) >= 0;
    }

    #region Matching

    //Whether the run has what the font and highlighting ask for, and its style.
    public bool MatchesRun(Run run, FormattingResolver resolver)
    {
        if (RunProperties != null && RunProperties.HasChildren)
        {
            RunProperties actual = resolver.ResolveRun(run);
            foreach (OpenXmlElement wanted in RunProperties.ChildElements)
            {
                if (!SameValue(wanted, Find(actual, wanted.LocalName)))
                {
                    return false;
                }
            }
        }

        if (Highlight.HasValue)
        {
            string value = resolver.ResolveRun(run).Highlight?.Val?.InnerText;
            bool highlighted = !string.IsNullOrEmpty(value) && value != "none";
            if (highlighted != Highlight.Value)
            {
                return false;
            }
        }

        if (Style != null && _styleType == StyleValues.Character)
        {
            return run.RunProperties?.RunStyle?.Val?.Value == _styleId;
        }
        return true;
    }

    //Whether the paragraph has what the paragraph format asks for, and its style.
    public bool MatchesParagraph(Paragraph paragraph, FormattingResolver resolver)
    {
        IList<OpenXmlElement> wanted = ParagraphProperties;
        if (wanted.Count > 0)
        {
            ParagraphProperties actual = resolver.ResolveParagraph(paragraph);
            foreach (OpenXmlElement element in wanted)
            {
                if (element.LocalName != "pStyle" && element.LocalName != "sectPr" && element.LocalName != "cnfStyle"
                    && !SameValue(element, Find(actual, element.LocalName)))
                {
                    return false;
                }
            }
        }

        if (Style != null && _styleType == StyleValues.Paragraph)
        {
            string used = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            return used == null ? StyleHelper.IsDefaultParagraphStyle(paragraph, _styleId) : used == _styleId;
        }
        return true;
    }

    /// <summary>Whether what is asked for is about paragraphs alone, so that a match is whole paragraphs.</summary>
    public bool IsParagraphOnly
    {
        get { return !HasRunFormatting && (HasParagraphFormatting || (Style != null && _styleType == StyleValues.Paragraph)); }
    }

    /// <summary>Looks up the style that is asked for in the document of the context, which a search does once before it matches anything.</summary>
    /// <exception cref="ArgumentException">The document has no such style.</exception>
    public void ResolveStyle(OpenXmlElement context)
    {
        _styleId = null;
        if (Style != null)
        {
            _styleId = StyleHelper.FindStyleId(context, Style, out _styleType);
        }
    }

    private static OpenXmlElement Find(OpenXmlElement properties, string localName)
    {
        return properties?.ChildElements.FirstOrDefault(e => e.LocalName == localName);
    }

    private static string AttributeName(string localName)
    {
        return localName == "left" ? "start" : localName == "right" ? "end" : localName;
    }

    private static bool SameValue(OpenXmlElement wanted, OpenXmlElement actual)
    {
        if (IsOnOff(wanted.LocalName))
        {
            return StyleWrapper.IsOn(wanted) == (actual != null && StyleWrapper.IsOn(actual));
        }
        if (actual == null)
        {
            return false;
        }

        List<OpenXmlAttribute> have = actual.GetAttributes().ToList();
        foreach (OpenXmlAttribute attribute in wanted.GetAttributes())
        {
            string name = AttributeName(attribute.LocalName);
            OpenXmlAttribute found = have.FirstOrDefault(a => AttributeName(a.LocalName) == name);
            if (found.LocalName == null || !string.Equals(found.Value, attribute.Value, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }
        return true;
    }

    #endregion

    #region Applying

    /// <summary>Gives the text of the range what the font, highlighting, paragraph format and style say.</summary>
    public void ApplyTo(RangeWrapper range)
    {
        if (Style != null)
        {
            range.Style = Style;
        }
        if (HasRunFormatting)
        {
            RunProperties source = RunProperties == null ? new RunProperties() : (RunProperties)RunProperties.CloneNode(true);
            string[] remove = Array.Empty<string>();
            if (Highlight == true)
            {
                source.SetRunProperty(new DocumentFormat.OpenXml.Wordprocessing.Highlight() { Val = HighlightColorValues.Yellow });
            }
            else if (Highlight == false)
            {
                remove = new[] { "highlight" };
            }
            new FontWrapper(range).ApplyFrom(source, remove);
        }
        if (HasParagraphFormatting)
        {
            new ParagraphFormatWrapper(range).ApplyFrom(ParagraphProperties);
        }
    }

    internal static void MergeInto(OpenXmlCompositeElement properties, OpenXmlElement source, Action<OpenXmlElement> place)
    {
        OpenXmlElement existing = properties.ChildElements.FirstOrDefault(e => e.LocalName == source.LocalName);
        if (existing != null && Array.IndexOf(Merged, source.LocalName) >= 0)
        {
            foreach (OpenXmlAttribute attribute in source.GetAttributes())
            {
                existing.SetAttribute(attribute);
            }
            return;
        }
        place(source.CloneNode(true));
    }

    internal static void SetParagraphProperty(ParagraphProperties properties, OpenXmlElement property)
    {
        foreach (OpenXmlElement existing in properties.ChildElements.Where(e => e.LocalName == property.LocalName).ToList())
        {
            existing.Remove();
        }

        int order = ParagraphOrder(property);
        OpenXmlElement predecessor = properties.ChildElements.LastOrDefault(e => ParagraphOrder(e) < order);
        if (predecessor == null)
        {
            properties.PrependChild(property);
        }
        else
        {
            predecessor.InsertAfterSelf(property);
        }
    }

    private static int ParagraphOrder(OpenXmlElement element)
    {
        int index = Array.IndexOf(ParagraphPropertyOrder, element.LocalName);
        return index >= 0 ? index : ParagraphPropertyOrder.Length;
    }

    #endregion
}
