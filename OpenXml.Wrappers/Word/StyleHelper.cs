using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Looks up the styles of the document that an element is in.
/// </summary>
internal static class StyleHelper
{
    /// <summary>
    /// Returns the styles of the document that holds the element, or null if it is not in a document or the document has no styles.
    /// </summary>
    internal static Styles GetStyles(OpenXmlElement context)
    {
        return GetMainDocumentPart(context)?.StyleDefinitionsPart?.Styles;
    }

    /// <summary>
    /// Returns the main document part of the document that holds the element, or null if it is not in a document.
    /// </summary>
    internal static MainDocumentPart GetMainDocumentPart(OpenXmlElement context)
    {
        OpenXmlPart part = (context.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;
        return part as MainDocumentPart ?? part?.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
    }

    /// <summary>
    /// Finds a style of the document by its name (ignoring case) or, failing that, by its id. A built-in style is also found by the name that
    /// Word shows for it, whatever name it is stored under in the document. Returns null if there is none.
    /// </summary>
    internal static Style FindStyle(Styles styles, string nameOrId)
    {
        if (styles == null || nameOrId == null)
        {
            return null;
        }

        return styles.Elements<Style>().FirstOrDefault(s => String.Equals(s.StyleName?.Val?.Value, nameOrId, StringComparison.OrdinalIgnoreCase))
            ?? styles.Elements<Style>().FirstOrDefault(s => !IsCustom(s) && String.Equals(BuiltInStyles.GetDisplayName(s.StyleName?.Val?.Value), nameOrId, StringComparison.OrdinalIgnoreCase))
            ?? styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == nameOrId);
    }

    internal static bool IsCustom(Style style)
    {
        return style.CustomStyle != null && style.CustomStyle.HasValue && style.CustomStyle.Value;
    }

    /// <summary>
    /// The name that Word gives the style: the name it is stored under, except that built-in styles have the name Word shows for them
    /// ("Heading 1" rather than "heading 1").
    /// </summary>
    internal static string GetName(Style style)
    {
        string stored = style.StyleName?.Val?.Value;
        if (stored == null)
        {
            return style.StyleId?.Value;
        }
        return IsCustom(style) ? stored : BuiltInStyles.GetDisplayName(stored) ?? stored;
    }

    /// <summary>
    /// Finds a style by its name (ignoring case) or, failing that, by its id. A built-in style that the document does not have yet is added.
    /// </summary>
    /// <exception cref="InvalidOperationException">The element is not in a document.</exception>
    /// <exception cref="ArgumentException">There is no style with that name.</exception>
    public static string FindStyleId(OpenXmlElement context, string nameOrId, out StyleValues type)
    {
        MainDocumentPart mainDocumentPart = GetMainDocumentPart(context);
        if (mainDocumentPart == null)
        {
            throw new InvalidOperationException("The document does not have any styles.");
        }

        Style style = BuiltInStyles.GetOrCreate(mainDocumentPart, nameOrId);
        if (style == null)
        {
            throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The document does not have a style named '{0}'.", nameOrId), nameof(nameOrId));
        }

        type = style.Type != null && style.Type.HasValue ? style.Type.Value : StyleValues.Paragraph;
        return style.StyleId.Value;
    }

    /// <summary>
    /// Returns the name of the style with the id. If the style cannot be found the id is returned.
    /// </summary>
    public static string GetStyleName(OpenXmlElement context, string styleId)
    {
        Style style = GetStyles(context)?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
        return style == null ? styleId : GetName(style);
    }

    /// <summary>
    /// Returns the name of the default style of the type (such as "Normal" for paragraphs), or null if it cannot be determined.
    /// </summary>
    public static string GetDefaultStyleName(OpenXmlElement context, StyleValues type)
    {
        Style style = GetStyles(context)?.Elements<Style>().FirstOrDefault(s => s.Type != null && s.Type.HasValue && s.Type.Value == type && s.Default != null && s.Default.HasValue && s.Default.Value);
        return style == null ? null : GetName(style);
    }
}
