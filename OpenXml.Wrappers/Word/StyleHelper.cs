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
    private static Styles GetStyles(OpenXmlElement context)
    {
        OpenXmlPart part = (context.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;
        MainDocumentPart mainDocumentPart = part as MainDocumentPart ?? part?.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
        return mainDocumentPart?.StyleDefinitionsPart?.Styles;
    }

    /// <summary>
    /// Finds a style by its name (ignoring case) or, failing that, by its id.
    /// </summary>
    /// <exception cref="InvalidOperationException">The document has no styles.</exception>
    /// <exception cref="ArgumentException">There is no style with that name.</exception>
    public static string FindStyleId(OpenXmlElement context, string nameOrId, out StyleValues type)
    {
        Styles styles = GetStyles(context);
        if (styles == null)
        {
            throw new InvalidOperationException("The document does not have any styles.");
        }

        Style style = styles.Elements<Style>().FirstOrDefault(s => String.Equals(s.StyleName?.Val?.Value, nameOrId, StringComparison.OrdinalIgnoreCase))
            ?? styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == nameOrId);
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
        return style?.StyleName?.Val?.Value ?? styleId;
    }

    /// <summary>
    /// Returns the name of the default style of the type (such as "Normal" for paragraphs), or null if it cannot be determined.
    /// </summary>
    public static string GetDefaultStyleName(OpenXmlElement context, StyleValues type)
    {
        Style style = GetStyles(context)?.Elements<Style>().FirstOrDefault(s => s.Type != null && s.Type.HasValue && s.Type.Value == type && s.Default != null && s.Default.HasValue && s.Default.Value);
        return style?.StyleName?.Val?.Value ?? style?.StyleId?.Value;
    }
}
