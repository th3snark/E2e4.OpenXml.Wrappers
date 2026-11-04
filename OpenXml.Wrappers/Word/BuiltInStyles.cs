using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The styles that come with Word. A document only holds the ones that have been used, but in Word they can all be asked for by name, so
/// the definitions are kept here and added to a document when one is asked for or applied.
/// </summary>
/// <remarks>
/// The definitions are the ones Word 365 writes when each style is applied in a new document (<c>Templates/BuiltInStyles.xml</c>, written by
/// <c>tools/WordReference</c>), so they are the definitions of a style in use: Word drops <c>w:semiHidden</c> from a style once it is used.
/// Adding a paragraph style adds its linked character style ("Heading 1 Char") too, as Word does.
/// </remarks>
internal static class BuiltInStyles
{
    private sealed class Definition
    {
        public BuiltInStyle? Id;
        public string Name;
        public string StoredName;
        public string StyleId;
        public Style Style;
    }

    //The style ids of the BuiltInStyle values whose id is not the name of the value.
    private static readonly Dictionary<BuiltInStyle, string> StyleIds = new Dictionary<BuiltInStyle, string>()
    {
        { BuiltInStyle.Toc1, "TOC1" }, { BuiltInStyle.Toc2, "TOC2" }, { BuiltInStyle.Toc3, "TOC3" }, { BuiltInStyle.Toc4, "TOC4" },
        { BuiltInStyle.Toc5, "TOC5" }, { BuiltInStyle.Toc6, "TOC6" }, { BuiltInStyle.Toc7, "TOC7" }, { BuiltInStyle.Toc8, "TOC8" },
        { BuiltInStyle.Toc9, "TOC9" }, { BuiltInStyle.HyperlinkFollowed, "FollowedHyperlink" }, { BuiltInStyle.NormalTable, "TableNormal" },
    };

    //The names Word shows for built-in styles stored under another name, other than the lower-case ones ("heading 1" shows as "Heading 1").
    private static readonly Dictionary<string, string> DisplayNames = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        { "annotation text", "Comment Text" }, { "annotation reference", "Comment Reference" }, { "annotation subject", "Comment Subject" }, { "FollowedHyperlink", "Followed Hyperlink" },
    };

    private static readonly List<Definition> Definitions = CreateDefinitions();
    private static readonly Dictionary<BuiltInStyle, Definition> ById = Definitions.Where(d => d.Id.HasValue).ToDictionary(d => d.Id.Value);

    /// <summary>The name Word shows for the built-in style that is stored under the name, or null if it is not one that is known.</summary>
    public static string GetDisplayName(string storedName)
    {
        return storedName == null ? null : Find(storedName)?.Name;
    }

    public static bool IsBuiltInName(string name)
    {
        return name != null && Find(name) != null;
    }

    /// <summary>The name or id of a built-in style that has a definition.</summary>
    public static string GetName(BuiltInStyle style)
    {
        if (!ById.TryGetValue(style, out Definition definition))
        {
            throw new ArgumentOutOfRangeException(nameof(style));
        }
        return definition.Name;
    }

    /// <summary>
    /// Finds the style in the document and, if the document does not have it but it is a built-in one, adds it. Returns null when there is no
    /// such style.
    /// </summary>
    public static Style GetOrCreate(MainDocumentPart mainDocumentPart, string nameOrId)
    {
        Style existing = StyleHelper.FindStyle(mainDocumentPart?.StyleDefinitionsPart?.Styles, nameOrId);
        if (existing != null || mainDocumentPart == null)
        {
            return existing;
        }

        Definition definition = Find(nameOrId);
        return definition == null ? null : Create(mainDocumentPart, definition);
    }

    /// <summary>
    /// Records that the style is used, as Word does when it is applied: a style with <c>w:unhideWhenUsed</c> loses <c>w:semiHidden</c>, and so
    /// does its linked style.
    /// </summary>
    public static void MarkUsed(MainDocumentPart mainDocumentPart, Style style)
    {
        if (style.UnhideWhenUsed == null || style.SemiHidden == null)
        {
            return;
        }
        style.SemiHidden.Remove();
        string linkedId = style.LinkedStyle?.Val?.Value;
        Style linked = linkedId == null ? null : mainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == linkedId);
        linked?.SemiHidden?.Remove();
    }

    private static Definition Find(string nameOrId)
    {
        return Definitions.FirstOrDefault(d => string.Equals(d.Name, nameOrId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(d.StoredName, nameOrId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(d.StyleId, nameOrId, StringComparison.OrdinalIgnoreCase));
    }

    private static Definition FindById(string styleId)
    {
        return Definitions.FirstOrDefault(d => d.StyleId == styleId);
    }

    private static Style Create(MainDocumentPart mainDocumentPart, Definition definition)
    {
        StyleDefinitionsPart part = mainDocumentPart.StyleDefinitionsPart ?? mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        if (part.Styles == null)
        {
            part.Styles = new Styles();
        }

        Style existing = part.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == definition.StyleId);
        if (existing != null)
        {
            return existing;
        }

        //The styles that this one is based on and followed by come first, so that the document never refers to a style it does not have.
        foreach (string dependency in new[] { definition.Style.BasedOn?.Val?.Value, definition.Style.NextParagraphStyle?.Val?.Value })
        {
            Definition other = dependency == null || dependency == definition.StyleId ? null : FindById(dependency);
            if (other != null)
            {
                Create(mainDocumentPart, other);
            }
        }

        Style style = (Style)definition.Style.CloneNode(true);
        if (definition.StyleId == "Header" || definition.StyleId == "Footer")
        {
            SetTabStopsToTextWidth(style, mainDocumentPart);
        }

        //Default styles go before the others, as Word writes them; the rest go at the end, in the order they are added.
        if (style.Default != null && style.Default.Value)
        {
            OpenXmlElement after = part.Styles.Elements<Style>().LastOrDefault(s => s.Default != null && s.Default.Value)
                ?? (OpenXmlElement)part.Styles.Elements<LatentStyles>().LastOrDefault()
                ?? part.Styles.Elements<DocDefaults>().LastOrDefault();
            if (after == null)
            {
                part.Styles.PrependChild(style);
            }
            else
            {
                after.InsertAfterSelf(style);
            }
        }
        else
        {
            part.Styles.AppendChild(style);
        }

        //The linked character style of a paragraph style (or the other way round) comes straight after it.
        Definition linked = style.LinkedStyle?.Val?.Value == null ? null : FindById(style.LinkedStyle.Val.Value);
        if (linked != null)
        {
            Create(mainDocumentPart, linked);
        }
        return style;
    }

    //Word's Header and Footer styles have a centre tab stop in the middle of the text and a right one at its right margin, for the page of a
    //new document; the definitions are for an A4 page, so they are moved to the text width of the document's last section.
    private static void SetTabStopsToTextWidth(Style style, MainDocumentPart mainDocumentPart)
    {
        SectionProperties section = mainDocumentPart.Document?.Body?.Elements<SectionProperties>().LastOrDefault();
        uint? pageWidth = section?.GetFirstChild<PageSize>()?.Width?.Value;
        PageMargin margin = section?.GetFirstChild<PageMargin>();
        if (!pageWidth.HasValue || margin?.Left == null || margin.Right == null)
        {
            return;
        }
        int textWidth = (int)pageWidth.Value - (int)margin.Left.Value - (int)margin.Right.Value - (int)(margin.Gutter?.Value ?? 0);
        List<TabStop> tabs = style.StyleParagraphProperties?.Tabs?.Elements<TabStop>().ToList();
        if (textWidth <= 0 || tabs == null || tabs.Count != 2)
        {
            return;
        }
        tabs[0].Position = textWidth / 2;
        tabs[1].Position = textWidth;
    }

    private static List<Definition> CreateDefinitions()
    {
        Styles styles;
        using (Stream stream = NewDocument.OpenTemplate("BuiltInStyles.xml"))
        using (StreamReader reader = new StreamReader(stream))
        {
            string xml = reader.ReadToEnd();
            styles = new Styles(xml.StartsWith("<?xml", StringComparison.Ordinal) ? xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2) : xml);
        }

        Dictionary<string, BuiltInStyle> ids = new Dictionary<string, BuiltInStyle>(StringComparer.Ordinal);
        //net48 has no generic Enum.GetValues.
#pragma warning disable CA2263
        foreach (BuiltInStyle id in Enum.GetValues(typeof(BuiltInStyle)))
#pragma warning restore CA2263
        {
            ids.Add(StyleIds.TryGetValue(id, out string styleId) ? styleId : id.ToString(), id);
        }

        List<Definition> list = new List<Definition>();
        foreach (Style style in styles.Elements<Style>())
        {
            string stored = style.StyleName.Val.Value;
            list.Add(new Definition()
            {
                Id = ids.TryGetValue(style.StyleId.Value, out BuiltInStyle id) ? id : (BuiltInStyle?)null,
                Name = ToDisplayName(stored),
                StoredName = stored,
                StyleId = style.StyleId.Value,
                Style = style,
            });
        }
        return list;
    }

    //"heading 1" is shown as "Heading 1" and "toc 1" as "TOC 1"; a few others have a different name altogether.
    private static string ToDisplayName(string storedName)
    {
        if (DisplayNames.TryGetValue(storedName, out string name))
        {
            return name;
        }
        if (storedName.Any(char.IsUpper))
        {
            return storedName;
        }
        return string.Join(" ", storedName.Split(' ').Select(w => w == "toc" ? "TOC" : w.Length == 0 ? w : char.ToUpper(w[0], CultureInfo.InvariantCulture) + w.Substring(1)));
    }
}
