using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The look a table of contents is given by the Table of Contents dialog, as <c>WdTocFormat</c> in Word (with the same values).
/// </summary>
public enum TocFormat
{
    /// <summary>The look of the document's template: setting it changes nothing.</summary>
    Template = 0,

    /// <summary>Classic.</summary>
    Classic = 1,

    /// <summary>Distinctive.</summary>
    Distinctive = 2,

    /// <summary>Fancy.</summary>
    Fancy = 3,

    /// <summary>Modern.</summary>
    Modern = 4,

    /// <summary>Formal.</summary>
    Formal = 5,

    /// <summary>Simple.</summary>
    Simple = 6
}

public partial class TableOfContentsWrapperCollection
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static Dictionary<TocFormat, List<Style>> ReadFormats()
    {
        XDocument document;
        using (Stream stream = NewDocument.OpenTemplate("TocFormats.xml"))
        {
            document = XDocument.Load(stream);
        }
        Dictionary<TocFormat, List<Style>> formats = new Dictionary<TocFormat, List<Style>>();
        foreach (XElement format in document.Root.Elements("format"))
        {
            formats[(TocFormat)int.Parse((string)format.Attribute("value"), CultureInfo.InvariantCulture)] =
                format.Elements(W + "style").Select(s => new Style(s.ToString(System.Xml.Linq.SaveOptions.DisableFormatting))).ToList();
        }
        return formats;
    }

    /// <summary>
    /// The look of the tables of contents, as <c>TablesOfContents.Format</c> in Word. A look is the definitions of the nine styles
    /// <c>TOC 1</c> to <c>TOC 9</c>, which belong to every table of the document, so setting it changes all of them; <see cref="TocFormat.Template"/>
    /// changes nothing, as in Word. Reading gives the look the styles have, or <see cref="TocFormat.Template"/> if they have none of the looks.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a value of <see cref="TocFormat"/>.</exception>
    /// <exception cref="NotSupportedException">The tables are not in a document.</exception>
    public TocFormat Format
    {
        get
        {
            Styles styles = GetStyles(false);
            if (styles == null)
            {
                return TocFormat.Template;
            }
            foreach (KeyValuePair<TocFormat, List<Style>> format in ReadFormats())
            {
                bool all = format.Value.All(look => styles.Elements<Style>().Any(s => s.StyleId?.Value == look.StyleId?.Value && SameStyle(s, look)));
                if (all)
                {
                    return format.Key;
                }
            }
            return TocFormat.Template;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            if (value == TocFormat.Template)
            {
                return;
            }
            Styles styles = GetStyles(true);
            foreach (Style look in ReadFormats()[value])
            {
                Style existing = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == look.StyleId?.Value);
                if (existing != null)
                {
                    existing.InsertBeforeSelf(look.CloneNode(true));
                    existing.Remove();
                }
                else
                {
                    styles.AppendChild(look.CloneNode(true));
                }
            }
        }
    }

    //The same definition, whatever namespace declarations the two are written with.
    private static bool SameStyle(Style one, Style other)
    {
        return XNode.DeepEquals(Plain(one), Plain(other));
    }

    private static XElement Plain(Style style)
    {
        XElement element = XElement.Parse(style.OuterXml);
        foreach (XElement each in element.DescendantsAndSelf())
        {
            each.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        }
        return element;
    }

    private Styles GetStyles(bool create)
    {
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(_range.AnchorElement);
        if (main == null)
        {
            throw new NotSupportedException("The look of a table of contents can be set only in a document.");
        }
        if (create)
        {
            BuiltInStyles.GetOrCreate(main, "TOC1");
        }
        return main.StyleDefinitionsPart?.Styles;
    }
}
