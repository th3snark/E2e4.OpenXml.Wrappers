using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The styles of a document, as <c>Styles</c> in Word.
/// </summary>
public class StyleWrapperCollection : IEnumerable<StyleWrapper>
{
    private readonly MainDocumentPart _mainDocumentPart;

    public StyleWrapperCollection(MainDocumentPart mainDocumentPart)
    {
        ArgumentNullException.ThrowIfNull(mainDocumentPart);
        _mainDocumentPart = mainDocumentPart;
    }

    private IList<Style> GetElementList()
    {
        return _mainDocumentPart.StyleDefinitionsPart?.Styles?.Elements<Style>().ToList() ?? new List<Style>();
    }

    public StyleWrapper this[int index]
    {
        get
        {
            IList<Style> styles = GetElementList();
            if (index < 0 || index >= styles.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new StyleWrapper(styles[index]);
        }
    }

    /// <summary>
    /// The style with the name (ignoring case) or, failing that, the id; null if there is none.
    /// </summary>
    public StyleWrapper this[string name]
    {
        get
        {
            IList<Style> styles = GetElementList();
            Style style = styles.FirstOrDefault(s => String.Equals(s.StyleName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase))
                ?? styles.FirstOrDefault(s => s.StyleId?.Value == name);
            return style == null ? null : new StyleWrapper(style);
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    public bool Exists(string name)
    {
        return this[name] != null;
    }

    /// <summary>
    /// Adds a style, as <c>Styles.Add</c> in Word. A paragraph style is based on the default paragraph style.
    /// </summary>
    /// <exception cref="ArgumentException">The name is empty or a style with the name (or the id made from it) already exists.</exception>
    public StyleWrapper Add(string name, StyleType type)
    {
        if (String.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A style needs a name.", nameof(name));
        }
        if (!Enum.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }

        //The id is the name without the characters that are not allowed in one, as Word makes it.
        string styleId = new string(name.Where(Char.IsLetterOrDigit).ToArray());
        if (styleId.Length == 0 || Exists(name) || Exists(styleId))
        {
            throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "A style named '{0}' already exists or cannot be given an id.", name), nameof(name));
        }

        StyleDefinitionsPart part = _mainDocumentPart.StyleDefinitionsPart ?? _mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
        if (part.Styles == null)
        {
            part.Styles = new Styles();
        }

        Style style = new Style(new StyleName() { Val = name })
        {
            StyleId = styleId,
            CustomStyle = true,
            Type = type switch
            {
                StyleType.Character => StyleValues.Character,
                StyleType.Table => StyleValues.Table,
                StyleType.List => StyleValues.Numbering,
                _ => StyleValues.Paragraph
            }
        };
        if (type == StyleType.Paragraph)
        {
            string defaultId = part.Styles.Elements<Style>().FirstOrDefault(s => s.Type != null && s.Type.HasValue && s.Type.Value == StyleValues.Paragraph && s.Default != null && s.Default.HasValue && s.Default.Value)?.StyleId?.Value;
            if (defaultId != null)
            {
                style.InsertAfter(new BasedOn() { Val = defaultId }, style.StyleName);
            }
        }
        part.Styles.AppendChild(style);
        return new StyleWrapper(style);
    }

    /// <summary>
    /// Deletes a style, as <c>Style.Delete</c> in Word. Paragraphs, runs and tables that use it go back to the default style, and styles
    /// based on it lose their base.
    /// </summary>
    /// <exception cref="InvalidOperationException">The style is built in, or is the default style of its type.</exception>
    public void Delete(StyleWrapper style)
    {
        ArgumentNullException.ThrowIfNull(style);
        if (style.Style.Parent == null || style.Style.GetRoot() != _mainDocumentPart.StyleDefinitionsPart?.Styles)
        {
            throw new ArgumentException("The style is not a style of this document.", nameof(style));
        }
        if (style.BuiltIn || (style.Style.Default != null && style.Style.Default.HasValue && style.Style.Default.Value))
        {
            throw new InvalidOperationException("Built-in styles and default styles cannot be deleted.");
        }

        string styleId = style.StyleId;
        style.Style.Remove();

        foreach (Style other in _mainDocumentPart.StyleDefinitionsPart.Styles.Elements<Style>())
        {
            if (other.BasedOn?.Val?.Value == styleId)
            {
                other.BasedOn.Remove();
            }
            if (other.NextParagraphStyle?.Val?.Value == styleId)
            {
                other.NextParagraphStyle.Remove();
            }
            if (other.LinkedStyle?.Val?.Value == styleId)
            {
                other.LinkedStyle.Remove();
            }
        }

        if (_mainDocumentPart.Document != null)
        {
            foreach (ParagraphStyleId reference in _mainDocumentPart.Document.Descendants<ParagraphStyleId>().Where(e => e.Val?.Value == styleId).ToList())
            {
                reference.Remove();
            }
            foreach (RunStyle reference in _mainDocumentPart.Document.Descendants<RunStyle>().Where(e => e.Val?.Value == styleId).ToList())
            {
                reference.Remove();
            }
            foreach (TableStyle reference in _mainDocumentPart.Document.Descendants<TableStyle>().Where(e => e.Val?.Value == styleId).ToList())
            {
                reference.Remove();
            }
        }
    }

    #region IEnumerable Members

    public IEnumerator<StyleWrapper> GetEnumerator()
    {
        foreach (Style style in GetElementList())
        {
            yield return new StyleWrapper(style);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
