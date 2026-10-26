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
        Guard.ThrowIfNull(mainDocumentPart);
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
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new StyleWrapper(styles[index]);
        }
    }

    /// <summary>
    /// The style with the name (ignoring case) or, failing that, the id; null if there is none. As in Word, a built-in style such as "Heading 1" is
    /// found whether or not the document has used it yet: it is added to the document when it is asked for.
    /// </summary>
    public StyleWrapper this[string name]
    {
        get
        {
            Style style = BuiltInStyles.GetOrCreate(_mainDocumentPart, name);
            return style == null ? null : new StyleWrapper(style);
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    /// <summary>
    /// Whether the document holds a style with the name or id. Unlike the indexer this does not add a built-in style that the document has not used.
    /// </summary>
    public bool Exists(string name)
    {
        return StyleHelper.FindStyle(_mainDocumentPart.StyleDefinitionsPart?.Styles, name) != null;
    }

    /// <summary>
    /// A built-in style, as <c>Styles(wdStyleHeading1)</c> in Word. It is added to the document if the document does not have it yet.
    /// </summary>
    public StyleWrapper this[BuiltInStyle style]
    {
        get { return this[BuiltInStyles.GetName(style)]; }
    }

    /// <summary>
    /// Adds a style, as <c>Styles.Add</c> in Word. A paragraph style is based on the default paragraph style.
    /// </summary>
    /// <exception cref="ArgumentException">The name is empty or a style with the name (or the id made from it) already exists.</exception>
    public StyleWrapper Add(string name, StyleType type)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A style needs a name.", nameof(name));
        }
        if (!EnumCompat.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }

        //The id is the name without the characters that are not allowed in one, as Word makes it.
        string styleId = new string(name.Where(char.IsLetterOrDigit).ToArray());
        if (styleId.Length == 0 || Exists(name) || Exists(styleId) || BuiltInStyles.IsBuiltInName(name) || BuiltInStyles.IsBuiltInName(styleId))
        {
            throw new ArgumentException(string.Format(CultureInfo.InvariantCulture, "A style named '{0}' already exists or cannot be given an id.", name), nameof(name));
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
