using System;
using System.Linq;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A style of the document, as <c>Style</c> in Word.
/// </summary>
/// <remarks>
/// <see cref="Name"/> is the name Word shows for a built-in style ("Heading 1"), whatever the document stores ("heading 1"). The formatting of a
/// style is read and set through <see cref="Font"/> and <see cref="ParagraphFormat"/>.
/// </remarks>
public class StyleWrapper : System.IEquatable<StyleWrapper>
{
    private readonly Style _style;

    public StyleWrapper(Style style)
    {
        Guard.ThrowIfNull(style);
        _style = style;
    }

    public Style Style
    {
        get { return _style; }
    }

    /// <summary>
    /// The name of the style, as <c>Style.NameLocal</c> in Word. Falls back to the style id for a style without a name.
    /// </summary>
    public string Name
    {
        get { return StyleHelper.GetName(_style); }
    }

    /// <summary>
    /// The name of the style, as <c>Style.NameLocal</c> in Word. The same as <see cref="Name"/>: the library does not localise style names.
    /// </summary>
    public string NameLocal
    {
        get { return Name; }
    }

    /// <summary>
    /// Deletes the style, as <c>Style.Delete</c> in Word. Paragraphs, runs and tables that use it go back to the default style, and styles
    /// based on it lose their base.
    /// </summary>
    /// <exception cref="InvalidOperationException">The style is built in, or is the default style of its type, or is not in a document.</exception>
    public void Delete()
    {
        Styles styles = _style.Parent as Styles;
        MainDocumentPart mainDocumentPart = StyleHelper.GetMainDocumentPart(_style);
        if (styles == null || mainDocumentPart == null)
        {
            throw new InvalidOperationException("The style is not a style of a document.");
        }
        if (BuiltIn || (_style.Default != null && _style.Default.HasValue && _style.Default.Value))
        {
            throw new InvalidOperationException("Built-in styles and default styles cannot be deleted.");
        }

        string styleId = StyleId;
        _style.Remove();

        foreach (Style other in styles.Elements<Style>())
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

        if (mainDocumentPart.Document != null)
        {
            foreach (ParagraphStyleId reference in mainDocumentPart.Document.Descendants<ParagraphStyleId>().Where(e => e.Val?.Value == styleId).ToList())
            {
                reference.Remove();
            }
            foreach (RunStyle reference in mainDocumentPart.Document.Descendants<RunStyle>().Where(e => e.Val?.Value == styleId).ToList())
            {
                reference.Remove();
            }
            foreach (TableStyle reference in mainDocumentPart.Document.Descendants<TableStyle>().Where(e => e.Val?.Value == styleId).ToList())
            {
                reference.Remove();
            }
        }
    }

    /// <summary>
    /// The id by which paragraphs, runs and tables refer to the style.
    /// </summary>
    public string StyleId
    {
        get { return _style.StyleId?.Value; }
    }

    public StyleType Type
    {
        get
        {
            StyleValues? type = _style.Type != null && _style.Type.HasValue ? _style.Type.Value : null;
            if (type == StyleValues.Character)
            {
                return StyleType.Character;
            }
            if (type == StyleValues.Table)
            {
                return StyleType.Table;
            }
            return type == StyleValues.Numbering ? StyleType.List : StyleType.Paragraph;
        }
    }

    /// <summary>
    /// True for the styles that come with Word, false for those added to the document (<c>Style.BuiltIn</c> in Word).
    /// </summary>
    public bool BuiltIn
    {
        get { return _style.CustomStyle == null || !_style.CustomStyle.HasValue || !_style.CustomStyle.Value; }
    }

    /// <summary>
    /// The name of the style this style is based on, or null.
    /// </summary>
    public string BaseStyle
    {
        get { return GetName(_style.BasedOn?.Val?.Value); }
    }

    /// <summary>
    /// The name of the style applied to the paragraph after one in this style, or null when it is the same style.
    /// </summary>
    public string NextParagraphStyle
    {
        get { return GetName(_style.NextParagraphStyle?.Val?.Value); }
    }

    /// <summary>
    /// The style this style is based on as an object, or null.
    /// </summary>
    public StyleWrapper BaseStyleObject
    {
        get { return StyleHelper.GetStyleObject(_style, BaseStyle); }
    }

    /// <summary>
    /// The style applied to the paragraph after one in this style as an object, or null when it is the same style.
    /// </summary>
    public StyleWrapper NextParagraphStyleObject
    {
        get { return StyleHelper.GetStyleObject(_style, NextParagraphStyle); }
    }

    /// <summary>
    /// The character formatting set in the style itself, as <c>Style.Font</c> in Word. Formatting inherited from the style it is based on is
    /// not included. See <see cref="FontWrapper"/> for the rules for values.
    /// </summary>
    public FontWrapper Font
    {
        get { return new FontWrapper(_style); }
    }

    /// <summary>
    /// The paragraph formatting set in the style itself, as <c>Style.ParagraphFormat</c> in Word. Formatting inherited from the style it is
    /// based on is not included. See <see cref="ParagraphFormatWrapper"/> for the rules for values.
    /// </summary>
    public ParagraphFormatWrapper ParagraphFormat
    {
        get { return new ParagraphFormatWrapper(_style); }
    }

    private string GetName(string styleId)
    {
        return styleId == null ? null : StyleHelper.GetStyleName(_style, styleId);
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as StyleWrapper);
    }

    public bool Equals(StyleWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_style, other._style);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_style);
    }

    public static bool operator ==(StyleWrapper left, StyleWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(StyleWrapper left, StyleWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
