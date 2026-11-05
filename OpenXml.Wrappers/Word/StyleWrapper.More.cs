using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class StyleWrapper
{
    /// <summary>Whether Word redefines the style from the formatting of a paragraph in it when that is changed (<c>Style.AutomaticallyUpdate</c>).</summary>
    public bool AutomaticallyUpdate
    {
        get { return IsOn(_style.AutoRedefine); }
        set { _style.AutoRedefine = value ? new AutoRedefine() : null; }
    }

    /// <summary>Whether the style is locked, so that its formatting cannot be changed when the document is protected (<c>Style.Locked</c>).</summary>
    public bool Locked
    {
        get { return IsOn(_style.Locked); }
        set { _style.Locked = value ? new Locked() : null; }
    }

    /// <summary>Whether the style is in the Styles gallery as a recommended style (<c>Style.QuickStyle</c>).</summary>
    public bool QuickStyle
    {
        get { return IsOn(_style.PrimaryStyle); }
        set { _style.PrimaryStyle = value ? new PrimaryStyle() : null; }
    }

    /// <summary>Whether the style, when it is hidden, shows again once it is used (<c>Style.UnhideWhenUsed</c>).</summary>
    public bool UnhideWhenUsed
    {
        get { return IsOn(_style.UnhideWhenUsed); }
        set { _style.UnhideWhenUsed = value ? new UnhideWhenUsed() : null; }
    }

    /// <summary>
    /// Whether the style is hidden from the Styles task pane until it is used (<c>Style.Visibility</c> in Word, which as measured in Word 16
    /// is true for a style that is hidden, such as <c>Heading 5</c>, and false for one that shows, such as <c>Normal</c>).
    /// </summary>
    public bool Visibility
    {
        get { return IsOn(_style.SemiHidden); }
        set { _style.SemiHidden = value ? new SemiHidden() : null; }
    }

    /// <summary>
    /// The place of the style in the Styles gallery and task pane, lowest first (<c>Style.Priority</c>). Word stores one less (<c>w:uiPriority</c>),
    /// and 1 when there is none.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int Priority
    {
        get { return (_style.UIPriority?.Val?.Value ?? 0) + 1; }
        set
        {
            Guard.ThrowIfLessThan(value, 1);
            SetPriority(_style, value);
            Style linked = LinkedStyleElement();
            if (linked != null)
            {
                SetPriority(linked, value);
            }
        }
    }

    private static void SetPriority(Style style, int value)
    {
        style.UIPriority = value == 1 ? null : new UIPriority() { Val = value - 1 };
    }

    /// <summary>Whether the style is in use. A style that is in the document is.</summary>
    public bool InUse
    {
        get { return true; }
    }

    /// <summary>Whether the style is linked to another, a paragraph style to a character style that carries its character formatting (<c>Style.Linked</c>).</summary>
    public bool Linked
    {
        get { return _style.LinkedStyle?.Val?.Value != null; }
    }

    /// <summary>The style this one is linked to, or null (<c>Style.LinkStyle</c>).</summary>
    public StyleWrapper LinkStyle
    {
        get
        {
            Style linked = LinkedStyleElement();
            return linked == null ? null : new StyleWrapper(linked);
        }
    }

    private Style LinkedStyleElement()
    {
        return LinkedOf(_style);
    }

    internal static Style LinkedOf(Style style)
    {
        string id = style.LinkedStyle?.Val?.Value;
        return id == null ? null : StyleHelper.GetStyles(style)?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == id);
    }

    /// <summary>
    /// The borders of a paragraph style (<c>Style.Borders</c>): the four edges, as for a paragraph. Null for a style of another type, as in Word.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return Type == StyleType.Paragraph ? new BordersWrapper(_style, false) : null; }
    }

    /// <summary>
    /// The shading of a paragraph style (<c>Style.Shading</c>). As Word does, it is given to the character style linked to it as well. Null for
    /// a style of another type.
    /// </summary>
    public ShadingWrapper Shading
    {
        get { return Type == StyleType.Paragraph ? ParagraphShading(_style) : null; }
    }

    private static Shading NewShading()
    {
        return new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
    }

    //The shading of the paragraphs of a style, and of the characters of the style linked to it.
    internal static ShadingWrapper ParagraphShading(Style style)
    {
        return new ShadingWrapper(create =>
        {
            List<Shading> shadings = new List<Shading>();
            StyleParagraphProperties paragraph = style.StyleParagraphProperties;
            if (paragraph == null && create)
            {
                paragraph = new StyleParagraphProperties();
                style.StyleParagraphProperties = paragraph;
            }
            Shading own = paragraph?.Shading;
            if (own == null && create)
            {
                own = NewShading();
                paragraph.Shading = own;
            }
            shadings.Add(own);

            Style linked = LinkedOf(style);
            if (linked != null)
            {
                shadings.Add(CharacterShadingElement(linked, create));
            }
            return shadings;
        });
    }

    //The shading behind the characters of a style.
    internal static ShadingWrapper CharacterShading(Style style)
    {
        return new ShadingWrapper(create => new List<Shading> { CharacterShadingElement(style, create) });
    }

    private static Shading CharacterShadingElement(Style style, bool create)
    {
        StyleRunProperties properties = style.StyleRunProperties;
        if (properties == null && create)
        {
            properties = new StyleRunProperties();
            style.StyleRunProperties = properties;
        }
        Shading shading = properties?.Shading;
        if (shading == null && create)
        {
            shading = NewShading();
            properties.Shading = shading;
        }
        return shading;
    }

    /// <summary>
    /// Whether the space between paragraphs of this style is left out (<c>Style.NoSpaceBetweenParagraphsOfSameStyle</c>), which is
    /// <c>w:contextualSpacing</c>.
    /// </summary>
    public bool NoSpaceBetweenParagraphsOfSameStyle
    {
        get { return IsOn(_style.StyleParagraphProperties?.ContextualSpacing); }
        set
        {
            if (value)
            {
                if (_style.StyleParagraphProperties == null)
                {
                    _style.StyleParagraphProperties = new StyleParagraphProperties();
                }
                _style.StyleParagraphProperties.ContextualSpacing = new ContextualSpacing();
            }
            else if (_style.StyleParagraphProperties != null)
            {
                _style.StyleParagraphProperties.ContextualSpacing = null;
                if (!_style.StyleParagraphProperties.HasChildren)
                {
                    _style.StyleParagraphProperties = null;
                }
            }
        }
    }

    /// <summary>
    /// Whether text in the style is left out of the proofing tools (<c>Style.NoProofing</c>). Like the other character formatting set
    /// here it is also given to the linked style.
    /// </summary>
    public bool NoProofing
    {
        get { return IsOn(_style.StyleRunProperties?.NoProof); }
        set { SetRunProperty(s => s.NoProof = value ? new NoProof() : null); }
    }

    /// <summary>The language of the text in the style, as the number Word uses (<c>WdLanguageID</c>, for example 1033 for English (United States)); 0 when there is none.</summary>
    public int LanguageID
    {
        get { return ReadLanguage(l => l.Val?.Value); }
        set { SetRunProperty(s => SetLanguage(s, value, false)); }
    }

    /// <summary>The language of East Asian text in the style, as <see cref="LanguageID"/>.</summary>
    public int LanguageIDFarEast
    {
        get { return ReadLanguage(l => l.EastAsia?.Value); }
        set { SetRunProperty(s => SetLanguage(s, value, true)); }
    }

    private int ReadLanguage(Func<Languages, string> read)
    {
        string name = read(_style.StyleRunProperties?.Languages ?? new Languages());
        if (string.IsNullOrEmpty(name))
        {
            return 0;
        }
        try
        {
            return CultureInfo.GetCultureInfo(name).LCID;
        }
        catch (CultureNotFoundException)
        {
            return 0;
        }
    }

    private static void SetLanguage(StyleRunProperties properties, int id, bool eastAsia)
    {
        string name = null;
        if (id != 0)
        {
            try
            {
                name = CultureInfo.GetCultureInfo(id).Name;
            }
            catch (CultureNotFoundException)
            {
                throw new ArgumentOutOfRangeException(nameof(id), "The number is not a language Word knows.");
            }
        }
        Languages languages = properties.Languages ?? new Languages();
        StringValue value = name == null ? null : new StringValue(name);
        if (eastAsia)
        {
            languages.EastAsia = value;
        }
        else
        {
            languages.Val = value;
        }
        properties.Languages = languages.Val == null && languages.EastAsia == null && languages.Bidi == null ? null : languages;
    }

    //Changes the character formatting of the style and of the character style it is linked to, as Word does.
    private void SetRunProperty(Action<StyleRunProperties> change)
    {
        foreach (Style style in new[] { _style, LinkedStyleElement() }.Where(s => s != null))
        {
            if (style.StyleRunProperties == null)
            {
                style.StyleRunProperties = new StyleRunProperties();
            }
            change(style.StyleRunProperties);
            if (!style.StyleRunProperties.HasChildren)
            {
                style.StyleRunProperties = null;
            }
        }
    }

    //An on/off element is on unless it says val 0, false or off.
    internal static bool IsOn(OpenXmlElement element)
    {
        if (element == null)
        {
            return false;
        }
        string value = element.GetAttributes().Where(a => a.LocalName == "val").Select(a => a.Value).FirstOrDefault();
        return value == null || (value != "0" && value != "false" && value != "off");
    }

    private string FindId(string name)
    {
        return StyleHelper.FindStyleId(_style, name, out StyleValues _);
    }

    private string SetBaseStyle(string name)
    {
        if (name == null)
        {
            _style.BasedOn = null;
            return null;
        }
        _style.BasedOn = new BasedOn() { Val = FindId(name) };

        //The character formatting of the new base is given to the linked style too, as Word does.
        Style linked = LinkedStyleElement();
        StyleRunProperties baseFormatting = StyleHelper.FindStyle(StyleHelper.GetStyles(_style), name)?.StyleRunProperties;
        if (linked != null && baseFormatting != null)
        {
            linked.StyleRunProperties = (StyleRunProperties)baseFormatting.CloneNode(true);
        }
        return name;
    }

    private void SetNextParagraphStyle(string name)
    {
        _style.NextParagraphStyle = name == null ? null : new NextParagraphStyle() { Val = FindId(name) };
    }
}
