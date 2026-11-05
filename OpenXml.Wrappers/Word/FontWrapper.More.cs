using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The mark that is put by every character of the text, as <c>WdEmphasisMark</c> in Word.
/// </summary>
public enum EmphasisMark
{
    /// <summary>No mark.</summary>
    None = 0,

    /// <summary>A solid dot over each character.</summary>
    OverSolid = 1,

    /// <summary>A comma over each character.</summary>
    OverComma = 2,

    /// <summary>A white circle over each character.</summary>
    OverWhiteCircle = 3,

    /// <summary>A solid dot under each character.</summary>
    UnderSolid = 4
}

/// <summary>
/// The complex script and East Asian members of the font, the fonts by script, <c>Grow</c> and <c>Shrink</c>, and <c>Duplicate</c>.
/// </summary>
public partial class FontWrapper
{
    /// <summary>Bold complex script text (Arabic, Hebrew and so on; <c>Font.BoldBi</c> in Word).</summary>
    public bool? BoldBi
    {
        get { return GetOnOff<BoldComplexScript>(e => e.Val); }
        set { SetOnOff("bCs", value, off => new BoldComplexScript() { Val = off }, properties => ReadOnOff<BoldComplexScript>(properties, e => e.Val)); }
    }

    /// <summary>Italic complex script text (<c>Font.ItalicBi</c> in Word).</summary>
    public bool? ItalicBi
    {
        get { return GetOnOff<ItalicComplexScript>(e => e.Val); }
        set { SetOnOff("iCs", value, off => new ItalicComplexScript() { Val = off }, properties => ReadOnOff<ItalicComplexScript>(properties, e => e.Val)); }
    }

    /// <summary>The size of complex script text in points (<c>Font.SizeBi</c> in Word), which may be a half point.</summary>
    public double? SizeBi
    {
        get
        {
            return GetValue<double?>(properties => ReadSizeBi(properties) ?? (_direct ? (double?)null : DefaultSize));
        }
        set
        {
            if (value.HasValue && (value.Value <= 0 || value.Value > 1638))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The font size must be greater than 0 and no more than 1638 points.");
            }

            SetValue("szCs", value, properties => ReadSizeBi(properties) ?? DefaultSize, points => new FontSizeComplexScript()
            {
                Val = Math.Round(points * 2, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture)
            });
        }
    }

    private static double? ReadSizeBi(OpenXmlCompositeElement properties)
    {
        FontSizeComplexScript size = properties?.Elements<FontSizeComplexScript>().FirstOrDefault();
        return size?.Val == null ? (double?)null : Convert.ToDouble(size.Val.Value, CultureInfo.InvariantCulture) / 2;
    }

    /// <summary>
    /// The colour of complex script text as a colour index (<c>Font.ColorIndexBi</c> in Word). Word keeps one colour for text of every script,
    /// so this is the same as <see cref="ColorIndex"/>.
    /// </summary>
    public ColorIndex? ColorIndexBi
    {
        get { return ColorIndex; }
        set { ColorIndex = value; }
    }

    /// <summary>The font of Latin text (<c>Font.NameAscii</c> in Word); setting it leaves the other scripts as they are. A font that comes from the theme reads <c>+Body</c> or <c>+Headings</c>.</summary>
    public string NameAscii
    {
        get { return GetFontName(f => f.Ascii?.Value, f => f.AsciiTheme?.Value, false); }
        set { SetFontName(value, f => f.Ascii = value, f => f.AsciiTheme = null, f => f.Ascii = null); }
    }

    /// <summary>The font of other text, the characters that are not Latin, East Asian or complex script (<c>Font.NameOther</c> in Word).</summary>
    public string NameOther
    {
        get { return GetFontName(f => f.HighAnsi?.Value, f => f.HighAnsiTheme?.Value, false); }
        set { SetFontName(value, f => f.HighAnsi = value, f => f.HighAnsiTheme = null, f => f.HighAnsi = null); }
    }

    /// <summary>The font of East Asian text (<c>Font.NameFarEast</c> in Word).</summary>
    public string NameFarEast
    {
        get { return GetFontName(f => f.EastAsia?.Value, f => f.EastAsiaTheme?.Value, false); }
        set { SetFontName(value, f => f.EastAsia = value, f => f.EastAsiaTheme = null, f => f.EastAsia = null); }
    }

    /// <summary>The font of complex script text (<c>Font.NameBi</c> in Word). A font that comes from the theme reads <c>+Body CS</c> or <c>+Headings CS</c>.</summary>
    public string NameBi
    {
        get { return GetFontName(f => f.ComplexScript?.Value, f => f.ComplexScriptTheme?.Value, true); }
        set { SetFontName(value, f => f.ComplexScript = value, f => f.ComplexScriptTheme = null, f => f.ComplexScript = null); }
    }

    private string GetFontName(Func<RunFonts, string> name, Func<RunFonts, ThemeFontValues?> theme, bool complexScript)
    {
        return GetValue((properties, resolver) =>
        {
            RunFonts fonts = properties?.Elements<RunFonts>().FirstOrDefault();
            if (fonts == null)
            {
                return _direct ? null : "Times New Roman";
            }

            ThemeFontValues? themeFont = theme(fonts);
            if (!_direct && themeFont.HasValue)
            {
                //Word shows the font of the theme by the part of the theme it is in, not by its name.
                bool major = themeFont.Value == ThemeFontValues.MajorAscii || themeFont.Value == ThemeFontValues.MajorHighAnsi
                    || themeFont.Value == ThemeFontValues.MajorEastAsia || themeFont.Value == ThemeFontValues.MajorBidi;
                return (major ? "+Headings" : "+Body") + (complexScript ? " CS" : string.Empty);
            }
            return name(fonts) ?? (_direct ? null : "Times New Roman");
        });
    }

    private void SetFontName(string value, Action<RunFonts> set, Action<RunFonts> clearTheme, Action<RunFonts> clear)
    {
        Apply(properties =>
        {
            RunFonts fonts = properties.Elements<RunFonts>().FirstOrDefault();
            if (value == null)
            {
                if (fonts != null)
                {
                    clear(fonts);
                    clearTheme(fonts);
                    if (!fonts.GetAttributes().Any())
                    {
                        properties.RemoveRunProperty("rFonts");
                    }
                }
                return;
            }

            if (fonts == null)
            {
                fonts = new RunFonts();
                properties.SetRunProperty(fonts);
            }
            set(fonts);
            clearTheme(fonts);
        });
    }

    /// <summary>The mark over or under each character, as <c>Font.EmphasisMark</c> in Word.</summary>
    public EmphasisMark? EmphasisMark
    {
        get { return GetValue<EmphasisMark?>(ReadEmphasis); }
        set
        {
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            SetValue("em", value, ReadEmphasis, mark => new Emphasis() { Val = ToEmphasisValue(mark) });
        }
    }

    private static EmphasisMark? ReadEmphasis(OpenXmlCompositeElement properties)
    {
        EmphasisMarkValues? value = properties?.Elements<Emphasis>().FirstOrDefault()?.Val?.Value;
        if (!value.HasValue || value.Value == EmphasisMarkValues.None)
        {
            return Word.EmphasisMark.None;
        }
        return value.Value == EmphasisMarkValues.Dot ? Word.EmphasisMark.OverSolid
            : value.Value == EmphasisMarkValues.Comma ? Word.EmphasisMark.OverComma
            : value.Value == EmphasisMarkValues.Circle ? Word.EmphasisMark.OverWhiteCircle
            : Word.EmphasisMark.UnderSolid;
    }

    private static EmphasisMarkValues ToEmphasisValue(EmphasisMark mark)
    {
        return mark == Word.EmphasisMark.OverSolid ? EmphasisMarkValues.Dot
            : mark == Word.EmphasisMark.OverComma ? EmphasisMarkValues.Comma
            : mark == Word.EmphasisMark.OverWhiteCircle ? EmphasisMarkValues.Circle
            : mark == Word.EmphasisMark.UnderSolid ? EmphasisMarkValues.UnderDot
            : EmphasisMarkValues.None;
    }

    /// <summary>
    /// True if the characters are not lined up on the character grid of the page (<c>Font.DisableCharacterSpaceGrid</c> in Word), which only
    /// matters in a document laid out on a grid.
    /// </summary>
    public bool? DisableCharacterSpaceGrid
    {
        get { return GetValue<bool?>(ReadNoGrid); }
        set { SetValue("snapToGrid", value, ReadNoGrid, disabled => new SnapToGrid() { Val = disabled ? OpenXmlExtensions.OffValue() : (OnOffValue)true }); }
    }

    private static bool? ReadNoGrid(OpenXmlCompositeElement properties)
    {
        SnapToGrid grid = properties?.Elements<SnapToGrid>().FirstOrDefault();
        return grid != null && grid.Val != null && grid.Val.HasValue && !grid.Val.Value;
    }

    //The sizes that Grow and Shrink step through above 8 points, as Word does; above 80 points they go in steps of 10.
    private static readonly double[] SizeSteps = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 80 };

    private const double LargestSize = 1638;

    /// <summary>
    /// Makes the characters one step larger, as <c>Font.Grow</c> in Word (the size of complex script text follows): to the next size in 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26,
    /// 28, 36, 48, 72, 80, 90, 100 and so on in steps of 10, or one point more below 8, up to 1638 points. Each run grows from its own size.
    /// </summary>
    public void Grow()
    {
        Apply((properties, inherited) => Step(properties, inherited, GrownSize));
    }

    /// <summary>
    /// Makes the characters one step smaller, as <c>Font.Shrink</c> in Word, through the sizes of <see cref="Grow"/>, or one point less up to
    /// 8, down to 1 point.
    /// </summary>
    public void Shrink()
    {
        Apply((properties, inherited) => Step(properties, inherited, ShrunkSize));
    }

    private static void Step(OpenXmlCompositeElement properties, Func<OpenXmlCompositeElement> inherited, Func<double, double> next)
    {
        double size = ReadSize(properties) ?? ReadSize(inherited()) ?? DefaultSize;
        properties.RemoveRunProperty("sz");
        properties.RemoveRunProperty("szCs");
        double changed = next(size);

        //As in Word the size of complex script text changes with it.
        string halfPoints = Math.Round(changed * 2, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture);
        if (ReadSize(inherited()) != changed)
        {
            properties.SetRunProperty(new FontSize() { Val = halfPoints });
        }
        if (ReadSizeBi(inherited()) != changed)
        {
            properties.SetRunProperty(new FontSizeComplexScript() { Val = halfPoints });
        }
    }

    private static double GrownSize(double size)
    {
        if (size < SizeSteps[0])
        {
            return size + 1;
        }
        foreach (double step in SizeSteps)
        {
            if (step > size)
            {
                return step;
            }
        }
        return Math.Min(LargestSize, (Math.Floor(size / 10) + 1) * 10);
    }

    private static double ShrunkSize(double size)
    {
        if (size <= SizeSteps[0])
        {
            return Math.Max(1, size - 1);
        }
        double previous = SizeSteps[0];
        foreach (double step in SizeSteps)
        {
            if (step >= size)
            {
                return previous;
            }
            previous = step;
        }
        double lower = Math.Ceiling(size / 10) * 10 - 10;
        return lower > SizeSteps[SizeSteps.Length - 1] ? lower : SizeSteps[SizeSteps.Length - 1];
    }

    /// <summary>
    /// A copy of the font that does not change with the text, as <c>Font.Duplicate</c> in Word: it keeps the formatting the text has now.
    /// Assign it to <see cref="RangeWrapper.Font"/> of a range to give that range the formatting.
    /// </summary>
    public FontWrapper Duplicate()
    {
        List<Run> runs = _style != null ? new List<Run>() : _range.GetRuns().ToList();
        Paragraph holder = new Paragraph();
        if (_style != null)
        {
            holder.AppendChild(new Run(StyledProperties(), new Text("x")));
        }
        else if (runs.Count == 0)
        {
            List<Paragraph> marks = _range.GetParagraphs().ToList();
            if (marks.Count > 0)
            {
                FormattingResolver resolver = new FormattingResolver(marks[0]);
                holder.AppendChild(new Run(AsRunProperties(resolver.ResolveParagraphMark(marks[0])), new Text("x")));
            }
        }
        else
        {
            FormattingResolver resolver = new FormattingResolver(runs[0]);
            foreach (Run run in runs)
            {
                holder.AppendChild(new Run(AsRunProperties(_direct ? run.RunProperties : resolver.ResolveRun(run)), new Text("x")));
            }
        }
        return new FontWrapper(new RangeWrapper(new Body(holder)), null, _direct);
    }

    private RunProperties StyledProperties()
    {
        FormattingResolver resolver = new FormattingResolver(_style);
        return AsRunProperties(_direct ? _style.StyleRunProperties : resolver.ResolveStyleRun(_style));
    }

    private static RunProperties AsRunProperties(OpenXmlCompositeElement properties)
    {
        RunProperties copy = new RunProperties();
        if (properties != null)
        {
            foreach (OpenXmlElement child in properties.ChildElements.Where(e => !(e is RunStyle)))
            {
                copy.AppendChild(child.CloneNode(true));
            }
        }
        return copy;
    }

    //Gives the text of the range what the font has, for the properties that it has a value for (mixed ones are left alone).
    internal void CopyTo(FontWrapper target)
    {
        target.Bold = Bold ?? target.Bold;
        target.Italic = Italic ?? target.Italic;
        target.StrikeThrough = StrikeThrough ?? target.StrikeThrough;
        target.DoubleStrikeThrough = DoubleStrikeThrough ?? target.DoubleStrikeThrough;
        target.Underline = Underline ?? target.Underline;
        target.UnderlineColor = UnderlineColor ?? target.UnderlineColor;
        target.Superscript = Superscript ?? target.Superscript;
        target.Subscript = Subscript ?? target.Subscript;
        target.AllCaps = AllCaps ?? target.AllCaps;
        target.SmallCaps = SmallCaps ?? target.SmallCaps;
        target.Hidden = Hidden ?? target.Hidden;
        target.Emboss = Emboss ?? target.Emboss;
        target.Engrave = Engrave ?? target.Engrave;
        target.Outline = Outline ?? target.Outline;
        target.Shadow = Shadow ?? target.Shadow;
        target.Spacing = Spacing ?? target.Spacing;
        target.Scaling = Scaling ?? target.Scaling;
        target.Position = Position ?? target.Position;
        target.Kerning = Kerning ?? target.Kerning;
        target.Size = Size ?? target.Size;
        target.SizeBi = SizeBi ?? target.SizeBi;
        target.BoldBi = BoldBi ?? target.BoldBi;
        target.ItalicBi = ItalicBi ?? target.ItalicBi;
        target.EmphasisMark = EmphasisMark ?? target.EmphasisMark;
        target.DisableCharacterSpaceGrid = DisableCharacterSpaceGrid ?? target.DisableCharacterSpaceGrid;
        target.Color = Color ?? target.Color;
        target.NameAscii = NameAscii ?? target.NameAscii;
        target.NameOther = NameOther ?? target.NameOther;
        target.NameFarEast = NameFarEast ?? target.NameFarEast;
        target.NameBi = NameBi ?? target.NameBi;
    }
}
