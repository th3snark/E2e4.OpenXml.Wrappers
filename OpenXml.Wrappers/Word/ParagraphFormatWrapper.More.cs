using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The direction in which a paragraph reads, as <c>WdReadingOrder</c> in Word.
/// </summary>
public enum ReadingOrder
{
    /// <summary>Right to left.</summary>
    RightToLeft = 0,

    /// <summary>Left to right. It is the default.</summary>
    LeftToRight = 1
}

/// <summary>
/// How characters of different sizes line up in a line, as <c>WdBaselineAlignment</c> in Word.
/// </summary>
public enum BaselineAlignment
{
    /// <summary>The tops of the characters line up.</summary>
    Top = 0,

    /// <summary>The centres line up.</summary>
    Center = 1,

    /// <summary>The baselines line up.</summary>
    Baseline = 2,

    /// <summary>The bottoms line up.</summary>
    Far = 3,

    /// <summary>Word decides (the default).</summary>
    Auto = 4
}

/// <summary>
/// Which lines of a paragraph in a text box follow the shape of the box, as <c>WdTextboxTightWrap</c> in Word.
/// </summary>
public enum TextboxTightWrap
{
    /// <summary>No lines. It is the default.</summary>
    None = 0,

    /// <summary>All lines.</summary>
    AllLines = 1,

    /// <summary>The first and last lines.</summary>
    FirstAndLastLine = 2,

    /// <summary>The first line only.</summary>
    FirstLineOnly = 3,

    /// <summary>The last line only.</summary>
    LastLineOnly = 4
}

/// <summary>
/// The shortcuts for spacing and indenting, the East Asian and reading order switches, and <c>Duplicate</c>.
/// </summary>
public partial class ParagraphFormatWrapper
{
    /// <summary>
    /// The borders of the paragraphs, as <c>ParagraphFormat.Borders</c> in Word: the four edges. It is the direct setting only (the borders a
    /// paragraph style gives are not read).
    /// </summary>
    public BordersWrapper Borders
    {
        get
        {
            if (_style != null)
            {
                return new BordersWrapper(_style, false);
            }
            return new BordersWrapper(() => GetParagraphs().Cast<OpenXmlElement>().ToList(), false);
        }
    }

    /// <summary>
    /// The shading of the paragraphs, as <c>ParagraphFormat.Shading</c> in Word. It is the direct setting only.
    /// </summary>
    public ShadingWrapper Shading
    {
        get
        {
            if (_style != null)
            {
                return StyleWrapper.ParagraphShading(_style);
            }
            return new ShadingWrapper(create => GetParagraphs().Select(paragraph =>
            {
                ParagraphProperties properties = create ? paragraph.GetOrAddParagraphProperties() : paragraph.ParagraphProperties;
                Shading shading = properties?.Shading;
                if (shading == null && create)
                {
                    shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };
                    properties.Shading = shading;
                }
                return shading;
            }).ToList());
        }
    }

    //Word opens up a paragraph by a line, 12 points.
    private const double OpenSpace = 12;

    /// <summary>Puts 12 points of space before the paragraphs, as <c>ParagraphFormat.OpenUp</c> in Word.</summary>
    public void OpenUp()
    {
        SpaceBefore = OpenSpace;
    }

    /// <summary>Takes away the space before the paragraphs, as <c>ParagraphFormat.CloseUp</c> in Word.</summary>
    public void CloseUp()
    {
        SpaceBefore = 0;
    }

    /// <summary>
    /// Takes away the space before the paragraphs if they have some, otherwise puts 12 points there, as <c>ParagraphFormat.OpenOrCloseUp</c>
    /// in Word. Paragraphs that differ are opened.
    /// </summary>
    public void OpenOrCloseUp()
    {
        SpaceBefore = SpaceBefore > 0 ? 0 : OpenSpace;
    }

    /// <summary>Single line spacing (<c>ParagraphFormat.Space1</c> in Word).</summary>
    public void Space1()
    {
        LineSpacingRule = Word.LineSpacingRule.SingleSpacing;
    }

    /// <summary>One and a half line spacing (<c>ParagraphFormat.Space15</c> in Word).</summary>
    public void Space15()
    {
        LineSpacingRule = Word.LineSpacingRule.OnePointFive;
    }

    /// <summary>Double line spacing (<c>ParagraphFormat.Space2</c> in Word).</summary>
    public void Space2()
    {
        LineSpacingRule = Word.LineSpacingRule.DoubleSpacing;
    }

    /// <summary>
    /// Moves the left indent by a number of default tab stops, as <c>ParagraphFormat.TabIndent</c> in Word: forward to the next stop when the
    /// count is positive, back to the stop before when it is negative, and not below 0.
    /// </summary>
    public void TabIndent(int count)
    {
        LeftIndent = TabStop(LeftIndent ?? 0, count);
    }

    /// <summary>
    /// Moves the left indent by a number of default tab stops and hangs the first line out to the margin, as
    /// <c>ParagraphFormat.TabHangingIndent</c> in Word.
    /// </summary>
    public void TabHangingIndent(int count)
    {
        double left = TabStop(LeftIndent ?? 0, count);
        LeftIndent = left;
        FirstLineIndent = -left;
    }

    //The indent a number of default tab stops from the current one, which is not a multiple of the stop when the indent was set by hand.
    private double TabStop(double indent, int count)
    {
        double stop = DefaultTabStopPoints();
        if (count == 0)
        {
            return indent;
        }
        double stops = indent / stop;
        double target = count > 0 ? Math.Floor(stops + 1e-9) + count : Math.Ceiling(stops - 1e-9) + count;
        return Math.Max(0, target * stop);
    }

    private double DefaultTabStopPoints()
    {
        Paragraph first = GetParagraphs().FirstOrDefault();
        MainDocumentPart main = _style != null ? StyleHelper.GetMainDocumentPart(_style) : first == null ? null : StyleHelper.GetMainDocumentPart(first);
        long? twips = main?.DocumentSettingsPart?.Settings?.GetFirstChild<DefaultTabStop>()?.Val?.Value;
        return twips.HasValue && twips.Value > 0 ? twips.Value / 20.0 : 36;
    }

    /// <summary>Whether text may break in the middle of a word to wrap (<c>ParagraphFormat.WordWrap</c> in Word); true by default.</summary>
    public bool? WordWrap
    {
        get { return GetSwitch(properties => properties?.WordWrap, e => e.Val, true); }
        set { SetSwitch(value, true, off => new WordWrap() { Val = off }, properties => properties?.WordWrap, e => e.Val, properties => properties.WordWrap = null, (properties, element) => properties.WordWrap = element); }
    }

    /// <summary>Whether the right indent follows the East Asian character grid (<c>ParagraphFormat.AutoAdjustRightIndent</c> in Word); true by default.</summary>
    public bool? AutoAdjustRightIndent
    {
        get { return GetSwitch(properties => properties?.AdjustRightIndent, e => e.Val, true); }
        set { SetSwitch(value, true, off => new AdjustRightIndent() { Val = off }, properties => properties?.AdjustRightIndent, e => e.Val, properties => properties.AdjustRightIndent = null, (properties, element) => properties.AdjustRightIndent = element); }
    }

    /// <summary>Whether space is added between East Asian and Latin text (<c>ParagraphFormat.AddSpaceBetweenFarEastAndAlpha</c> in Word); true by default.</summary>
    public bool? AddSpaceBetweenFarEastAndAlpha
    {
        get { return GetSwitch(properties => properties?.AutoSpaceDE, e => e.Val, true); }
        set { SetSwitch(value, true, off => new AutoSpaceDE() { Val = off }, properties => properties?.AutoSpaceDE, e => e.Val, properties => properties.AutoSpaceDE = null, (properties, element) => properties.AutoSpaceDE = element); }
    }

    /// <summary>Whether space is added between East Asian text and numbers (<c>ParagraphFormat.AddSpaceBetweenFarEastAndDigit</c> in Word); true by default.</summary>
    public bool? AddSpaceBetweenFarEastAndDigit
    {
        get { return GetSwitch(properties => properties?.AutoSpaceDN, e => e.Val, true); }
        set { SetSwitch(value, true, off => new AutoSpaceDN() { Val = off }, properties => properties?.AutoSpaceDN, e => e.Val, properties => properties.AutoSpaceDN = null, (properties, element) => properties.AutoSpaceDN = element); }
    }

    /// <summary>Whether the East Asian line break rules are used (<c>ParagraphFormat.FarEastLineBreakControl</c> in Word); true by default.</summary>
    public bool? FarEastLineBreakControl
    {
        get { return GetSwitch(properties => properties?.Kinsoku, e => e.Val, true); }
        set { SetSwitch(value, true, off => new Kinsoku() { Val = off }, properties => properties?.Kinsoku, e => e.Val, properties => properties.Kinsoku = null, (properties, element) => properties.Kinsoku = element); }
    }

    /// <summary>Whether punctuation may hang past the margin (<c>ParagraphFormat.HangingPunctuation</c> in Word); true by default.</summary>
    public bool? HangingPunctuation
    {
        get { return GetSwitch(properties => properties?.OverflowPunctuation, e => e.Val, true); }
        set { SetSwitch(value, true, off => new OverflowPunctuation() { Val = off }, properties => properties?.OverflowPunctuation, e => e.Val, properties => properties.OverflowPunctuation = null, (properties, element) => properties.OverflowPunctuation = element); }
    }

    /// <summary>Whether half-width punctuation is at the top of the line (<c>ParagraphFormat.HalfWidthPunctuationOnTopOfLine</c> in Word); false by default.</summary>
    public bool? HalfWidthPunctuationOnTopOfLine
    {
        get { return GetSwitch(properties => properties?.TopLinePunctuation, e => e.Val, false); }
        set { SetSwitch(value, false, off => new TopLinePunctuation() { Val = off }, properties => properties?.TopLinePunctuation, e => e.Val, properties => properties.TopLinePunctuation = null, (properties, element) => properties.TopLinePunctuation = element); }
    }

    /// <summary>
    /// True if the lines of the paragraphs do not snap to the line grid of the page (<c>ParagraphFormat.DisableLineHeightGrid</c> in Word); false by default.
    /// </summary>
    public bool? DisableLineHeightGrid
    {
        get
        {
            bool? snaps = GetSwitch(properties => properties?.SnapToGrid, e => e.Val, true);
            return snaps.HasValue ? !snaps.Value : (bool?)null;
        }
        set
        {
            SetSwitch(value.HasValue ? !value.Value : (bool?)null, true, off => new SnapToGrid() { Val = off }, properties => properties?.SnapToGrid, e => e.Val, properties => properties.SnapToGrid = null, (properties, element) => properties.SnapToGrid = element);
        }
    }

    /// <summary>The direction the paragraphs read in (<c>ParagraphFormat.ReadingOrder</c> in Word).</summary>
    public ReadingOrder? ReadingOrder
    {
        get
        {
            bool? rightToLeft = GetSwitch(properties => properties?.BiDi, e => e.Val, false);
            return rightToLeft.HasValue ? (rightToLeft.Value ? Word.ReadingOrder.RightToLeft : Word.ReadingOrder.LeftToRight) : (ReadingOrder?)null;
        }
        set
        {
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            SetSwitch(value.HasValue ? value.Value == Word.ReadingOrder.RightToLeft : (bool?)null, false, off => new BiDi() { Val = off }, properties => properties?.BiDi, e => e.Val, properties => properties.BiDi = null, (properties, element) => properties.BiDi = element);
        }
    }

    /// <summary>How characters of different sizes line up in a line (<c>ParagraphFormat.BaseLineAlignment</c> in Word).</summary>
    public BaselineAlignment? BaseLineAlignment
    {
        get { return GetValue<BaselineAlignment?>(ReadBaseline); }
        set
        {
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            ApplyUnlessInherited(properties => properties.TextAlignment = null, inherited => !value.HasValue || ReadBaseline(inherited) == value,
                properties => properties.TextAlignment = new TextAlignment() { Val = ToTextAlignment(value.Value) });
        }
    }

    private static BaselineAlignment? ReadBaseline(ParagraphProperties properties)
    {
        VerticalTextAlignmentValues? value = properties?.TextAlignment?.Val?.Value;
        if (!value.HasValue)
        {
            return BaselineAlignment.Auto;
        }
        return value.Value == VerticalTextAlignmentValues.Top ? BaselineAlignment.Top
            : value.Value == VerticalTextAlignmentValues.Center ? BaselineAlignment.Center
            : value.Value == VerticalTextAlignmentValues.Baseline ? BaselineAlignment.Baseline
            : value.Value == VerticalTextAlignmentValues.Bottom ? BaselineAlignment.Far
            : BaselineAlignment.Auto;
    }

    private static VerticalTextAlignmentValues ToTextAlignment(BaselineAlignment value)
    {
        return value == BaselineAlignment.Top ? VerticalTextAlignmentValues.Top
            : value == BaselineAlignment.Center ? VerticalTextAlignmentValues.Center
            : value == BaselineAlignment.Baseline ? VerticalTextAlignmentValues.Baseline
            : value == BaselineAlignment.Far ? VerticalTextAlignmentValues.Bottom
            : VerticalTextAlignmentValues.Auto;
    }

    /// <summary>Which lines follow the shape of the text box the paragraphs are in (<c>ParagraphFormat.TextboxTightWrap</c> in Word).</summary>
    public TextboxTightWrap? TextboxTightWrap
    {
        get { return GetValue<TextboxTightWrap?>(ReadTightWrap); }
        set
        {
            if (value.HasValue && !EnumCompat.IsDefined(value.Value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            ApplyUnlessInherited(properties => properties.TextBoxTightWrap = null, inherited => !value.HasValue || ReadTightWrap(inherited) == value,
                properties => properties.TextBoxTightWrap = new TextBoxTightWrap() { Val = ToTightWrap(value.Value) });
        }
    }

    private static TextboxTightWrap? ReadTightWrap(ParagraphProperties properties)
    {
        TextBoxTightWrapValues? value = properties?.TextBoxTightWrap?.Val?.Value;
        if (!value.HasValue)
        {
            return Word.TextboxTightWrap.None;
        }
        return value.Value == TextBoxTightWrapValues.AllLines ? Word.TextboxTightWrap.AllLines
            : value.Value == TextBoxTightWrapValues.FirstAndLastLine ? Word.TextboxTightWrap.FirstAndLastLine
            : value.Value == TextBoxTightWrapValues.FirstLineOnly ? Word.TextboxTightWrap.FirstLineOnly
            : value.Value == TextBoxTightWrapValues.LastLineOnly ? Word.TextboxTightWrap.LastLineOnly
            : Word.TextboxTightWrap.None;
    }

    private static TextBoxTightWrapValues ToTightWrap(TextboxTightWrap value)
    {
        return value == Word.TextboxTightWrap.AllLines ? TextBoxTightWrapValues.AllLines
            : value == Word.TextboxTightWrap.FirstAndLastLine ? TextBoxTightWrapValues.FirstAndLastLine
            : value == Word.TextboxTightWrap.FirstLineOnly ? TextBoxTightWrapValues.FirstLineOnly
            : value == Word.TextboxTightWrap.LastLineOnly ? TextBoxTightWrapValues.LastLineOnly
            : TextBoxTightWrapValues.None;
    }

    //An on and off property that is on, or off, when the paragraph says nothing: def.
    private bool? GetSwitch<T>(Func<ParagraphProperties, T> find, Func<T, OnOffValue> getValue, bool def) where T : OpenXmlElement
    {
        return GetValue<bool?>(properties =>
        {
            T element = find(properties);
            if (element == null)
            {
                return _direct ? false : def;
            }
            OnOffValue value = getValue(element);
            return value == null || !value.HasValue || value.Value;
        });
    }

    //As Word writes them: a value the paragraph has anyway from its styles (or by default) is not written, on is the bare element and off is w:val="0".
    private void SetSwitch<T>(bool? value, bool def, Func<OnOffValue, T> create, Func<ParagraphProperties, T> find, Func<T, OnOffValue> getValue, Action<ParagraphProperties> remove, Action<ParagraphProperties, T> set) where T : OpenXmlElement
    {
        ApplyUnlessInherited(remove, inherited =>
        {
            T element = find(inherited);
            OnOffValue on = element == null ? null : getValue(element);
            bool isOn = element == null ? def : on == null || !on.HasValue || on.Value;
            return !value.HasValue || isOn == value.Value;
        }, properties => set(properties, create(value.Value ? null : OpenXmlExtensions.OffValue())));
    }

    /// <summary>
    /// A copy of the formatting that does not change with the paragraphs, as <c>ParagraphFormat.Duplicate</c> in Word: it keeps the
    /// formatting they have now. Assign it to <see cref="RangeWrapper.ParagraphFormat"/> of a range to give that range the formatting.
    /// </summary>
    public ParagraphFormatWrapper Duplicate()
    {
        List<Paragraph> holders = new List<Paragraph>();
        if (_style != null)
        {
            holders.Add(Holder(_direct ? CopyStyleProperties() : new FormattingResolver(_style).ResolveStyleParagraph(_style)));
        }
        else
        {
            IList<Paragraph> paragraphs = GetParagraphs();
            FormattingResolver resolver = paragraphs.Count == 0 ? null : new FormattingResolver(paragraphs[0]);
            foreach (Paragraph paragraph in paragraphs)
            {
                holders.Add(Holder(_direct ? paragraph.ParagraphProperties : resolver.ResolveParagraph(paragraph)));
            }
        }
        return new ParagraphFormatWrapper(new RangeWrapper(new Body(holders))) { _direct = _direct };
    }

    private static Paragraph Holder(ParagraphProperties properties)
    {
        Paragraph paragraph = new Paragraph();
        if (properties != null)
        {
            ParagraphProperties copy = new ParagraphProperties();
            foreach (OpenXmlElement child in properties.ChildElements.Where(e => !(e is ParagraphStyleId || e is SectionProperties || e is ParagraphMarkRunProperties || e is ParagraphPropertiesChange)))
            {
                copy.AppendChild(child.CloneNode(true));
            }
            if (copy.HasChildren)
            {
                paragraph.AppendChild(copy);
            }
        }
        return paragraph;
    }

    //Gives the paragraphs what the format has, for the properties it has a single value for (mixed ones are left alone).
    internal void CopyTo(ParagraphFormatWrapper target)
    {
        target.Alignment = Alignment ?? target.Alignment;
        target.LeftIndent = LeftIndent ?? target.LeftIndent;
        target.RightIndent = RightIndent ?? target.RightIndent;
        target.FirstLineIndent = FirstLineIndent ?? target.FirstLineIndent;
        target.SpaceBefore = SpaceBefore ?? target.SpaceBefore;
        target.SpaceAfter = SpaceAfter ?? target.SpaceAfter;
        target.SpaceBeforeAuto = SpaceBeforeAuto ?? target.SpaceBeforeAuto;
        target.SpaceAfterAuto = SpaceAfterAuto ?? target.SpaceAfterAuto;
        Word.LineSpacingRule? rule = LineSpacingRule;
        double? spacing = LineSpacing;
        if (rule.HasValue)
        {
            target.LineSpacingRule = rule;
            if (spacing.HasValue && (rule == Word.LineSpacingRule.AtLeast || rule == Word.LineSpacingRule.Exactly || rule == Word.LineSpacingRule.Multiple))
            {
                target.LineSpacing = spacing;
            }
        }
        target.KeepWithNext = KeepWithNext ?? target.KeepWithNext;
        target.KeepTogether = KeepTogether ?? target.KeepTogether;
        target.PageBreakBefore = PageBreakBefore ?? target.PageBreakBefore;
        target.WidowControl = WidowControl ?? target.WidowControl;
        target.NoLineNumber = NoLineNumber ?? target.NoLineNumber;
        target.MirrorIndents = MirrorIndents ?? target.MirrorIndents;
        target.Hyphenation = Hyphenation ?? target.Hyphenation;
        target.OutlineLevel = OutlineLevel ?? target.OutlineLevel;
        target.WordWrap = WordWrap ?? target.WordWrap;
        target.AutoAdjustRightIndent = AutoAdjustRightIndent ?? target.AutoAdjustRightIndent;
        target.AddSpaceBetweenFarEastAndAlpha = AddSpaceBetweenFarEastAndAlpha ?? target.AddSpaceBetweenFarEastAndAlpha;
        target.AddSpaceBetweenFarEastAndDigit = AddSpaceBetweenFarEastAndDigit ?? target.AddSpaceBetweenFarEastAndDigit;
        target.FarEastLineBreakControl = FarEastLineBreakControl ?? target.FarEastLineBreakControl;
        target.HangingPunctuation = HangingPunctuation ?? target.HangingPunctuation;
        target.HalfWidthPunctuationOnTopOfLine = HalfWidthPunctuationOnTopOfLine ?? target.HalfWidthPunctuationOnTopOfLine;
        target.DisableLineHeightGrid = DisableLineHeightGrid ?? target.DisableLineHeightGrid;
        target.ReadingOrder = ReadingOrder ?? target.ReadingOrder;
        target.BaseLineAlignment = BaseLineAlignment ?? target.BaseLineAlignment;
        target.TextboxTightWrap = TextboxTightWrap ?? target.TextboxTightWrap;
    }
}
