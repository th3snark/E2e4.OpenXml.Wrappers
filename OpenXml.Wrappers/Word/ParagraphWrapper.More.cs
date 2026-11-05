namespace OpenXml.Wrappers.Word;

/// <summary>
/// The members of <see cref="ParagraphFormatWrapper"/> that <c>Paragraph</c> has in Word too, each the same as on <see cref="Format"/>.
/// </summary>
public partial class ParagraphWrapper
{
    /// <summary>Whether the paragraph is left out when lines are numbered.</summary>
    public bool? NoLineNumber
    {
        get { return Format.NoLineNumber; }
        set { Format.NoLineNumber = value; }
    }

    /// <summary>Whether the left and right indents are swapped on facing pages.</summary>
    public bool? MirrorIndents
    {
        get { return Format.MirrorIndents; }
        set { Format.MirrorIndents = value; }
    }

    /// <summary>Whether the paragraph is hyphenated when the document is set to hyphenate automatically.</summary>
    public bool? Hyphenation
    {
        get { return Format.Hyphenation; }
        set { Format.Hyphenation = value; }
    }

    /// <summary>Whether the space before the paragraph is set automatically (as for web pages).</summary>
    public bool? SpaceBeforeAuto
    {
        get { return Format.SpaceBeforeAuto; }
        set { Format.SpaceBeforeAuto = value; }
    }

    /// <summary>Whether the space after the paragraph is set automatically (as for web pages).</summary>
    public bool? SpaceAfterAuto
    {
        get { return Format.SpaceAfterAuto; }
        set { Format.SpaceAfterAuto = value; }
    }

    /// <summary>The outline level of the paragraph.</summary>
    public ParagraphOutlineLevel? OutlineLevel
    {
        get { return Format.OutlineLevel; }
        set { Format.OutlineLevel = value; }
    }

    /// <summary>The tab stops of the paragraph.</summary>
    public TabStopWrapperCollection TabStops
    {
        get { return Format.TabStops; }
    }

    /// <summary>Whether text may break in the middle of a word to wrap.</summary>
    public bool? WordWrap
    {
        get { return Format.WordWrap; }
        set { Format.WordWrap = value; }
    }

    /// <summary>Whether the right indent follows the East Asian character grid.</summary>
    public bool? AutoAdjustRightIndent
    {
        get { return Format.AutoAdjustRightIndent; }
        set { Format.AutoAdjustRightIndent = value; }
    }

    /// <summary>Whether space is added between East Asian and Latin text.</summary>
    public bool? AddSpaceBetweenFarEastAndAlpha
    {
        get { return Format.AddSpaceBetweenFarEastAndAlpha; }
        set { Format.AddSpaceBetweenFarEastAndAlpha = value; }
    }

    /// <summary>Whether space is added between East Asian text and numbers.</summary>
    public bool? AddSpaceBetweenFarEastAndDigit
    {
        get { return Format.AddSpaceBetweenFarEastAndDigit; }
        set { Format.AddSpaceBetweenFarEastAndDigit = value; }
    }

    /// <summary>Whether the East Asian line break rules are used.</summary>
    public bool? FarEastLineBreakControl
    {
        get { return Format.FarEastLineBreakControl; }
        set { Format.FarEastLineBreakControl = value; }
    }

    /// <summary>Whether punctuation may hang past the margin.</summary>
    public bool? HangingPunctuation
    {
        get { return Format.HangingPunctuation; }
        set { Format.HangingPunctuation = value; }
    }

    /// <summary>Whether half-width punctuation is at the top of the line.</summary>
    public bool? HalfWidthPunctuationOnTopOfLine
    {
        get { return Format.HalfWidthPunctuationOnTopOfLine; }
        set { Format.HalfWidthPunctuationOnTopOfLine = value; }
    }

    /// <summary>Whether the lines do not snap to the line grid of the page.</summary>
    public bool? DisableLineHeightGrid
    {
        get { return Format.DisableLineHeightGrid; }
        set { Format.DisableLineHeightGrid = value; }
    }

    /// <summary>The direction the paragraph reads in.</summary>
    public ReadingOrder? ReadingOrder
    {
        get { return Format.ReadingOrder; }
        set { Format.ReadingOrder = value; }
    }

    /// <summary>How characters of different sizes line up in a line.</summary>
    public BaselineAlignment? BaseLineAlignment
    {
        get { return Format.BaseLineAlignment; }
        set { Format.BaseLineAlignment = value; }
    }

    /// <summary>Which lines follow the shape of the text box the paragraph is in.</summary>
    public TextboxTightWrap? TextboxTightWrap
    {
        get { return Format.TextboxTightWrap; }
        set { Format.TextboxTightWrap = value; }
    }

    /// <summary>The borders of the paragraph; the same as <see cref="Format"/>.<c>Borders</c>.</summary>
    public BordersWrapper Borders
    {
        get { return Format.Borders; }
    }

    /// <summary>The shading of the paragraph; the same as <see cref="Format"/>.<c>Shading</c>.</summary>
    public ShadingWrapper Shading
    {
        get { return Format.Shading; }
    }

    /// <summary>Removes the direct paragraph formatting so that the styles apply; see <see cref="ParagraphFormatWrapper.Reset"/>.</summary>
    public void Reset()
    {
        Format.Reset();
    }

    /// <summary>Puts 12 points of space before the paragraph.</summary>
    public void OpenUp()
    {
        Format.OpenUp();
    }

    /// <summary>Takes away the space before the paragraph.</summary>
    public void CloseUp()
    {
        Format.CloseUp();
    }

    /// <summary>Takes away the space before the paragraph if it has some, otherwise puts 12 points there.</summary>
    public void OpenOrCloseUp()
    {
        Format.OpenOrCloseUp();
    }

    /// <summary>Single line spacing.</summary>
    public void Space1()
    {
        Format.Space1();
    }

    /// <summary>One and a half line spacing.</summary>
    public void Space15()
    {
        Format.Space15();
    }

    /// <summary>Double line spacing.</summary>
    public void Space2()
    {
        Format.Space2();
    }

    /// <summary>Moves the left indent by a number of default tab stops; see <see cref="ParagraphFormatWrapper.TabIndent"/>.</summary>
    public void TabIndent(int count)
    {
        Format.TabIndent(count);
    }

    /// <summary>Moves the left indent by a number of default tab stops and hangs the first line out; see <see cref="ParagraphFormatWrapper.TabHangingIndent"/>.</summary>
    public void TabHangingIndent(int count)
    {
        Format.TabHangingIndent(count);
    }

    /// <summary>Indents the paragraph to the next default tab stop, as <c>Paragraph.Indent</c> in Word.</summary>
    public void Indent()
    {
        Format.TabIndent(1);
    }

    /// <summary>Moves the left indent back to the default tab stop before it, but not below 0, as <c>Paragraph.Outdent</c> in Word.</summary>
    public void Outdent()
    {
        Format.TabIndent(-1);
    }
}
