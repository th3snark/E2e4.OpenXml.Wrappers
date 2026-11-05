using System;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The members of <see cref="ParagraphFormatWrapper"/> that <c>Paragraphs</c> has in Word too: reading gives null where the paragraphs
/// differ, and setting or calling applies to every one. Each is the same as on <see cref="Format"/>.
/// </summary>
public partial class ParagraphWrapperCollection
{
    /// <summary>
    /// Adds an empty paragraph before the paragraph a range is in, or splits the paragraph at the range when it is collapsed in the middle of
    /// text, as <c>Paragraphs.Add(Range)</c> in Word: a paragraph mark is put at the start of the range. The result is the paragraph the new
    /// mark ends.
    /// </summary>
    public ParagraphWrapper Add(RangeWrapper range)
    {
        Guard.ThrowIfNull(range, nameof(range));
        RangeWrapper at = range.Duplicate();
        at.Collapse(CollapseDirection.Start);
        at.InsertParagraph();
        return at.Paragraphs[0];
    }

    /// <summary>The alignment of the paragraphs.</summary>
    public ParagraphAlignment? Alignment
    {
        get { return Format.Alignment; }
        set { Format.Alignment = value; }
    }

    /// <summary>The distance in points between the left margin and the left edge of the paragraphs.</summary>
    public double? LeftIndent
    {
        get { return Format.LeftIndent; }
        set { Format.LeftIndent = value; }
    }

    /// <summary>The distance in points between the right margin and the right edge of the paragraphs.</summary>
    public double? RightIndent
    {
        get { return Format.RightIndent; }
        set { Format.RightIndent = value; }
    }

    /// <summary>The indent in points of the first line of the paragraphs; negative for a hanging indent.</summary>
    public double? FirstLineIndent
    {
        get { return Format.FirstLineIndent; }
        set { Format.FirstLineIndent = value; }
    }

    /// <summary>The space in points before the paragraphs.</summary>
    public double? SpaceBefore
    {
        get { return Format.SpaceBefore; }
        set { Format.SpaceBefore = value; }
    }

    /// <summary>The space in points after the paragraphs.</summary>
    public double? SpaceAfter
    {
        get { return Format.SpaceAfter; }
        set { Format.SpaceAfter = value; }
    }

    /// <summary>Whether the space before the paragraphs is set automatically.</summary>
    public bool? SpaceBeforeAuto
    {
        get { return Format.SpaceBeforeAuto; }
        set { Format.SpaceBeforeAuto = value; }
    }

    /// <summary>Whether the space after the paragraphs is set automatically.</summary>
    public bool? SpaceAfterAuto
    {
        get { return Format.SpaceAfterAuto; }
        set { Format.SpaceAfterAuto = value; }
    }

    /// <summary>The line spacing in points.</summary>
    public double? LineSpacing
    {
        get { return Format.LineSpacing; }
        set { Format.LineSpacing = value; }
    }

    /// <summary>The rule used for the line spacing.</summary>
    public LineSpacingRule? LineSpacingRule
    {
        get { return Format.LineSpacingRule; }
        set { Format.LineSpacingRule = value; }
    }

    /// <summary>Whether the paragraphs are kept on the same page as the one after.</summary>
    public bool? KeepWithNext
    {
        get { return Format.KeepWithNext; }
        set { Format.KeepWithNext = value; }
    }

    /// <summary>Whether all the lines of the paragraphs are kept on the same page.</summary>
    public bool? KeepTogether
    {
        get { return Format.KeepTogether; }
        set { Format.KeepTogether = value; }
    }

    /// <summary>Whether the paragraphs start a new page.</summary>
    public bool? PageBreakBefore
    {
        get { return Format.PageBreakBefore; }
        set { Format.PageBreakBefore = value; }
    }

    /// <summary>Whether the first or last line of the paragraphs is kept from being left alone on a page.</summary>
    public bool? WidowControl
    {
        get { return Format.WidowControl; }
        set { Format.WidowControl = value; }
    }

    /// <summary>Whether the paragraphs are left out when lines are numbered.</summary>
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

    /// <summary>Whether the paragraphs are hyphenated when the document is set to hyphenate automatically.</summary>
    public bool? Hyphenation
    {
        get { return Format.Hyphenation; }
        set { Format.Hyphenation = value; }
    }

    /// <summary>The outline level of the paragraphs.</summary>
    public ParagraphOutlineLevel? OutlineLevel
    {
        get { return Format.OutlineLevel; }
        set { Format.OutlineLevel = value; }
    }

    /// <summary>The tab stops of the paragraphs.</summary>
    public TabStopWrapperCollection TabStops
    {
        get { return Format.TabStops; }
    }

    /// <summary>The borders of the paragraphs.</summary>
    public BordersWrapper Borders
    {
        get { return Format.Borders; }
    }

    /// <summary>The shading of the paragraphs.</summary>
    public ShadingWrapper Shading
    {
        get { return Format.Shading; }
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

    /// <summary>The direction the paragraphs read in.</summary>
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

    /// <summary>Which lines follow the shape of the text box the paragraphs are in.</summary>
    public TextboxTightWrap? TextboxTightWrap
    {
        get { return Format.TextboxTightWrap; }
        set { Format.TextboxTightWrap = value; }
    }

    /// <summary>Removes the direct paragraph formatting so that the styles apply.</summary>
    public void Reset()
    {
        Format.Reset();
    }

    /// <summary>Puts 12 points of space before the paragraphs.</summary>
    public void OpenUp()
    {
        Format.OpenUp();
    }

    /// <summary>Takes away the space before the paragraphs.</summary>
    public void CloseUp()
    {
        Format.CloseUp();
    }

    /// <summary>Takes away the space before the paragraphs if they have some, otherwise puts 12 points there.</summary>
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

    /// <summary>Moves the left indent by a number of default tab stops.</summary>
    public void TabIndent(int count)
    {
        Format.TabIndent(count);
    }

    /// <summary>Moves the left indent by a number of default tab stops and hangs the first line out.</summary>
    public void TabHangingIndent(int count)
    {
        Format.TabHangingIndent(count);
    }

    /// <summary>Indents each paragraph to the next default tab stop.</summary>
    public void Indent()
    {
        foreach (ParagraphWrapper paragraph in this)
        {
            paragraph.Indent();
        }
    }

    /// <summary>Moves the left indent of each paragraph back to the default tab stop before it.</summary>
    public void Outdent()
    {
        foreach (ParagraphWrapper paragraph in this)
        {
            paragraph.Outdent();
        }
    }

    /// <summary>
    /// Adds 6 points of space before and after each paragraph, as <c>Paragraphs.IncreaseSpacing</c> in Word.
    /// </summary>
    public void IncreaseSpacing()
    {
        ChangeSpacing(6);
    }

    /// <summary>
    /// Takes 6 points of space away before and after each paragraph, but not below 0, as <c>Paragraphs.DecreaseSpacing</c> in Word.
    /// </summary>
    public void DecreaseSpacing()
    {
        ChangeSpacing(-6);
    }

    private void ChangeSpacing(double points)
    {
        foreach (ParagraphWrapper paragraph in this)
        {
            ParagraphFormatWrapper format = paragraph.Format;
            format.SpaceBefore = Math.Max(0, (format.SpaceBefore ?? 0) + points);
            format.SpaceAfter = Math.Max(0, (format.SpaceAfter ?? 0) + points);
        }
    }

    /// <summary>Moves each paragraph one level down in the outline; see <see cref="ParagraphWrapper.OutlineDemote"/>.</summary>
    public void OutlineDemote()
    {
        foreach (ParagraphWrapper paragraph in this)
        {
            paragraph.OutlineDemote();
        }
    }

    /// <summary>Moves each paragraph one level up in the outline; see <see cref="ParagraphWrapper.OutlinePromote"/>.</summary>
    public void OutlinePromote()
    {
        foreach (ParagraphWrapper paragraph in this)
        {
            paragraph.OutlinePromote();
        }
    }

    /// <summary>Makes each paragraph body text; see <see cref="ParagraphWrapper.OutlineDemoteToBody"/>.</summary>
    public void OutlineDemoteToBody()
    {
        foreach (ParagraphWrapper paragraph in this)
        {
            paragraph.OutlineDemoteToBody();
        }
    }
}
