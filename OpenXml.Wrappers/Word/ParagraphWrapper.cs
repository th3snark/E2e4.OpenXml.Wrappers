using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class ParagraphWrapper : System.IEquatable<ParagraphWrapper>
{
    Paragraph _paragraph;

    public ParagraphWrapper(Paragraph paragraph)
    {
        _paragraph = paragraph;
    }

    public Paragraph Paragraph
    {
        get { return _paragraph; }
    }

    /// <summary>
    /// A range over the element that contains the paragraph: the table cell or the body.
    /// </summary>
    public RangeWrapper Parent
    {
        get { return new RangeWrapper(_paragraph.Parent); }
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_paragraph); }
    }

    /// <summary>
    /// The formatting of the paragraph, as <c>Paragraph.Format</c> in Word.
    /// </summary>
    public ParagraphFormatWrapper Format
    {
        get { return Range.ParagraphFormat; }
    }

    /// <summary>
    /// The text of the paragraph, ending with its paragraph mark as a carriage return, as <c>Paragraph.Range.Text</c> in Word (for a paragraph
    /// in a table cell the last one ends with the end-of-cell mark <c>\a</c>). Setting it keeps the paragraph mark: one at the end of the
    /// new text is ignored, so that <c>paragraph.Text = paragraph.Text</c> changes nothing.
    /// </summary>
    public string Text
    {
        get { return Range.Text; }
        set { Range.Text = value; }
    }

    /// <summary>The alignment of the paragraph. The same as <see cref="Format"/>.<c>Alignment</c>.</summary>
    public ParagraphAlignment Alignment
    {
        get { return Format.Alignment ?? ParagraphAlignment.Left; }
        set { Format.Alignment = value; }
    }

    /// <summary>The distance in points between the left margin and the left edge of the paragraph. The same as <see cref="Format"/>.<c>LeftIndent</c>.</summary>
    public double LeftIndent
    {
        get { return Format.LeftIndent ?? 0; }
        set { Format.LeftIndent = value; }
    }

    /// <summary>The distance in points between the right margin and the right edge of the paragraph. The same as <see cref="Format"/>.<c>RightIndent</c>.</summary>
    public double RightIndent
    {
        get { return Format.RightIndent ?? 0; }
        set { Format.RightIndent = value; }
    }

    /// <summary>The indent of the first line in points (negative for a hanging indent). The same as <see cref="Format"/>.<c>FirstLineIndent</c>.</summary>
    public double FirstLineIndent
    {
        get { return Format.FirstLineIndent ?? 0; }
        set { Format.FirstLineIndent = value; }
    }

    /// <summary>The line spacing in points. The same as <see cref="Format"/>.<c>LineSpacing</c>.</summary>
    public double LineSpacing
    {
        get { return Format.LineSpacing ?? 12; }
        set { Format.LineSpacing = value; }
    }

    /// <summary>The rule for the line spacing. The same as <see cref="Format"/>.<c>LineSpacingRule</c>.</summary>
    public LineSpacingRule LineSpacingRule
    {
        get { return Format.LineSpacingRule ?? LineSpacingRule.SingleSpacing; }
        set { Format.LineSpacingRule = value; }
    }

    /// <summary>Whether all the lines of the paragraph are kept on the same page. The same as <see cref="Format"/>.<c>KeepTogether</c>.</summary>
    public bool KeepTogether
    {
        get { return Format.KeepTogether == true; }
        set { Format.KeepTogether = value; }
    }

    /// <summary>Whether the paragraph starts a new page. The same as <see cref="Format"/>.<c>PageBreakBefore</c>.</summary>
    public bool PageBreakBefore
    {
        get { return Format.PageBreakBefore == true; }
        set { Format.PageBreakBefore = value; }
    }

    /// <summary>Whether the first and last lines are kept from being alone on a page. The same as <see cref="Format"/>.<c>WidowControl</c>.</summary>
    public bool WidowControl
    {
        get { return Format.WidowControl == true; }
        set { Format.WidowControl = value; }
    }

    /// <summary>The name of the paragraph style. The same as <see cref="RangeWrapper.Style"/> of <see cref="Range"/>.</summary>
    public string Style
    {
        get { return Range.Style; }
        set { Range.Style = value; }
    }

    /// <summary>
    /// The space before the paragraph in points, 0 when none is set. The same as <see cref="Format"/>.<c>SpaceBefore</c>.
    /// </summary>
    public double SpaceBefore
    {
        get { return Format.SpaceBefore ?? 0; }
        set { Format.SpaceBefore = value; }
    }

    /// <summary>
    /// The space after the paragraph in points, 0 when none is set. The same as <see cref="Format"/>.<c>SpaceAfter</c>.
    /// </summary>
    public double SpaceAfter
    {
        get { return Format.SpaceAfter ?? 0; }
        set { Format.SpaceAfter = value; }
    }

    /// <summary>
    /// The paragraph <paramref name="count"/> after this one in the document, including paragraphs in table cells, or null if there is none (<c>Paragraph.Next</c> in Word).
    /// </summary>
    public ParagraphWrapper Next(int count = 1)
    {
        return GetNeighbour(count);
    }

    /// <summary>
    /// The paragraph <paramref name="count"/> before this one in the document, or null if there is none (<c>Paragraph.Previous</c> in Word).
    /// </summary>
    public ParagraphWrapper Previous(int count = 1)
    {
        return GetNeighbour(-count);
    }

    private ParagraphWrapper GetNeighbour(int offset)
    {
        System.Collections.Generic.List<Paragraph> paragraphs = _paragraph.GetRoot().Descendants<Paragraph>().ToList();
        int index = paragraphs.IndexOf(_paragraph) + offset;
        return index < 0 || index >= paragraphs.Count ? null : new ParagraphWrapper(paragraphs[index]);
    }

    public void Delete()
    {
        Range.Delete();
    }

    public void InsertBreak(BreakValues type)
    {
        if (type == BreakValues.Page || type == BreakValues.Column || type == BreakValues.TextWrapping)
        {
            _paragraph.AppendChild(new Run(new Break() { Type = type }));
        }
        else
        {
            throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The break type '{0}' is not a recognised type.", type), nameof(type));
        }
    }

    public void DeleteBreaks()
    {
        foreach (Run run in _paragraph.Elements<Run>())
        {
            run.RemoveAllChildren<Break>();
        }
    }

    /// <summary>
    /// Ends a section with this paragraph, so that the paragraph is the last one of the section. <paramref name="start"/> says how the section
    /// that follows starts. (<see cref="RangeWrapper.InsertBreak"/> is Word's way to insert a section break, at a place in the text; this is
    /// the same break with the paragraph as its place.)
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The start is not a section start.</exception>
    public void InsertSectionBreak(SectionStart start = SectionStart.NewPage)
    {
        if (!EnumCompat.IsDefined(start))
        {
            throw new ArgumentOutOfRangeException(nameof(start));
        }

        EndSection(start == SectionStart.Continuous ? SectionMarkValues.Continuous
            : start == SectionStart.EvenPage ? SectionMarkValues.EvenPage
            : start == SectionStart.OddPage ? SectionMarkValues.OddPage
            : start == SectionStart.NewColumn ? SectionMarkValues.NextColumn
            : SectionMarkValues.NextPage);
    }

    //Ends a section with this paragraph and returns the properties of the section that ends here, which are a copy of the final section's.
    internal SectionProperties EndSection(SectionMarkValues markType)
    {
        //Add a new paragraph properties if one does not exist.
        ParagraphProperties paragraphProperties = _paragraph.ParagraphProperties;
        if (paragraphProperties == null)
        {
            paragraphProperties = new ParagraphProperties();
            _paragraph.PrependChild(paragraphProperties);
        }

        //Remove the existing section properties from the paragraph properties.
        SectionProperties sectionProperties = paragraphProperties.SectionProperties;
        if (sectionProperties != null)
        {
            sectionProperties.Parent.RemoveChild(sectionProperties);
        }

        //Get the last section properties in the body, clone it and then append it to the paragraph properties of this paragraph.
        SectionProperties finalSectionProperties = _paragraph.GetDocument().Body.Elements<SectionProperties>().First();
        SectionProperties newSectionProperties = (SectionProperties)finalSectionProperties.Clone();
        paragraphProperties.AppendChild(newSectionProperties);

        //The break type belongs to the section that starts after the break: the next section properties in the document.
        System.Collections.Generic.List<SectionProperties> all = _paragraph.GetRoot().Descendants<SectionProperties>().ToList();
        SectionProperties following = all[all.IndexOf(newSectionProperties) + 1];
        SectionType sectionType = following.GetFirstChild<SectionType>();
        if (sectionType == null)
        {
            sectionType = new SectionType();
            OpenXmlElement before = following.ChildElements.LastOrDefault(e => e is HeaderReference || e is FooterReference || e is FootnoteProperties || e is EndnoteDocumentWideProperties || e is EndnoteProperties);
            if (before != null)
            {
                before.InsertAfterSelf(sectionType);
            }
            else
            {
                following.PrependChild(sectionType);
            }
        }
        sectionType.Val = markType;
        return newSectionProperties;
    }
    public void CopyPropertiesTo(ParagraphWrapper paragraphWrapper)
    {
        //Make the target paragraph properties the same as the source paragraph,
        ParagraphProperties fromParagraphProperties = (ParagraphProperties)_paragraph.ParagraphProperties;
        if (fromParagraphProperties != null)
        {
            if (paragraphWrapper.Paragraph.ParagraphProperties != null)
            {
                paragraphWrapper.Paragraph.RemoveChild(paragraphWrapper.Paragraph.ParagraphProperties);
            }
            ParagraphProperties toParagraphProperties = (ParagraphProperties)fromParagraphProperties.Clone();
            paragraphWrapper.Paragraph.PrependChild(toParagraphProperties);
        }

        //Make the all the target runs properties the same as the run properties of the first run of the source paragraph.
        Run firstRun = _paragraph.Elements<Run>().FirstOrDefault();
        if (firstRun != null)
        {
            RunProperties fromRunProperties = firstRun.RunProperties;
            if (fromRunProperties != null)
            {
                foreach (Run toRun in paragraphWrapper.Paragraph.Elements<Run>())
                {
                    RunProperties toRunProperties = toRun.RunProperties;
                    if (toRunProperties != null)
                    {
                        toRun.RemoveChild(toRunProperties);
                    }
                    toRun.PrependChild((RunProperties)fromRunProperties.Clone());
                }
            }
        }
    }

    /// <summary>
    /// Whether the paragraph is kept on the same page as the next one; the same as <c>Format.KeepWithNext</c>.
    /// </summary>
    public bool KeepWithNext
    {
        get { return Format.KeepWithNext == true; }
        set { Format.KeepWithNext = value ? true : null; }
    }
    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as ParagraphWrapper);
    }

    public bool Equals(ParagraphWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_paragraph, other._paragraph);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_paragraph);
    }

    public static bool operator ==(ParagraphWrapper left, ParagraphWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(ParagraphWrapper left, ParagraphWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
