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

    public string Text
    {
        get { return Range.Text.TrimEnd('\r', '\a'); }
        set { Range.Text = value; }
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

    private SpacingBetweenLines GetOrAddSpacing()
    {
        //The typed properties put the elements in schema order.
        if (_paragraph.ParagraphProperties == null)
        {
            _paragraph.ParagraphProperties = new ParagraphProperties();
        }
        if (_paragraph.ParagraphProperties.SpacingBetweenLines == null)
        {
            _paragraph.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines();
        }
        return _paragraph.ParagraphProperties.SpacingBetweenLines;
    }
    /// <summary>
    /// The paragraph after this one in the document, including paragraphs in table cells, or null for the last paragraph (<c>Paragraph.Next</c> in Word).
    /// </summary>
    public ParagraphWrapper Next()
    {
        return GetNeighbour(1);
    }

    /// <summary>
    /// The paragraph before this one in the document, or null for the first paragraph (<c>Paragraph.Previous</c> in Word).
    /// </summary>
    public ParagraphWrapper Previous()
    {
        return GetNeighbour(-1);
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
    /// Ends a section with this paragraph. <paramref name="type"/> says how the section that follows starts: "nextPage", "continuous",
    /// "evenPage", "oddPage" or "nextColumn". Returns the properties of the section that ends here, which are a copy of the final section's.
    /// </summary>
    public SectionProperties InsertSectionBreak(string type)
    {
        SectionMarkValues markType;
        switch (type)
        {
            case "nextPage":
                markType = SectionMarkValues.NextPage;
                break;
            case "continuous":
                markType = SectionMarkValues.Continuous;
                break;
            case "evenPage":
                markType = SectionMarkValues.EvenPage;
                break;
            case "oddPage":
                markType = SectionMarkValues.OddPage;
                break;
            case "nextColumn":
                markType = SectionMarkValues.NextColumn;
                break;
            default:
                throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The section break type '{0}' is not a recognised type.", type), nameof(type));
        }

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
    public bool KeepNext
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
