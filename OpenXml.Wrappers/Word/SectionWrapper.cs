using System.Linq;

using DocumentFormat.OpenXml;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A section of the document, as <c>Section</c> in Word.
/// </summary>
/// <remarks>
/// The section properties of every section but the last are in the paragraph that ends the section, and those of the last section are at
/// the end of the body. A document may not have the latter, in which case the last section has the defaults of Word until a
/// <see cref="PageSetup"/> value is set.
/// </remarks>
public class SectionWrapper : System.IEquatable<SectionWrapper>
{
    private readonly Body _body;
    private readonly SectionProperties _sectionProperties;

    internal SectionWrapper(Body body, SectionProperties sectionProperties)
    {
        _body = body;
        _sectionProperties = sectionProperties;
    }

    /// <summary>
    /// The section properties of the section, or null for the last section of a document that does not have any.
    /// </summary>
    public SectionProperties SectionProperties
    {
        get { return GetSectionProperties(false); }
    }

    /// <summary>
    /// The 0-based position of the section in the document.
    /// </summary>
    public int Index
    {
        get { return new SectionWrapperCollection(_body).ToList().IndexOf(this); }
    }

    /// <summary>
    /// The sections of the document this section is in (<c>Section.Parent</c> in Word).
    /// </summary>
    public SectionWrapperCollection Parent
    {
        get { return new SectionWrapperCollection(_body); }
    }

    /// <summary>
    /// The text of the section, as <c>Section.Range</c> in Word: from the end of the section before it to the paragraph that ends it (the
    /// section break), or to the end of the document for the last section. It is a range of characters of the document.
    /// </summary>
    public RangeWrapper Range
    {
        get
        {
            int sectionIndex = Index;
            int current = 0;
            OpenXmlElement first = null;
            OpenXmlElement last = null;
            foreach (OpenXmlElement block in _body.Elements().Where(e => !(e is SectionProperties)))
            {
                if (current == sectionIndex)
                {
                    first ??= block;
                    last = block;
                }
                if (block is Paragraph paragraph && paragraph.ParagraphProperties?.SectionProperties != null)
                {
                    if (current == sectionIndex)
                    {
                        break;
                    }
                    current++;
                }
            }

            OpenXmlElement root = (OpenXmlElement)_body.Ancestors<Document>().FirstOrDefault() ?? _body;
            RangeWrapper range = new RangeWrapper(root);
            if (first == null)
            {
                int length = range.StoryLength;
                range.SetRange(length, length);
            }
            else
            {
                range.SetRange(new RangeWrapper(first).Start, new RangeWrapper(last).End);
            }
            return range;
        }
    }

    /// <summary>
    /// The page borders of the section, as <c>Section.Borders</c> in Word. See <see cref="PageSetupWrapper.Borders"/>.
    /// </summary>
    public BordersWrapper Borders
    {
        get { return PageSetup.Borders; }
    }

    /// <summary>
    /// The page setup of the section.
    /// </summary>
    public PageSetupWrapper PageSetup
    {
        get { return new PageSetupWrapper(this); }
    }

    /// <summary>
    /// The headers of the section.
    /// </summary>
    public HeaderFooterWrapperCollection Headers
    {
        get { return new HeaderFooterWrapperCollection(this, true); }
    }

    /// <summary>
    /// The footers of the section.
    /// </summary>
    public HeaderFooterWrapperCollection Footers
    {
        get { return new HeaderFooterWrapperCollection(this, false); }
    }

    internal DocumentFormat.OpenXml.Packaging.MainDocumentPart GetMainDocumentPart()
    {
        return _body.Ancestors<Document>().FirstOrDefault()?.MainDocumentPart;
    }

    internal System.Collections.Generic.IList<SectionWrapper> Siblings()
    {
        return new SectionWrapperCollection(_body).ToList();
    }

    internal SectionProperties GetSectionProperties(bool create)
    {
        if (_sectionProperties != null)
        {
            return _sectionProperties;
        }

        //This is the last section: its properties are the last child of the body.
        SectionProperties properties = _body.Elements<SectionProperties>().LastOrDefault();
        if (properties == null && create)
        {
            properties = _body.AppendChild(new SectionProperties());
        }
        return properties;
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as SectionWrapper);
    }

    public bool Equals(SectionWrapper other)
    {
        if (ReferenceEquals(other, null))
        {
            return false;
        }
        SectionProperties mine = GetSectionProperties(false);
        SectionProperties theirs = other.GetSectionProperties(false);
        return mine != null || theirs != null ? ReferenceEquals(mine, theirs) : ReferenceEquals(_body, other._body);
    }

    public override int GetHashCode()
    {
        //The properties of the last section can appear later, so all of the sections hash by the body to stay stable.
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_body);
    }

    public static bool operator ==(SectionWrapper left, SectionWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(SectionWrapper left, SectionWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
