using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A header or footer of a section, as <c>HeaderFooter</c> in Word.
/// </summary>
/// <remarks>
/// A section without a header of its own shows that of the nearest section before it ("linked to previous"). <see cref="Range"/> follows
/// that link when reading, and creates the header when there is none to link to. Turning on <see cref="SectionWrapper.PageSetup"/>'s
/// <c>DifferentFirstPageHeaderFooter</c> is up to the caller: creating a first-page header does not switch it on, and the even-page headers
/// are only shown when the odd and even pages setting of the document is on, which the library does not expose.
/// </remarks>
public class HeaderFooterWrapper : System.IEquatable<HeaderFooterWrapper>
{
    private readonly SectionWrapper _section;
    private readonly bool _isHeader;
    private readonly HeaderFooterIndex _index;

    internal HeaderFooterWrapper(SectionWrapper section, bool isHeader, HeaderFooterIndex index)
    {
        _section = section;
        _isHeader = isHeader;
        _index = index;
    }

    /// <summary>
    /// The section the header or footer belongs to.
    /// </summary>
    public SectionWrapper Parent
    {
        get { return _section; }
    }

    public HeaderFooterIndex Index
    {
        get { return _index; }
    }

    public bool IsHeader
    {
        get { return _isHeader; }
    }

    /// <summary>
    /// True when the section has a header or footer of its own of this kind (not when it only shows that of an earlier section), as
    /// <c>HeaderFooter.Exists</c> in Word. Setting it to true gives the section an empty header or footer of its own, and false removes it.
    /// </summary>
    public bool Exists
    {
        get { return GetReference(_section) != null; }
        set
        {
            if (value == Exists)
            {
                return;
            }
            if (value)
            {
                Create(null);
            }
            else
            {
                RemoveReference();
            }
        }
    }

    /// <summary>
    /// The page numbers of the header or footer, as <c>HeaderFooter.PageNumbers</c> in Word.
    /// </summary>
    public PageNumbersWrapper PageNumbers
    {
        get { return new PageNumbersWrapper(this); }
    }

    /// <summary>
    /// True when the section shows the header or footer of the previous section. Always false for the first section. Setting it to true
    /// removes the header or footer of this section; setting it to false gives the section a copy of the one it was showing.
    /// </summary>
    /// <exception cref="InvalidOperationException">Set to true for the first section, which has nothing to link to.</exception>
    public bool LinkToPrevious
    {
        get { return _section.Index > 0 && GetReference(_section) == null; }
        set
        {
            if (value == LinkToPrevious)
            {
                return;
            }

            if (value)
            {
                if (_section.Index == 0)
                {
                    throw new InvalidOperationException("The first section has no previous section to link to.");
                }
                RemoveReference();
            }
            else
            {
                HeaderFooterWrapper source = FindEffective();
                OpenXmlCompositeElement content = source == null ? null : source.GetRoot();
                Create(content);
            }
        }
    }

    /// <summary>
    /// The content of the header or footer, as <c>HeaderFooter.Range</c> in Word. The header or footer is created if the section and the
    /// sections before it have none.
    /// </summary>
    public RangeWrapper Range
    {
        get
        {
            HeaderFooterWrapper source = FindEffective();
            OpenXmlCompositeElement root = source == null ? Create(null) : source.GetRoot();
            return new RangeWrapper(root);
        }
    }

    //The wrapper in this or the nearest earlier section that has the header or footer, or null.
    private HeaderFooterWrapper FindEffective()
    {
        IList<SectionWrapper> sections = _section.Siblings();
        for (int i = _section.Index; i >= 0; i--)
        {
            HeaderFooterWrapper candidate = new HeaderFooterWrapper(sections[i], _isHeader, _index);
            if (GetReference(sections[i]) != null)
            {
                return candidate;
            }
        }
        return null;
    }

    private HeaderFooterValues GetReferenceType()
    {
        if (_index == HeaderFooterIndex.FirstPage)
        {
            return HeaderFooterValues.First;
        }
        return _index == HeaderFooterIndex.EvenPages ? HeaderFooterValues.Even : HeaderFooterValues.Default;
    }

    private HeaderFooterReferenceType GetReference(SectionWrapper section)
    {
        SectionProperties properties = section.SectionProperties;
        if (properties == null)
        {
            return null;
        }

        HeaderFooterValues type = GetReferenceType();
        IEnumerable<HeaderFooterReferenceType> references = _isHeader
            ? properties.Elements<HeaderReference>()
            : properties.Elements<FooterReference>();
        return references.FirstOrDefault(r => (r.Type != null && r.Type.HasValue ? r.Type.Value : HeaderFooterValues.Default) == type);
    }

    private MainDocumentPart GetMainDocumentPart()
    {
        MainDocumentPart part = _section.GetMainDocumentPart();
        if (part == null)
        {
            throw new InvalidOperationException("The section is not part of a document, so it cannot have headers and footers.");
        }
        return part;
    }

    private OpenXmlCompositeElement GetRoot()
    {
        HeaderFooterReferenceType reference = GetReference(_section);
        OpenXmlPart part = GetMainDocumentPart().GetPartById(reference.Id.Value);
        return _isHeader ? (OpenXmlCompositeElement)((HeaderPart)part).Header : ((FooterPart)part).Footer;
    }

    //Adds a header or footer part to the section, with a copy of the content or else an empty paragraph, and returns its root.
    private OpenXmlCompositeElement Create(OpenXmlCompositeElement content)
    {
        MainDocumentPart main = GetMainDocumentPart();
        SectionProperties properties = _section.GetSectionProperties(true);

        //Replace a reference that is left over without a part.
        GetReference(_section)?.Remove();

        OpenXmlCompositeElement root;
        string id;
        if (_isHeader)
        {
            HeaderPart part = main.AddNewPart<HeaderPart>();
            part.Header = content == null ? new Header(new Paragraph()) : (Header)content.CloneNode(true);
            root = part.Header;
            id = main.GetIdOfPart(part);
            properties.PrependChild(new HeaderReference() { Type = GetReferenceType(), Id = id });
        }
        else
        {
            FooterPart part = main.AddNewPart<FooterPart>();
            part.Footer = content == null ? new Footer(new Paragraph()) : (Footer)content.CloneNode(true);
            root = part.Footer;
            id = main.GetIdOfPart(part);
            properties.PrependChild(new FooterReference() { Type = GetReferenceType(), Id = id });
        }
        return root;
    }

    private void RemoveReference()
    {
        HeaderFooterReferenceType reference = GetReference(_section);
        if (reference == null)
        {
            return;
        }

        string id = reference.Id.Value;
        reference.Remove();

        //Another section may use the same part.
        MainDocumentPart main = GetMainDocumentPart();
        bool stillUsed = main.Document.Descendants<HeaderFooterReferenceType>().Any(r => r.Id != null && r.Id.Value == id);
        if (!stillUsed)
        {
            main.DeletePart(id);
        }
    }

    #region Equality

    // Two wrappers are equal when they are the same kind of header or footer of the same section.

    public override bool Equals(object obj)
    {
        return Equals(obj as HeaderFooterWrapper);
    }

    public bool Equals(HeaderFooterWrapper other)
    {
        return !ReferenceEquals(other, null) && _section == other._section && _isHeader == other._isHeader && _index == other._index;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(_section, _isHeader, _index);
    }

    public static bool operator ==(HeaderFooterWrapper left, HeaderFooterWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(HeaderFooterWrapper left, HeaderFooterWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
