using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>The three galleries of list templates, as <c>WdListGalleryType</c> in Word (with the same values).</summary>
public enum ListGalleryType
{
    /// <summary>The bullets.</summary>
    Bullet = 1,

    /// <summary>The numbers.</summary>
    Number = 2,

    /// <summary>The outlines.</summary>
    Outline = 3
}

/// <summary>
/// A gallery of list templates, as <c>ListGallery</c> in Word: seven templates that can be applied to the paragraphs of any document.
/// The templates are Word 16's own; changing one is not kept.
/// </summary>
public class ListGalleryWrapper
{
    private const int TemplatesPerGallery = 7;

    private readonly ListGalleryType _type;

    internal ListGalleryWrapper(ListGalleryType type)
    {
        _type = type;
    }

    /// <summary>Which gallery this is.</summary>
    public ListGalleryType GalleryType
    {
        get { return _type; }
    }

    /// <summary>The templates of the gallery, as <c>ListGallery.ListTemplates</c> in Word.</summary>
    public ListGalleryTemplateCollection ListTemplates
    {
        get { return new ListGalleryTemplateCollection(_type); }
    }

    /// <summary>True if the template at the index (0-based) has been changed from Word's own, as <c>ListGallery.Modified</c> in Word; never here.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no template at the index.</exception>
    public bool Modified(int index)
    {
        if (index < 0 || index >= TemplatesPerGallery)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }
        return false;
    }

    /// <summary>Puts the template at the index (0-based) back to Word's own, as <c>ListGallery.Reset</c> in Word; there is nothing to put back.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no template at the index.</exception>
    public void Reset(int index)
    {
        Modified(index);
    }

    internal static List<AbstractNum> ReadDefinitions()
    {
        XDocument document;
        using (Stream stream = NewDocument.OpenTemplate("ListGalleries.xml"))
        {
            document = XDocument.Load(stream);
        }
        //The elements are read as the numbering part's own, so that they have a parent that is not in any document: that is how a template of
        //a gallery is told from one of a document.
        Numbering holder = new Numbering(document.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting));
        return holder.Elements<AbstractNum>().ToList();
    }
}

/// <summary>The templates of one gallery, as <c>ListTemplates</c> of a <c>ListGallery</c> in Word.</summary>
public class ListGalleryTemplateCollection : IEnumerable<ListTemplateWrapper>
{
    private readonly ListGalleryType _type;

    internal ListGalleryTemplateCollection(ListGalleryType type)
    {
        _type = type;
    }

    /// <summary>The number of templates: seven.</summary>
    public int Count
    {
        get { return 7; }
    }

    /// <summary>The template at the index (0-based), as <c>ListTemplates.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no template at the index.</exception>
    public ListTemplateWrapper this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new ListTemplateWrapper(ListGalleryWrapper.ReadDefinitions()[((int)_type - 1) * Count + index]);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ListTemplateWrapper> GetEnumerator()
    {
        for (int index = 0; index < Count; index++)
        {
            yield return this[index];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>The galleries of list templates, as <c>ListGalleries</c> in Word.</summary>
public class ListGalleryWrapperCollection : IEnumerable<ListGalleryWrapper>
{
    internal ListGalleryWrapperCollection()
    {
    }

    /// <summary>The number of galleries: three.</summary>
    public int Count
    {
        get { return 3; }
    }

    /// <summary>The gallery of a kind, as <c>ListGalleries.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a value of <see cref="ListGalleryType"/>.</exception>
    public ListGalleryWrapper this[ListGalleryType type]
    {
        get
        {
            if (!EnumCompat.IsDefined(type))
            {
                throw new ArgumentOutOfRangeException(nameof(type));
            }
            return new ListGalleryWrapper(type);
        }
    }

    /// <summary>The gallery at the index (0-based): bullets, numbers, outlines.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no gallery at the index.</exception>
    public ListGalleryWrapper this[int index]
    {
        get
        {
            if (index < 0 || index >= Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return this[(ListGalleryType)(index + 1)];
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ListGalleryWrapper> GetEnumerator()
    {
        for (int index = 0; index < Count; index++)
        {
            yield return this[index];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

public partial class ApplicationWrapper
{
    /// <summary>The galleries of list templates, as <c>Application.ListGalleries</c> in Word.</summary>
    public ListGalleryWrapperCollection ListGalleries
    {
        get { return new ListGalleryWrapperCollection(); }
    }

    /// <summary>
    /// Takes the characters out of a text that Word does not keep in a text it hands on, as <c>Application.CleanString</c> in Word: most
    /// control characters and the non-breaking space become a space, a line feed becomes a paragraph mark, code 30 becomes a hyphen, and the
    /// bell, code 31, the not sign and the pilcrow are dropped. The rule is the one measured for each of the characters 0 to 255 alone.
    /// </summary>
    public string CleanString(string text)
    {
        Guard.ThrowIfNull(text);
        StringBuilder result = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (c == 7 || c == 31 || c == 172 || c == 182)
            {
                continue;
            }
            if (c == 10)
            {
                result.Append('\r');
            }
            else if (c == 30)
            {
                result.Append('-');
            }
            else if (c <= 6 || c == 8 || c == 11 || c == 12 || (c >= 14 && c <= 29) || c == 160 || c == 176 || c == 183)
            {
                result.Append(' ');
            }
            else
            {
                result.Append(c);
            }
        }
        return result.ToString();
    }
}
