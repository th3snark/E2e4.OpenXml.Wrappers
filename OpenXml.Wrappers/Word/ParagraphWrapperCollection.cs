using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class ParagraphWrapperCollection : IEnumerable<ParagraphWrapper>
{
    private RangeWrapper _range;

    public ParagraphWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    public RangeWrapper Parent
    {
        get { return _range; }
    }

    public ParagraphWrapper this[int index]
    {
        get
        {
            try
            {
                Guard.ThrowIfNegative(index);
                return new ParagraphWrapper(GetElements().Skip(index).First());
            }
            catch (InvalidOperationException)
            {
                throw new ArgumentOutOfRangeException(nameof(index), String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new ArgumentOutOfRangeException(nameof(index), String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
        }
    }

    public int Count
    {
        get { return GetElements().Count(); }
    }

    /// <summary>The first paragraph, or null if there are none.</summary>
    public ParagraphWrapper First
    {
        get
        {
            Paragraph paragraph = GetElements().FirstOrDefault();
            return paragraph == null ? null : new ParagraphWrapper(paragraph);
        }
    }

    /// <summary>The last paragraph, or null if there are none.</summary>
    public ParagraphWrapper Last
    {
        get
        {
            Paragraph paragraph = GetElements().LastOrDefault();
            return paragraph == null ? null : new ParagraphWrapper(paragraph);
        }
    }

    /// <summary>
    /// The paragraph formatting of all the paragraphs, as <c>Paragraphs.Format</c> in Word: reading gives null for a property where they
    /// differ and setting applies to every one.
    /// </summary>
    public ParagraphFormatWrapper Format
    {
        get { return _range.ParagraphFormat; }
    }

    /// <summary>The style of all the paragraphs (null where they differ). Setting applies it to every one.</summary>
    public string Style
    {
        get { return _range.Style; }
        set { _range.Style = value; }
    }

    //The paragraphs of the range, produced as they are asked for so that the first few do not need the rest to be looked at.
    private IEnumerable<Paragraph> GetElements()
    {
        if (_range.IsSpan)
        {
            return _range.GetTouchedParagraphs();
        }
        return ReferenceEquals(_range.StartElement, _range.EndElement)
            ? StoryMap.ParagraphsIn(_range.StartElement)
            : _range.StartElement.GetElements<Paragraph>(_range.EndElement);
    }

    public ParagraphWrapper Add()
    {
        Paragraph paragraph = new Paragraph();
        _range.LastParagraph.Paragraph.InsertAfterSelf(paragraph);
        return new ParagraphWrapper(paragraph);
    }

    #region IEnumerable Members

    public IEnumerator<ParagraphWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (Paragraph element in GetElements().ToList())
        {
            yield return new ParagraphWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
