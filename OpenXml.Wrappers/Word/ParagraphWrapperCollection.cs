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
                return new ParagraphWrapper(GetElementList()[index]);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new ArgumentOutOfRangeException(String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index), ex);
            }
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    private IList<Paragraph> GetElementList()
    {
        return _range.IsSpan ? _range.GetTouchedParagraphs() : _range.StartElement.GetElements<Paragraph>(_range.EndElement).ToList();
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
        foreach (Paragraph element in GetElementList())
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
