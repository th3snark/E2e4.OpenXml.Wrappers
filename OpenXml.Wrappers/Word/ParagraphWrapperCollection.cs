using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class ParagraphWrapperCollection : IEnumerable
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
        return _range.Start.GetElements<Paragraph>(_range.End).ToList();
    }

    public ParagraphWrapper Add()
    {
        Paragraph paragraph = new Paragraph();
        _range.LastParagraph.Paragraph.InsertAfterSelf(paragraph);
        return new ParagraphWrapper(paragraph);
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new ParagraphWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class ParagraphWrapperEnumerator : IEnumerator, IDisposable
    {
        private ParagraphWrapperCollection _paragraphWrapperCollection;
        private int _index;
        private ParagraphWrapper _current;

        public ParagraphWrapperEnumerator(ParagraphWrapperCollection paragraphWrapperCollection)
        {
            _paragraphWrapperCollection = paragraphWrapperCollection;
            _index = -1;
            _current = null;
        }

        #region IEnumerator Members

        public object Current
        {
            get { return _current; }
        }

        public bool MoveNext()
        {
            if (++_index >= _paragraphWrapperCollection.Count)
                return false;
            else
                _current = _paragraphWrapperCollection[_index];
            return true;
        }

        public void Reset()
        {
            _current = null;
            _index = -1;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            _paragraphWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
