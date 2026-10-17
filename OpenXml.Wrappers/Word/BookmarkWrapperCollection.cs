using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class BookmarkWrapperCollection : IEnumerable
{
    private RangeWrapper _rangeWrapper;

    public BookmarkWrapperCollection(RangeWrapper rangeWrapper)
    {
        _rangeWrapper = rangeWrapper;
    }

    public BookmarkWrapper this[int index]
    {
        get
        {
            try
            {
                return new BookmarkWrapper(_rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).ElementAt(index));
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new ArgumentOutOfRangeException(String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index), ex);
            }
        }
    }

    public BookmarkWrapper this[string name]
    {
        get
        {
            BookmarkStart bookmarkStart = _rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).Where(e => e.Name.Value == name).FirstOrDefault();
            if (bookmarkStart != null)
            {
                return new BookmarkWrapper(bookmarkStart);
            }
            return null;
        }
    }

    public int Count
    {
        get { return _rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).Count(); }
    }

    public void DeleteAll()
    {
        //It is important that the list is fetched since we are deleting items inside the iteration.
        //Each bookmark is removed as a whole (its start and its matching end) so that no dangling start or end markers are left behind
        //when the two markers are not both inside the range.
        IList<BookmarkStart> bookmarkStartList = _rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).ToList();
        foreach (BookmarkStart bookmarkStart in bookmarkStartList)
        {
            BookmarkEnd bookmarkEnd = bookmarkStart.GetRoot().Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id.Value == bookmarkStart.Id.Value);
            bookmarkStart.Parent.RemoveChild(bookmarkStart);
            if (bookmarkEnd != null)
            {
                bookmarkEnd.Parent.RemoveChild(bookmarkEnd);
            }
        }
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new BookmarkWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class BookmarkWrapperEnumerator : IEnumerator, IDisposable
    {
        private BookmarkWrapperCollection _bookmarklWrapperCollection;
        private int _index;
        private BookmarkWrapper _current;

        public BookmarkWrapperEnumerator(BookmarkWrapperCollection bookmarkWrapperCollection)
        {
            _bookmarklWrapperCollection = bookmarkWrapperCollection;
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
            if (++_index >= _bookmarklWrapperCollection.Count)
                return false;
            else
                _current = _bookmarklWrapperCollection[_index];
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
            _bookmarklWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
