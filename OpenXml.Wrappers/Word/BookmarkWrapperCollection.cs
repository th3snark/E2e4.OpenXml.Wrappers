using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class BookmarkWrapperCollection : IEnumerable<BookmarkWrapper>
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

    private IList<BookmarkStart> GetElementList()
    {
        return _rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).ToList();
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

    public IEnumerator<BookmarkWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (BookmarkStart element in GetElementList())
        {
            yield return new BookmarkWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
