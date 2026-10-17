using System;
using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class BookmarkWrapper
{
    BookmarkStart _bookmarkStart;
    BookmarkEnd _bookmarkEnd;

    public BookmarkWrapper(BookmarkStart bookmarkStart)
    {
        _bookmarkStart = bookmarkStart;
        _bookmarkEnd = bookmarkStart.GetBookmarkEnd();
    }

    public string Name
    {
        get { return _bookmarkStart.Name.Value; }
    }

    public int Id
    {
        get { return Convert.ToInt32(_bookmarkStart.Id.Value, CultureInfo.InvariantCulture); }
    }

    public RangeWrapper Range
    {
        get { return new RangeWrapper(_bookmarkStart, _bookmarkEnd); }
    }

    public void Delete()
    {
        _bookmarkStart.Parent.RemoveChild(_bookmarkStart);
        _bookmarkEnd.Parent.RemoveChild(_bookmarkEnd);
    }
}
