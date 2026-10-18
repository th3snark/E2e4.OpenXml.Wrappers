using System;
using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class BookmarkWrapper : System.IEquatable<BookmarkWrapper>
{
    BookmarkStart _bookmarkStart;
    BookmarkEnd _bookmarkEnd;

    public BookmarkWrapper(BookmarkStart bookmarkStart)
    {
        _bookmarkStart = bookmarkStart;
        _bookmarkEnd = bookmarkStart.GetBookmarkEnd();
    }

    /// <summary>
    /// A range over the element that contains the bookmark (usually its paragraph).
    /// </summary>
    public RangeWrapper Parent
    {
        get { return new RangeWrapper(_bookmarkStart.Parent); }
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

    /// <summary>The character position where the bookmark starts (<c>Bookmark.Start</c> in Word).</summary>
    public int Start
    {
        get { return Range.Start; }
    }

    /// <summary>The character position where the bookmark ends (<c>Bookmark.End</c> in Word).</summary>
    public int End
    {
        get { return Range.End; }
    }

    /// <summary>True if the bookmark covers no text, only a point in the document (<c>Bookmark.Empty</c> in Word).</summary>
    public bool Empty
    {
        get { return Start == End; }
    }

    /// <summary>
    /// Adds a bookmark with another name over the same range and returns it (<c>Bookmark.Copy</c> in Word). A bookmark of that name is moved.
    /// </summary>
    /// <exception cref="ArgumentException">The name is not valid.</exception>
    public BookmarkWrapper Copy(string name)
    {
        RangeWrapper range = Range;
        return new BookmarkWrapperCollection(range).Add(name, range);
    }

    public void Delete()
    {
        _bookmarkStart.Parent.RemoveChild(_bookmarkStart);
        _bookmarkEnd.Parent.RemoveChild(_bookmarkEnd);
    }

    #region Equality

    // Two wrappers are equal when they wrap the same underlying element(s), so a new wrapper obtained for the same item compares equal.

    public override bool Equals(object obj)
    {
        return Equals(obj as BookmarkWrapper);
    }

    public bool Equals(BookmarkWrapper other)
    {
        return !ReferenceEquals(other, null) && ReferenceEquals(_bookmarkStart, other._bookmarkStart);
    }

    public override int GetHashCode()
    {
        return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_bookmarkStart);
    }

    public static bool operator ==(BookmarkWrapper left, BookmarkWrapper right)
    {
        return ReferenceEquals(left, null) ? ReferenceEquals(right, null) : left.Equals(right);
    }

    public static bool operator !=(BookmarkWrapper left, BookmarkWrapper right)
    {
        return !(left == right);
    }

    #endregion
}
