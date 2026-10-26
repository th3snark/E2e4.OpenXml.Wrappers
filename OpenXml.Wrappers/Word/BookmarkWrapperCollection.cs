using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// How the bookmarks of a document are ordered, as <c>WdBookmarkSortBy</c> in Word (with the same values).
/// </summary>
public enum BookmarkSorting
{
    /// <summary>Alphabetically by name, ignoring case (an underscore sorts before letters).</summary>
    Name = 0,

    /// <summary>In the order they are in the document.</summary>
    Location = 1
}

/// <summary>
/// The bookmarks of a range, as <c>Bookmarks</c> in Word.
/// </summary>
/// <remarks>
/// As in Word the bookmarks are in alphabetical order by default (<see cref="DefaultSorting"/>), and a bookmark whose name starts with an
/// underscore is hidden and left out unless <see cref="ShowHidden"/> is true. Both settings belong to the document, so setting them on one
/// collection applies to every collection of the document.
/// </remarks>
public class BookmarkWrapperCollection : IEnumerable<BookmarkWrapper>
{
    //The settings are those of the document, held by the root of its tree.
    private sealed class Settings
    {
        public bool ShowHidden;
        public BookmarkSorting Sorting = BookmarkSorting.Name;
    }

    private static readonly ConditionalWeakTable<OpenXmlElement, Settings> DocumentSettings = new ConditionalWeakTable<OpenXmlElement, Settings>();

    private RangeWrapper _rangeWrapper;

    public BookmarkWrapperCollection(RangeWrapper rangeWrapper)
    {
        _rangeWrapper = rangeWrapper;
    }

    /// <summary>
    /// The range the collection was taken from.
    /// </summary>
    public RangeWrapper Parent
    {
        get { return _rangeWrapper; }
    }

    private OpenXmlElement Start
    {
        get { return _rangeWrapper.StartElement ?? _rangeWrapper.AnchorElement; }
    }

    private Settings GetSettings()
    {
        return DocumentSettings.GetValue(Start.GetRoot(), key => new Settings());
    }

    /// <summary>True to include the hidden bookmarks, those whose names start with an underscore (<c>Bookmarks.ShowHidden</c> in Word).</summary>
    public bool ShowHidden
    {
        get { return GetSettings().ShowHidden; }
        set { GetSettings().ShowHidden = value; }
    }

    /// <summary>How the bookmarks are ordered: by name (the default, as in Word) or by location in the document.</summary>
    public BookmarkSorting DefaultSorting
    {
        get { return GetSettings().Sorting; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            GetSettings().Sorting = value;
        }
    }

    private static bool IsHidden(BookmarkStart bookmark)
    {
        string name = bookmark.Name?.Value;
        return name != null && name.Length > 0 && name[0] == '_';
    }

    //The bookmarks that the collection holds, in its order.
    private IList<BookmarkStart> GetElementList()
    {
        Settings settings = GetSettings();
        IEnumerable<BookmarkStart> bookmarks = Start.GetElements<BookmarkStart>(_rangeWrapper.EndElement ?? Start);
        if (!settings.ShowHidden)
        {
            bookmarks = bookmarks.Where(b => !IsHidden(b));
        }
        if (settings.Sorting == BookmarkSorting.Name)
        {
            //A stable sort, so bookmarks that differ only in case stay in document order.
            bookmarks = bookmarks.Select((b, i) => (Bookmark: b, Index: i)).OrderBy(p => p.Bookmark.Name?.Value, StringComparer.InvariantCultureIgnoreCase).ThenBy(p => p.Index).Select(p => p.Bookmark);
        }
        return bookmarks.ToList();
    }

    public BookmarkWrapper this[int index]
    {
        get
        {
            IList<BookmarkStart> bookmarks = GetElementList();
            if (index < 0 || index >= bookmarks.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return new BookmarkWrapper(bookmarks[index]);
        }
    }

    public BookmarkWrapper this[string name]
    {
        get
        {
            BookmarkStart bookmarkStart = GetElementList().Where(e => string.Equals(e.Name?.Value, name, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            if (bookmarkStart != null)
            {
                return new BookmarkWrapper(bookmarkStart);
            }
            return null;
        }
    }

    /// <summary>
    /// True if the collection has a bookmark with the name. Names are matched without regard to case, as in Word.
    /// </summary>
    public bool Exists(string name)
    {
        return GetElementList().Any(e => string.Equals(e.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Adds a bookmark over the range, as <c>Bookmarks.Add</c> in Word. A bookmark of the same name anywhere in the document is
    /// moved to the new range.
    /// </summary>
    /// <remarks>
    /// Supported for the ranges of a paragraph, table, table cell, bookmark or comment, and for any span of characters (a collapsed one gives a bookmark at an insertion point). A table is bookmarked by markers before and after it. The name must start with a letter or an underscore (which makes the bookmark hidden), contain only
    /// letters, digits and underscores, and be at most 40 characters long.
    /// </remarks>
    /// <exception cref="ArgumentException">The name is not valid.</exception>
    /// <exception cref="NotSupportedException">The range is of a type that cannot be bookmarked.</exception>
    public BookmarkWrapper Add(string name, RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
        if (string.IsNullOrEmpty(name) || name.Length > 40 || !(char.IsLetter(name[0]) || name[0] == '_') || !name.All(c => char.IsLetterOrDigit(c) || c == '_'))
        {
            throw new ArgumentException("A bookmark name must start with a letter or underscore, contain only letters, digits and underscores, and be at most 40 characters long.", nameof(name));
        }
        if (!range.IsSpan && !(range.StartElement is Paragraph || range.StartElement is TableCell || range.StartElement is Table || range.StartElement is BookmarkStart || range.StartElement is CommentRangeStart))
        {
            throw new NotSupportedException(string.Format(CultureInfo.InvariantCulture, "Adding a bookmark to a range starting with the type {0} is not currently supported.", range.StartElement.GetType().Name));
        }

        if (!range.IsSpan && range.StartElement is BookmarkStart sameBookmark && string.Equals(sameBookmark.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
        {
            //The range already is the bookmark.
            return new BookmarkWrapper(sameBookmark);
        }

        OpenXmlElement root = range.AnchorElement.GetRoot();
        foreach (BookmarkStart existing in root.Descendants<BookmarkStart>().Where(e => string.Equals(e.Name?.Value, name, StringComparison.OrdinalIgnoreCase)).ToList())
        {
            new BookmarkWrapper(existing).Delete();
        }

        int id = 0;
        foreach (BookmarkStart other in root.Descendants<BookmarkStart>())
        {
            id = Math.Max(id, Convert.ToInt32(other.Id.Value, CultureInfo.InvariantCulture) + 1);
        }
        string idText = id.ToString(CultureInfo.InvariantCulture);
        BookmarkStart bookmarkStart = new BookmarkStart() { Id = idText, Name = name };
        BookmarkEnd bookmarkEnd = new BookmarkEnd() { Id = idText };

        if (range.IsSpan)
        {
            range.PlaceMarkers(bookmarkStart, bookmarkEnd);
        }
        else if (range.StartElement is Paragraph paragraph)
        {
            AddToParagraph(paragraph, bookmarkStart, paragraph, bookmarkEnd);
        }
        else if (range.StartElement is TableCell cell)
        {
            AddToParagraph(cell.Elements<Paragraph>().First(), bookmarkStart, cell.Elements<Paragraph>().Last(), bookmarkEnd);
        }
        else
        {
            range.StartElement.InsertBeforeSelf(bookmarkStart);
            range.EndElement.InsertAfterSelf(bookmarkEnd);
        }
        return new BookmarkWrapper(bookmarkStart);
    }

    private static void AddToParagraph(Paragraph first, BookmarkStart bookmarkStart, Paragraph last, BookmarkEnd bookmarkEnd)
    {
        ParagraphProperties properties = first.ParagraphProperties;
        if (properties != null)
        {
            properties.InsertAfterSelf(bookmarkStart);
        }
        else
        {
            first.PrependChild(bookmarkStart);
        }
        last.AppendChild(bookmarkEnd);
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    public void DeleteAll()
    {
        //It is important that the list is fetched since we are deleting items inside the iteration.
        //Each bookmark is removed as a whole (its start and its matching end) so that no dangling start or end markers are left behind
        //when the two markers are not both inside the range.
        IList<BookmarkStart> bookmarkStartList = GetElementList();
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
