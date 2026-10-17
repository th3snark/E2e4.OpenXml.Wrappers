using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class BookmarkWrapperCollection : IEnumerable<BookmarkWrapper>
{
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
            BookmarkStart bookmarkStart = _rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).Where(e => String.Equals(e.Name?.Value, name, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
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
        return _rangeWrapper.Start.GetElements<BookmarkStart>(_rangeWrapper.End).Any(e => String.Equals(e.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Adds a bookmark over the range, as <c>Bookmarks.Add</c> in Word. A bookmark of the same name anywhere in the document is
    /// moved to the new range.
    /// </summary>
    /// <remarks>
    /// Supported for the ranges of a paragraph, table, table cell, bookmark or comment. A table is bookmarked by markers before and after it. The name must start with a letter, contain only
    /// letters, digits and underscores, and be at most 40 characters long.
    /// </remarks>
    /// <exception cref="ArgumentException">The name is not valid.</exception>
    /// <exception cref="NotSupportedException">The range is of a type that cannot be bookmarked.</exception>
    public BookmarkWrapper Add(string name, RangeWrapper range)
    {
        ArgumentNullException.ThrowIfNull(range);
        if (String.IsNullOrEmpty(name) || name.Length > 40 || !Char.IsLetter(name[0]) || !name.All(c => Char.IsLetterOrDigit(c) || c == '_'))
        {
            throw new ArgumentException("A bookmark name must start with a letter, contain only letters, digits and underscores, and be at most 40 characters long.", nameof(name));
        }
        if (!(range.Start is Paragraph || range.Start is TableCell || range.Start is Table || range.Start is BookmarkStart || range.Start is CommentRangeStart))
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Adding a bookmark to a range starting with the type {0} is not currently supported.", range.Start.GetType().Name));
        }

        if (range.Start is BookmarkStart sameBookmark && String.Equals(sameBookmark.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
        {
            //The range already is the bookmark.
            return new BookmarkWrapper(sameBookmark);
        }

        OpenXmlElement root = range.Start.GetRoot();
        foreach (BookmarkStart existing in root.Descendants<BookmarkStart>().Where(e => String.Equals(e.Name?.Value, name, StringComparison.OrdinalIgnoreCase)).ToList())
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

        if (range.Start is Paragraph paragraph)
        {
            AddToParagraph(paragraph, bookmarkStart, paragraph, bookmarkEnd);
        }
        else if (range.Start is TableCell cell)
        {
            AddToParagraph(cell.Elements<Paragraph>().First(), bookmarkStart, cell.Elements<Paragraph>().Last(), bookmarkEnd);
        }
        else
        {
            range.Start.InsertBeforeSelf(bookmarkStart);
            range.End.InsertAfterSelf(bookmarkEnd);
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
