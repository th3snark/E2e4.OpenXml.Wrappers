using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// An entry of the list of a drop-down list or combo box content control, as <c>ContentControlListEntry</c> in Word.
/// </summary>
public class ContentControlListEntry : IEquatable<ContentControlListEntry>
{
    private readonly ListItem _item;

    internal ContentControlListEntry(ListItem item)
    {
        _item = item;
    }

    private OpenXmlElement List
    {
        get { return _item.Parent; }
    }

    /// <summary>The position of the entry in the list (0-based).</summary>
    public int Index
    {
        get { return List.Elements<ListItem>().ToList().IndexOf(_item); }
    }

    /// <summary>The text shown for the entry, as <c>ContentControlListEntry.Text</c> in Word.</summary>
    public string Text
    {
        get { return _item.DisplayText?.Value ?? _item.Value?.Value ?? string.Empty; }
        set
        {
            Guard.ThrowIfNull(value);
            _item.DisplayText = value;
        }
    }

    /// <summary>The value stored for the entry, as <c>ContentControlListEntry.Value</c> in Word.</summary>
    public string Value
    {
        get { return _item.Value?.Value ?? string.Empty; }
        set
        {
            Guard.ThrowIfNull(value);
            _item.Value = value;
        }
    }

    /// <summary>Takes the entry out of the list, as <c>ContentControlListEntry.Delete</c> in Word.</summary>
    public void Delete()
    {
        _item.Remove();
    }

    /// <summary>Moves the entry one place towards the start of the list, as <c>ContentControlListEntry.MoveUp</c> in Word.</summary>
    public void MoveUp()
    {
        if (_item.PreviousSibling<ListItem>() is ListItem before)
        {
            _item.Remove();
            before.InsertBeforeSelf(_item);
        }
    }

    /// <summary>Moves the entry one place towards the end of the list, as <c>ContentControlListEntry.MoveDown</c> in Word.</summary>
    public void MoveDown()
    {
        if (_item.NextSibling<ListItem>() is ListItem after)
        {
            _item.Remove();
            after.InsertAfterSelf(_item);
        }
    }

    /// <summary>
    /// Makes the entry the one the control shows, as <c>ContentControlListEntry.Select</c> in Word: the control then holds the text of the
    /// entry and no longer shows its placeholder.
    /// </summary>
    public void Select()
    {
        if (!(List.Parent is SdtProperties properties) || !(properties.Parent is SdtElement sdt))
        {
            return;
        }
        new ContentControlWrapper(sdt).ReplaceContents(Text);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as ContentControlListEntry);
    }

    /// <inheritdoc/>
    public bool Equals(ContentControlListEntry other)
    {
        return !(other is null) && ReferenceEquals(_item, other._item);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _item.GetHashCode();
    }

    public static bool operator ==(ContentControlListEntry left, ContentControlListEntry right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(ContentControlListEntry left, ContentControlListEntry right)
    {
        return !(left == right);
    }
}

/// <summary>
/// The entries of the list of a drop-down list or combo box content control, as <c>ContentControlListEntries</c> in Word.
/// </summary>
public class ContentControlListEntries : IEnumerable<ContentControlListEntry>
{
    private readonly OpenXmlCompositeElement _list;

    internal ContentControlListEntries(OpenXmlCompositeElement list)
    {
        _list = list;
    }

    private List<ListItem> Items()
    {
        return _list.Elements<ListItem>().ToList();
    }

    /// <summary>The number of entries.</summary>
    public int Count
    {
        get { return Items().Count; }
    }

    /// <summary>The entry at the index (0-based), as <c>ContentControlListEntries.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no entry at the index.</exception>
    public ContentControlListEntry this[int index]
    {
        get
        {
            List<ListItem> items = Items();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new ContentControlListEntry(items[index]);
        }
    }

    /// <summary>
    /// Adds an entry to the list, as <c>ContentControlListEntries.Add</c> in Word.
    /// </summary>
    /// <param name="text">The text shown for the entry.</param>
    /// <param name="value">The value stored for it; the text when null.</param>
    /// <param name="index">Where in the list to put it (0-based); at the end when null.</param>
    public ContentControlListEntry Add(string text, string value = null, int? index = null)
    {
        Guard.ThrowIfNull(text);
        List<ListItem> items = Items();
        if (index.HasValue && (index.Value < 0 || index.Value > items.Count))
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }
        ListItem item = new ListItem() { DisplayText = text, Value = value ?? text };
        if (!index.HasValue || index.Value == items.Count)
        {
            _list.AppendChild(item);
        }
        else
        {
            items[index.Value].InsertBeforeSelf(item);
        }
        return new ContentControlListEntry(item);
    }

    /// <summary>Takes every entry out of the list, as <c>ContentControlListEntries.Clear</c> in Word.</summary>
    public void Clear()
    {
        foreach (ListItem item in Items())
        {
            item.Remove();
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ContentControlListEntry> GetEnumerator()
    {
        return Items().Select(i => new ContentControlListEntry(i)).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
