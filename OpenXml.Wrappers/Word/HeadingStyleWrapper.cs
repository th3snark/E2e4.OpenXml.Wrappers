using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A style that a table of contents makes entries from beyond the heading styles, with the level of the entries, as <c>HeadingStyle</c> in
/// Word: one pair of the <c>\t</c> switch of the <c>TOC</c> field.
/// </summary>
public class HeadingStyleWrapper : IEquatable<HeadingStyleWrapper>
{
    private readonly TableOfContentsWrapper _toc;
    private readonly string _style;

    internal HeadingStyleWrapper(TableOfContentsWrapper toc, string style)
    {
        _toc = toc;
        _style = style;
    }

    private int IndexOf()
    {
        List<KeyValuePair<string, int>> styles = _toc.AddedStyles();
        int index = styles.FindIndex(s => string.Equals(s.Key, _style, StringComparison.OrdinalIgnoreCase));
        return index >= 0 ? index : throw new InvalidOperationException("The heading style has been deleted.");
    }

    /// <summary>The name of the style, as <c>HeadingStyle.Style</c> in Word.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public string Style
    {
        get { return _toc.AddedStyles()[IndexOf()].Key; }
        set
        {
            Guard.ThrowIfNull(value);
            int index = IndexOf();
            _toc.ChangeAddedStyles(styles => styles[index] = new KeyValuePair<string, int>(value, styles[index].Value));
        }
    }

    /// <summary>The level of the entries made from the style (1 to 9), as <c>HeadingStyle.Level</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The level is not from 1 to 9.</exception>
    public int Level
    {
        get { return _toc.AddedStyles()[IndexOf()].Value; }
        set
        {
            if (value < 1 || value > 9)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            int index = IndexOf();
            _toc.ChangeAddedStyles(styles => styles[index] = new KeyValuePair<string, int>(styles[index].Key, value));
        }
    }

    /// <summary>Takes the style out of the table, as <c>HeadingStyle.Delete</c> in Word.</summary>
    public void Delete()
    {
        int index = IndexOf();
        _toc.ChangeAddedStyles(styles => styles.RemoveAt(index));
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as HeadingStyleWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(HeadingStyleWrapper other)
    {
        return !(other is null) && _toc == other._toc && string.Equals(_style, other._style, StringComparison.OrdinalIgnoreCase);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return StringComparer.OrdinalIgnoreCase.GetHashCode(_style);
    }

    public static bool operator ==(HeadingStyleWrapper left, HeadingStyleWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(HeadingStyleWrapper left, HeadingStyleWrapper right)
    {
        return !(left == right);
    }
}

/// <summary>
/// The styles a table of contents makes entries from in addition to the heading styles, as <c>HeadingStyles</c> in Word.
/// </summary>
public class HeadingStyleWrapperCollection : IEnumerable<HeadingStyleWrapper>
{
    private readonly TableOfContentsWrapper _toc;

    internal HeadingStyleWrapperCollection(TableOfContentsWrapper toc)
    {
        _toc = toc;
    }

    /// <summary>The number of styles.</summary>
    public int Count
    {
        get { return _toc.AddedStyles().Count; }
    }

    /// <summary>The style at the index (0-based), as <c>HeadingStyles.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no style at the index.</exception>
    public HeadingStyleWrapper this[int index]
    {
        get
        {
            List<KeyValuePair<string, int>> styles = _toc.AddedStyles();
            if (index < 0 || index >= styles.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new HeadingStyleWrapper(_toc, styles[index].Key);
        }
    }

    /// <summary>Adds a style to make entries from, at a level, as <c>HeadingStyles.Add</c> in Word.</summary>
    /// <param name="style">The name of the style.</param>
    /// <param name="level">The level of the entries (1 to 9).</param>
    /// <exception cref="ArgumentNullException">The style is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The level is not from 1 to 9.</exception>
    public HeadingStyleWrapper Add(string style, int level)
    {
        Guard.ThrowIfNull(style);
        if (level < 1 || level > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(level));
        }
        _toc.ChangeAddedStyles(styles => styles.Add(new KeyValuePair<string, int>(style, level)));
        return new HeadingStyleWrapper(_toc, style);
    }

    /// <inheritdoc/>
    public IEnumerator<HeadingStyleWrapper> GetEnumerator()
    {
        return _toc.AddedStyles().Select(s => new HeadingStyleWrapper(_toc, s.Key)).ToList().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
