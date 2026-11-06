using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A table of figures, as <c>TableOfFigures</c> in Word: a <c>TOC</c> field with the <c>\c</c> (or <c>\a</c>) switch whose result is a
/// paragraph for each caption with a label. It is made and kept like a table of contents (see <see cref="TableOfContentsWrapper"/>).
/// </summary>
public class TableOfFiguresWrapper : IEquatable<TableOfFiguresWrapper>
{
    private readonly TableOfContentsWrapper _toc;

    internal TableOfFiguresWrapper(TableOfContentsWrapper toc)
    {
        _toc = toc;
    }

    /// <summary>The label of the captions in the table, as <c>TableOfFigures.Caption</c> in Word.</summary>
    /// <exception cref="ArgumentNullException">The value is null.</exception>
    public string Caption
    {
        get { return _toc.CaptionLabel; }
        set
        {
            Guard.ThrowIfNull(value);
            _toc.CaptionLabel = value;
        }
    }

    /// <summary>True if the entries keep the label and number of the caption, as <c>TableOfFigures.IncludeLabel</c> in Word.</summary>
    public bool IncludeLabel
    {
        get { return _toc.CaptionIncludesLabel; }
        set { _toc.CaptionIncludesLabel = value; }
    }

    /// <summary>The table, from the start of its field to its end, as <c>TableOfFigures.Range</c> in Word.</summary>
    public RangeWrapper Range
    {
        get { return _toc.Range; }
    }

    /// <summary>The range the table is in, as <c>TableOfFigures.Parent</c> in Word.</summary>
    public RangeWrapper Parent
    {
        get { return _toc.Parent; }
    }

    /// <summary>True if the table has page numbers, as <c>TableOfFigures.IncludePageNumbers</c> in Word.</summary>
    public bool IncludePageNumbers
    {
        get { return _toc.IncludePageNumbers; }
        set { _toc.IncludePageNumbers = value; }
    }

    /// <summary>True if the page numbers are at the right margin, as <c>TableOfFigures.RightAlignPageNumbers</c> in Word.</summary>
    public bool RightAlignPageNumbers
    {
        get { return _toc.RightAlignPageNumbers; }
        set { _toc.RightAlignPageNumbers = value; }
    }

    /// <summary>True if the entries are hyperlinks to the captions, as <c>TableOfFigures.UseHyperlinks</c> in Word.</summary>
    public bool UseHyperlinks
    {
        get { return _toc.UseHyperlinks; }
        set { _toc.UseHyperlinks = value; }
    }

    /// <summary>True if a web page leaves the page numbers out, as <c>TableOfFigures.HidePageNumbersInWeb</c> in Word.</summary>
    public bool HidePageNumbersInWeb
    {
        get { return _toc.HidePageNumbersInWeb; }
        set { _toc.HidePageNumbersInWeb = value; }
    }

    /// <summary>The leader between an entry and its page number, as <c>TableOfFigures.TabLeader</c> in Word.</summary>
    public TabLeader TabLeader
    {
        get { return _toc.TabLeader; }
        set { _toc.TabLeader = value; }
    }

    /// <summary>The identifier of the kind of TC fields the table uses, as <c>TableOfFigures.TableID</c> in Word.</summary>
    public string TableID
    {
        get { return _toc.TableID; }
        set { _toc.TableID = value; }
    }

    /// <summary>True if the table uses TC fields, as <c>TableOfFigures.UseFields</c> in Word.</summary>
    public bool UseFields
    {
        get { return _toc.UseFields; }
        set { _toc.UseFields = value; }
    }

    /// <summary>Takes the table away, as <c>TableOfFigures.Delete</c> in Word.</summary>
    public void Delete()
    {
        _toc.Delete();
    }

    /// <summary>Makes the entries again from the captions, as <c>TableOfFigures.Update</c> in Word.</summary>
    public void Update()
    {
        _toc.Update();
    }

    #region Equality

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as TableOfFiguresWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(TableOfFiguresWrapper other)
    {
        return !(other is null) && _toc == other._toc;
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _toc.GetHashCode();
    }

    public static bool operator ==(TableOfFiguresWrapper left, TableOfFiguresWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(TableOfFiguresWrapper left, TableOfFiguresWrapper right)
    {
        return !(left == right);
    }

    #endregion
}

/// <summary>
/// The tables of figures of a range or a document, as <c>TablesOfFigures</c> in Word.
/// </summary>
public class TableOfFiguresWrapperCollection : IEnumerable<TableOfFiguresWrapper>
{
    private readonly RangeWrapper _range;

    internal TableOfFiguresWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the tables are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<TableOfFiguresWrapper> GetItems()
    {
        return _range.Fields
            .Where(f => f.Type == FieldType.TableOfContents && Regex.IsMatch(f.CodeText, "\\\\[ac]\\b"))
            .Select(f => new TableOfFiguresWrapper(new TableOfContentsWrapper(f))).ToList();
    }

    /// <summary>The number of tables of figures.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The table at the index (0-based), as <c>TablesOfFigures.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no table at the index.</exception>
    public TableOfFiguresWrapper this[int index]
    {
        get
        {
            List<TableOfFiguresWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Makes a table of figures, as <c>TablesOfFigures.Add</c> in Word. Its entries go in front of the paragraph the range is at the start of.
    /// </summary>
    /// <param name="range">Where to put the table; it must be at the start of a paragraph.</param>
    /// <param name="caption">The label of the captions to list, for example <c>Figure</c>.</param>
    /// <param name="includeLabel">True to keep the label and number in the entries.</param>
    /// <param name="useHeadingStyles">True to make entries from the heading styles too.</param>
    /// <param name="upperHeadingLevel">The first heading level (1 to 9).</param>
    /// <param name="lowerHeadingLevel">The last heading level (1 to 9).</param>
    /// <param name="useFields">True if the table uses TC fields.</param>
    /// <param name="tableID">The identifier of the kind of TC fields.</param>
    /// <param name="rightAlignPageNumbers">True to put the page numbers at the right margin.</param>
    /// <param name="includePageNumbers">True to give the entries page numbers.</param>
    /// <param name="addedStyles">Other styles to make entries from, as <see cref="TableOfContentsWrapperCollection.Add"/> takes them.</param>
    /// <param name="useHyperlinks">True to make the entries hyperlinks to the captions.</param>
    /// <param name="hidePageNumbersInWeb">True to leave the page numbers out of a web page.</param>
    /// <exception cref="ArgumentNullException">The caption is null.</exception>
    /// <exception cref="NotSupportedException">The range is not at the start of a paragraph.</exception>
    public TableOfFiguresWrapper Add(RangeWrapper range, string caption = "Figure", bool includeLabel = true, bool useHeadingStyles = false,
        int upperHeadingLevel = 1, int lowerHeadingLevel = 9, bool useFields = false, string tableID = null, bool rightAlignPageNumbers = true,
        bool includePageNumbers = true, string addedStyles = null, bool useHyperlinks = false, bool hidePageNumbersInWeb = false)
    {
        Guard.ThrowIfNull(range);
        Guard.ThrowIfNull(caption);
        TableOfContentsWrapper toc = new TableOfContentsWrapperCollection(_range).AddCaptionTable(range, caption, includeLabel, useHeadingStyles, upperHeadingLevel,
            lowerHeadingLevel, useFields, tableID, rightAlignPageNumbers, includePageNumbers, addedStyles, useHyperlinks, hidePageNumbersInWeb);
        return new TableOfFiguresWrapper(toc);
    }

    /// <inheritdoc/>
    public IEnumerator<TableOfFiguresWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
