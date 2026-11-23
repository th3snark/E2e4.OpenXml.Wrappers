using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>What separates the groups of entries of an index, as <c>WdHeadingSeparator</c> in Word (with the same values).</summary>
public enum HeadingSeparator
{
    /// <summary>Nothing.</summary>
    None = 0,

    /// <summary>A blank line.</summary>
    BlankLine = 1,

    /// <summary>The letter of the group (<c>\h "A"</c>).</summary>
    Letter = 2,

    /// <summary>The letter of the group; Word writes it as <see cref="Letter"/>.</summary>
    LetterLow = 3,

    /// <summary>The letter of the group; Word writes it as <see cref="Letter"/>.</summary>
    LetterFull = 4
}

/// <summary>How the subentries of an index are laid out, as <c>WdIndexType</c> in Word (with the same values).</summary>
public enum IndexType
{
    /// <summary>Each on a line of its own, indented.</summary>
    Indent = 0,

    /// <summary>On the line of the main entry (<c>\r</c>).</summary>
    RunIn = 1
}

/// <summary>
/// An index, as <c>Index</c> in Word: an <c>INDEX</c> field whose result is a paragraph for each entry of the document's <c>XE</c> fields.
/// </summary>
/// <remarks>
/// The entries are sorted, grouped and written as Word writes them. Their page numbers need the pages laid out, so <see cref="Update()"/>
/// estimates them from the breaks in the text and marks the field for Word to update when the document is opened; a caller that knows the
/// pages gives them to <see cref="Update(Func{RangeWrapper, string})"/>.
/// </remarks>
public class IndexWrapper : IEquatable<IndexWrapper>
{
    private FieldWrapper _field;

    internal IndexWrapper(FieldWrapper field)
    {
        _field = field;
    }

    /// <summary>The index, from the start of its field to its end, as <c>Index.Range</c> in Word.</summary>
    public RangeWrapper Range
    {
        get { return _field.Range; }
    }

    /// <summary>The range the index is in, as <c>Index.Parent</c> in Word.</summary>
    public RangeWrapper Parent
    {
        get { return _field.Parent; }
    }

    /// <summary>What separates the groups of entries, as <c>Index.HeadingSeparator</c> in Word (the <c>\h</c> switch).</summary>
    public HeadingSeparator HeadingSeparator
    {
        get
        {
            string heading = IndexCode.Get(_field.CodeText, 'h');
            return heading == null || heading.Length == 0 ? HeadingSeparator.None
                : heading.Trim().Length == 0 ? HeadingSeparator.BlankLine : HeadingSeparator.Letter;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Change('h', value == HeadingSeparator.None ? null : value == HeadingSeparator.BlankLine ? " " : "A", value != HeadingSeparator.None);
        }
    }

    /// <summary>True if the page numbers are at the right margin, as <c>Index.RightAlignPageNumbers</c> in Word (<c>\e</c> with a tab).</summary>
    public bool RightAlignPageNumbers
    {
        get { return IndexCode.Get(_field.CodeText, 'e') == "\t"; }
        set
        {
            if (value != RightAlignPageNumbers)
            {
                Change('e', value ? "\t" : null, value, true);
            }
        }
    }

    /// <summary>How the subentries are laid out, as <c>Index.Type</c> in Word (<c>\r</c> for run-in).</summary>
    public IndexType Type
    {
        get { return IndexCode.Has(_field.CodeText, 'r') ? IndexType.RunIn : IndexType.Indent; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            Change('r', null, value == IndexType.RunIn, true);
        }
    }

    /// <summary>The number of columns, as <c>Index.NumberOfColumns</c> in Word (<c>\c</c>); 0 when the index is not set in columns.</summary>
    public int NumberOfColumns
    {
        get
        {
            return int.TryParse(IndexCode.Get(_field.CodeText, 'c'), NumberStyles.Integer, CultureInfo.InvariantCulture, out int columns) ? columns : 0;
        }
        set
        {
            if (value < 0 || value > 4)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "An index has 0 to 4 columns.");
            }
            Change('c', value.ToString(CultureInfo.InvariantCulture), value > 0);
        }
    }

    /// <summary>True if letters with accents have groups of their own, as <c>Index.AccentedLetters</c> in Word (<c>\a</c>).</summary>
    public bool AccentedLetters
    {
        get { return IndexCode.Has(_field.CodeText, 'a'); }
        set { Change('a', null, value); }
    }

    /// <summary>The language the index is sorted for, as <c>Index.IndexLanguage</c> in Word (<c>\z</c>, a language id such as 2057); 0 when
    /// none is given.</summary>
    public int IndexLanguage
    {
        get { return int.TryParse(IndexCode.Get(_field.CodeText, 'z'), NumberStyles.Integer, CultureInfo.InvariantCulture, out int language) ? language : 0; }
        set { Change('z', value.ToString(CultureInfo.InvariantCulture), value > 0); }
    }

    /// <summary>The leader before right-aligned page numbers, as <c>Index.TabLeader</c> in Word: the tab stop of the index's paragraphs.</summary>
    public TabLeader TabLeader
    {
        get
        {
            TabStop tab = Paragraphs().Select(p => p.ParagraphProperties?.Tabs?.Elements<TabStop>().FirstOrDefault()).FirstOrDefault(t => t != null);
            return tab?.Leader?.Value == null ? TabLeader.Dots : TableOfContentsWrapper.FromLeader(tab.Leader.Value);
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            foreach (TabStop tab in Paragraphs().SelectMany(p => p.ParagraphProperties?.Tabs?.Elements<TabStop>() ?? Enumerable.Empty<TabStop>()))
            {
                tab.Leader = TableOfContentsWrapper.ToLeader(value);
            }
        }
    }

    //The paragraphs of the index, from the one its field starts in to the one before its end.
    private IEnumerable<Paragraph> Paragraphs()
    {
        Paragraph first = _field.Key.Ancestors<Paragraph>().FirstOrDefault();
        FieldChar end = _field.EndChar();
        Paragraph last = end?.Ancestors<Paragraph>().FirstOrDefault();
        for (Paragraph p = first; p != null && !ReferenceEquals(p, last); p = p.NextSibling<Paragraph>())
        {
            yield return p;
        }
    }

    private void Change(char name, string value, bool present, bool atStart = false)
    {
        _field.CodeText = IndexCode.Set(_field.CodeText, name, value, present, atStart);
    }

    /// <summary>
    /// Takes the index away, as <c>Index.Delete</c> in Word: its field and entries go, and the text around it stays.
    /// </summary>
    public void Delete()
    {
        _field.EnsureComplex();
        _field.ClearResult(out Paragraph begin, out Paragraph end, out Run _);
        if (!ReferenceEquals(begin, end) && begin.ParagraphProperties?.SectionProperties == null
            && !begin.Elements<Run>().Any(r => r.Elements<Text>().Any()))
        {
            //The paragraph the field started in held only the field.
            _field.JoinResultParagraphs();
        }
        _field.Delete();
    }

    /// <summary>
    /// Builds the entries again from the <c>XE</c> fields, as <c>Index.Update</c> in Word. The page numbers are estimated from the breaks in the
    /// text (page breaks, section breaks and where Word last broke the pages, which a document saved by Word records), in each section's
    /// numbering, and the field is marked for Word to update when the document is opened.
    /// </summary>
    public void Update()
    {
        IndexBuilder.Update(_field, null);
    }

    /// <summary>
    /// Builds the entries again from the <c>XE</c> fields, with page numbers from <paramref name="pageNumber"/>, for a caller that has laid the
    /// pages out. It is called with the range of each <c>XE</c> field, and with the start and the end (collapsed ranges) of the bookmark of an
    /// <c>XE</c> field with <c>\r</c>, and returns the page number shown there as a <c>PAGE</c> field there would show it, or null to leave
    /// that page out. The chapter numbers of <c>\s</c> and the separators are added here.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="pageNumber"/> is null.</exception>
    public void Update(Func<RangeWrapper, string> pageNumber)
    {
        Guard.ThrowIfNull(pageNumber);
        IndexBuilder.Update(_field, pageNumber);
    }

    #region Equality

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as IndexWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(IndexWrapper other)
    {
        return !(other is null) && _field == other._field;
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _field.GetHashCode();
    }

    public static bool operator ==(IndexWrapper left, IndexWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(IndexWrapper left, IndexWrapper right)
    {
        return !(left == right);
    }

    #endregion
}

/// <summary>
/// Reads and changes the switches of an <c>INDEX</c> field code, keeping the others as they are.
/// </summary>
internal static class IndexCode
{
    public static bool Has(string code, char name)
    {
        return IndexSettings.Parse(code).Has(name);
    }

    public static string Get(string code, char name)
    {
        return IndexSettings.Parse(code).Get(name);
    }

    //Puts the switch in (with its argument), replaces it, or takes it out. A new switch goes at the start when atStart (as Word puts \r and
    //\e), otherwise before \z, as Word does.
    public static string Set(string code, char name, string value, bool present, bool atStart)
    {
        List<KeyValuePair<bool, string>> tokens = FieldCodeTokens.Split(code, false);
        List<List<string>> switches = new List<List<string>>();
        for (int i = 1; i < tokens.Count; i++)
        {
            if (tokens[i].Key)
            {
                switches.Add(new List<string> { "\\" + tokens[i].Value });
            }
            else if (switches.Count > 0)
            {
                switches[switches.Count - 1].Add(Quote(tokens[i].Value));
            }
        }

        int at = switches.FindIndex(s => string.Equals(s[0], "\\" + name, StringComparison.OrdinalIgnoreCase));
        List<string> wanted = present ? new List<string> { "\\" + name } : null;
        if (wanted != null && value != null)
        {
            wanted.Add(Quote(value));
        }
        if (at >= 0)
        {
            if (wanted == null)
            {
                switches.RemoveAt(at);
            }
            else
            {
                switches[at] = wanted;
            }
        }
        else if (wanted != null)
        {
            int z = switches.FindIndex(s => string.Equals(s[0], "\\z", StringComparison.OrdinalIgnoreCase));
            switches.Insert(atStart ? 0 : z >= 0 ? z : switches.Count, wanted);
        }

        StringBuilder text = new StringBuilder("INDEX");
        foreach (List<string> item in switches)
        {
            text.Append(' ').Append(string.Join(" ", item));
        }
        return text.ToString();
    }

    private static string Quote(string value)
    {
        return "\"" + value + "\"";
    }
}

/// <summary>
/// The indexes of a range or a document, as <c>Indexes</c> in Word.
/// </summary>
public class IndexWrapperCollection : IEnumerable<IndexWrapper>
{
    private readonly RangeWrapper _range;

    internal IndexWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the indexes are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<IndexWrapper> GetItems()
    {
        return _range.Fields.Where(f => f.Type == FieldType.Index).Select(f => new IndexWrapper(f)).ToList();
    }

    /// <summary>The number of indexes.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The index at the index (0-based), as <c>Indexes.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no index there.</exception>
    public IndexWrapper this[int index]
    {
        get
        {
            List<IndexWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Makes an index of the document's <c>XE</c> fields, as <c>Indexes.Add</c> in Word, in place of the range. Its page numbers are estimated
    /// as <see cref="IndexWrapper.Update()"/> does.
    /// </summary>
    /// <param name="range">Where the index goes; text in it is replaced.</param>
    /// <param name="headingSeparator">What separates the groups of entries.</param>
    /// <param name="rightAlignPageNumbers">True to put the page numbers at the right margin.</param>
    /// <param name="type">How the subentries are laid out.</param>
    /// <param name="numberOfColumns">The number of columns (0 to 4); with any, the index is a section of its own, as in Word.</param>
    /// <param name="accentedLetters">True to give letters with accents groups of their own.</param>
    /// <param name="indexLanguage">The language to sort for (a language id); 0 for that of the current culture.</param>
    public IndexWrapper Add(RangeWrapper range, HeadingSeparator headingSeparator = HeadingSeparator.Letter, bool rightAlignPageNumbers = false,
        IndexType type = IndexType.Indent, int numberOfColumns = 0, bool accentedLetters = false, int indexLanguage = 0)
    {
        Guard.ThrowIfNull(range);
        if (!EnumCompat.IsDefined(headingSeparator))
        {
            throw new ArgumentOutOfRangeException(nameof(headingSeparator));
        }
        if (!EnumCompat.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }
        if (numberOfColumns < 0 || numberOfColumns > 4)
        {
            throw new ArgumentOutOfRangeException(nameof(numberOfColumns), "An index has 0 to 4 columns.");
        }

        //The switches in the order Word writes them.
        StringBuilder code = new StringBuilder("INDEX");
        if (type == IndexType.RunIn)
        {
            code.Append(" \\r");
        }
        if (rightAlignPageNumbers)
        {
            code.Append(" \\e \"\t\"");
        }
        if (numberOfColumns > 0)
        {
            code.Append(" \\c \"").Append(numberOfColumns.ToString(CultureInfo.InvariantCulture)).Append('"');
        }
        if (accentedLetters)
        {
            code.Append(" \\a");
        }
        if (headingSeparator != HeadingSeparator.None)
        {
            code.Append(headingSeparator == HeadingSeparator.BlankLine ? " \\h \" \"" : " \\h \"A\"");
        }
        int language = indexLanguage > 0 ? indexLanguage : CultureInfo.CurrentCulture.LCID;
        code.Append(" \\z \"").Append(language.ToString(CultureInfo.InvariantCulture)).Append("\" ");

        if (range.Start != range.End)
        {
            range.Text = string.Empty;
        }
        FieldWrapper field = FieldWrapperCollection.AddCode(range, " " + code.ToString() + " ");
        return new IndexWrapper(field);
    }

    /// <summary>
    /// Marks an entry for the index with an <c>XE</c> field at the end of the range, as <c>Indexes.MarkEntry</c> in Word, and returns the field.
    /// </summary>
    /// <param name="range">The text the entry is for; the field goes after it.</param>
    /// <param name="entry">The entry, with a colon before a subentry ("Fruit:Apple"); the text of the range when null.</param>
    /// <param name="entryAutoText">Not supported: there are no AutoText entries here.</param>
    /// <param name="crossReference">A cross-reference in place of the page number, such as "See Fruit" (<c>\t</c>).</param>
    /// <param name="crossReferenceAutoText">Not supported: there are no AutoText entries here.</param>
    /// <param name="bookmarkName">A bookmark whose pages the entry gives as a range (<c>\r</c>).</param>
    /// <param name="bold">True for a bold page number (<c>\b</c>).</param>
    /// <param name="italic">True for an italic page number (<c>\i</c>).</param>
    /// <param name="reading">The reading the entry is sorted by in East Asian languages (<c>\y</c>).</param>
    /// <exception cref="NotSupportedException">An AutoText entry is given.</exception>
    public FieldWrapper MarkEntry(RangeWrapper range, string entry = null, string entryAutoText = null, string crossReference = null,
        string crossReferenceAutoText = null, string bookmarkName = null, bool bold = false, bool italic = false, string reading = null)
    {
        Guard.ThrowIfNull(range);
        if (entryAutoText != null || crossReferenceAutoText != null)
        {
            throw new NotSupportedException("AutoText entries are not supported.");
        }

        StringBuilder code = new StringBuilder(" XE \"").Append(Escape(entry ?? range.Text.TrimEnd('\r'))).Append('"');
        if (!string.IsNullOrEmpty(reading))
        {
            code.Append(" \\y \"").Append(Escape(reading)).Append('"');
        }
        if (!string.IsNullOrEmpty(crossReference))
        {
            code.Append(" \\t \"").Append(Escape(crossReference)).Append('"');
        }
        if (!string.IsNullOrEmpty(bookmarkName))
        {
            code.Append(" \\r ").Append(bookmarkName);
        }
        if (bold)
        {
            code.Append(" \\b");
        }
        if (italic)
        {
            code.Append(" \\i");
        }
        code.Append(' ');

        //An XE field has no result: Word writes its begin, code and end.
        RangeWrapper at = range.Duplicate();
        at.Collapse(CollapseDirection.End);
        FieldChar begin = new FieldChar() { FieldCharType = FieldCharValues.Begin };
        at.InsertFieldRuns(_ => new List<Run>
        {
            new Run(begin),
            new Run(new FieldCode(code.ToString()) { Space = SpaceProcessingModeValues.Preserve }),
            new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
        });
        return new FieldWrapper(begin);
    }

    private static string Escape(string text)
    {
        return text.Replace("\"", "\\\"");
    }

    /// <inheritdoc/>
    public IEnumerator<IndexWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
