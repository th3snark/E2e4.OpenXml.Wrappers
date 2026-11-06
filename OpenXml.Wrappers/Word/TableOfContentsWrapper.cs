using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// How a table of contents is made, read from and written to its <c>TOC</c> field code.
/// </summary>
internal sealed class TocSettings
{
    internal bool UseHeadingStyles = true;
    internal int Upper = 1;
    internal int Lower = 9;
    internal bool UseFields;
    internal string TableId;
    internal bool RightAlign = true;
    internal bool IncludePageNumbers = true;
    internal bool UseHyperlinks;
    internal bool HidePageNumbersInWeb;
    internal bool UseOutlineLevels;
    internal List<KeyValuePair<string, int>> AddedStyles = new List<KeyValuePair<string, int>>();

    private static string Separator
    {
        get { return CultureInfo.CurrentCulture.TextInfo.ListSeparator; }
    }

    //The styles of the \t switch: pairs of a style name and a level, separated by the list separator.
    internal static List<KeyValuePair<string, int>> ParseStyles(string text)
    {
        List<KeyValuePair<string, int>> styles = new List<KeyValuePair<string, int>>();
        if (string.IsNullOrEmpty(text))
        {
            return styles;
        }
        string[] parts = text.Split(new[] { Separator }, StringSplitOptions.None);
        if (parts.Length % 2 != 0)
        {
            throw new ArgumentException("The styles must be pairs of a style name and a level.", nameof(text));
        }
        for (int i = 0; i < parts.Length; i += 2)
        {
            if (!int.TryParse(parts[i + 1].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int level) || level < 1 || level > 9)
            {
                throw new ArgumentException("The level of a style must be from 1 to 9.", nameof(text));
            }
            styles.Add(new KeyValuePair<string, int>(parts[i].Trim(), level));
        }
        return styles;
    }

    internal string ToCode()
    {
        StringBuilder code = new StringBuilder(" TOC");
        if (UseHeadingStyles)
        {
            code.Append(string.Format(CultureInfo.InvariantCulture, " \\o \"{0}-{1}\"", Upper, Lower));
        }
        if (UseFields)
        {
            code.Append(" \\f");
            if (!string.IsNullOrEmpty(TableId))
            {
                code.Append(' ').Append(TableId);
            }
        }
        if (!IncludePageNumbers)
        {
            code.Append(" \\n");
        }
        if (!RightAlign)
        {
            code.Append(" \\p \" \"");
        }
        if (UseHyperlinks)
        {
            code.Append(" \\h");
        }
        if (HidePageNumbersInWeb)
        {
            code.Append(" \\z");
        }
        if (UseOutlineLevels)
        {
            code.Append(" \\u");
        }
        if (AddedStyles.Count > 0)
        {
            code.Append(" \\t \"").Append(string.Join(Separator, AddedStyles.Select(s => s.Key + Separator + s.Value.ToString(CultureInfo.InvariantCulture)))).Append('"');
        }
        return code.Append(' ').ToString();
    }

    internal static TocSettings Parse(string code)
    {
        TocSettings settings = new TocSettings();
        Match levels = Regex.Match(code, "\\\\o\\s+\"(?<a>\\d+)-(?<b>\\d+)\"");
        settings.UseHeadingStyles = levels.Success;
        if (levels.Success)
        {
            settings.Upper = int.Parse(levels.Groups["a"].Value, CultureInfo.InvariantCulture);
            settings.Lower = int.Parse(levels.Groups["b"].Value, CultureInfo.InvariantCulture);
        }
        Match fields = Regex.Match(code, "\\\\f(?:\\s+(?<id>[^\\s\\\\\"]+))?");
        settings.UseFields = fields.Success;
        settings.TableId = fields.Success && fields.Groups["id"].Success ? fields.Groups["id"].Value : null;
        settings.UseHyperlinks = Regex.IsMatch(code, "\\\\h\\b");
        settings.IncludePageNumbers = !Regex.IsMatch(code, "\\\\n\\b");
        settings.RightAlign = !Regex.IsMatch(code, "\\\\p\\b");
        settings.HidePageNumbersInWeb = Regex.IsMatch(code, "\\\\z\\b");
        settings.UseOutlineLevels = Regex.IsMatch(code, "\\\\u\\b");
        Match added = Regex.Match(code, "\\\\t\\s+\"(?<v>[^\"]*)\"");
        if (added.Success)
        {
            settings.AddedStyles = ParseStyles(added.Groups["v"].Value);
        }
        return settings;
    }
}

/// <summary>
/// A table of contents, as <c>TableOfContents</c> in Word: a <c>TOC</c> field whose result is a paragraph for each heading.
/// </summary>
/// <remarks>
/// Page numbers need layout, so they are not worked out: the table is written with the page numbers it had (1 for a new entry) and the
/// field is marked for Word to update when the document is opened.
/// </remarks>
public class TableOfContentsWrapper : IEquatable<TableOfContentsWrapper>
{
    private FieldWrapper _field;

    internal TableOfContentsWrapper(FieldWrapper field)
    {
        _field = field;
    }

    internal FieldWrapper Field
    {
        get { return _field; }
    }

    private TocSettings Settings
    {
        get { return TocSettings.Parse(_field.CodeText); }
    }

    private void Change(Action<TocSettings> change)
    {
        TocSettings settings = Settings;
        change(settings);
        _field.CodeText = settings.ToCode().Trim();
    }

    /// <summary>
    /// The styles the table makes entries from besides the heading styles, as <c>TableOfContents.HeadingStyles</c> in Word (the <c>\t</c>
    /// switch of the field). Changing them changes the field code and makes the entries again, as <see cref="Update"/> does.
    /// </summary>
    public HeadingStyleWrapperCollection HeadingStyles
    {
        get { return new HeadingStyleWrapperCollection(this); }
    }

    internal List<KeyValuePair<string, int>> AddedStyles()
    {
        return Settings.AddedStyles;
    }

    internal void ChangeAddedStyles(Action<List<KeyValuePair<string, int>>> change)
    {
        //Word makes the entries again as soon as the styles change.
        Change(settings => change(settings.AddedStyles));
        Update();
    }

    /// <summary>The table of contents, from the start of its field to its end, as <c>TableOfContents.Range</c> in Word.</summary>
    public RangeWrapper Range
    {
        get { return _field.Range; }
    }

    /// <summary>The range the table of contents is in, as <c>TableOfContents.Parent</c> in Word.</summary>
    public RangeWrapper Parent
    {
        get { return _field.Parent; }
    }

    /// <summary>True if the table is made from the heading styles, as <c>TableOfContents.UseHeadingStyles</c> in Word.</summary>
    public bool UseHeadingStyles
    {
        get { return Settings.UseHeadingStyles; }
        set { Change(s => s.UseHeadingStyles = value); }
    }

    /// <summary>The first heading level in the table (1 to 9), as <c>TableOfContents.UpperHeadingLevel</c> in Word.</summary>
    public int UpperHeadingLevel
    {
        get { return Settings.Upper; }
        set
        {
            CheckLevel(value);
            Change(s => s.Upper = value);
        }
    }

    /// <summary>The last heading level in the table (1 to 9), as <c>TableOfContents.LowerHeadingLevel</c> in Word.</summary>
    public int LowerHeadingLevel
    {
        get { return Settings.Lower; }
        set
        {
            CheckLevel(value);
            Change(s => s.Lower = value);
        }
    }

    private static void CheckLevel(int level)
    {
        if (level < 1 || level > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(level));
        }
    }

    /// <summary>True if the table has page numbers, as <c>TableOfContents.IncludePageNumbers</c> in Word.</summary>
    public bool IncludePageNumbers
    {
        get { return Settings.IncludePageNumbers; }
        set { Change(s => s.IncludePageNumbers = value); }
    }

    /// <summary>True if the page numbers are at the right margin, as <c>TableOfContents.RightAlignPageNumbers</c> in Word.</summary>
    public bool RightAlignPageNumbers
    {
        get { return Settings.RightAlign; }
        set { Change(s => s.RightAlign = value); }
    }

    /// <summary>True if the entries are hyperlinks to their headings, as <c>TableOfContents.UseHyperlinks</c> in Word.</summary>
    public bool UseHyperlinks
    {
        get { return Settings.UseHyperlinks; }
        set { Change(s => s.UseHyperlinks = value); }
    }

    /// <summary>True if a web page leaves the page numbers out, as <c>TableOfContents.HidePageNumbersInWeb</c> in Word.</summary>
    public bool HidePageNumbersInWeb
    {
        get { return Settings.HidePageNumbersInWeb; }
        set { Change(s => s.HidePageNumbersInWeb = value); }
    }

    /// <summary>True if the table uses TC fields, as <c>TableOfContents.UseFields</c> in Word. <see cref="Update"/> does not read TC fields.</summary>
    public bool UseFields
    {
        get { return Settings.UseFields; }
        set { Change(s => s.UseFields = value); }
    }

    /// <summary>The identifier of the kind of TC fields the table uses, as <c>TableOfContents.TableID</c> in Word.</summary>
    public string TableID
    {
        get { return Settings.TableId ?? string.Empty; }
        set { Change(s => s.TableId = string.IsNullOrEmpty(value) ? null : value); }
    }

    private IEnumerable<Paragraph> EntryParagraphs()
    {
        FieldChar begin = BeginChar();
        Paragraph last = EndParagraph();
        if (begin == null || last == null)
        {
            yield break;
        }
        for (Paragraph p = begin.Ancestors<Paragraph>().First(); p != null && !ReferenceEquals(p, last); p = p.NextSibling<Paragraph>())
        {
            yield return p;
        }
    }

    private FieldChar BeginChar()
    {
        return (_field.Range.StartElement as Run)?.GetFirstChild<FieldChar>();
    }

    private Paragraph EndParagraph()
    {
        return (_field.Range.EndElement as Run)?.Ancestors<Paragraph>().FirstOrDefault();
    }

    /// <summary>The leader between an entry and its page number, as <c>TableOfContents.TabLeader</c> in Word.</summary>
    public TabLeader TabLeader
    {
        get
        {
            TabStop tab = EntryParagraphs().Select(p => p.ParagraphProperties?.Tabs?.Elements<TabStop>().FirstOrDefault()).FirstOrDefault(t => t != null);
            return tab?.Leader?.Value == null ? TabLeader.Dots : FromLeader(tab.Leader.Value);
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            foreach (TabStop tab in EntryParagraphs().SelectMany(p => p.ParagraphProperties?.Tabs?.Elements<TabStop>() ?? Enumerable.Empty<TabStop>()))
            {
                tab.Leader = ToLeader(value);
            }
        }
    }

    internal static TabStopLeaderCharValues ToLeader(TabLeader leader)
    {
        switch (leader)
        {
            case TabLeader.Spaces:
                return TabStopLeaderCharValues.None;
            case TabLeader.Dashes:
                return TabStopLeaderCharValues.Hyphen;
            case TabLeader.Lines:
                return TabStopLeaderCharValues.Underscore;
            case TabLeader.Heavy:
                return TabStopLeaderCharValues.Heavy;
            case TabLeader.MiddleDot:
                return TabStopLeaderCharValues.MiddleDot;
            default:
                return TabStopLeaderCharValues.Dot;
        }
    }

    internal static TabLeader FromLeader(TabStopLeaderCharValues leader)
    {
        return leader == TabStopLeaderCharValues.None ? TabLeader.Spaces
            : leader == TabStopLeaderCharValues.Hyphen ? TabLeader.Dashes
            : leader == TabStopLeaderCharValues.Underscore ? TabLeader.Lines
            : leader == TabStopLeaderCharValues.Heavy ? TabLeader.Heavy
            : leader == TabStopLeaderCharValues.MiddleDot ? TabLeader.MiddleDot : TabLeader.Dots;
    }

    /// <summary>
    /// Takes the table of contents away, as <c>TableOfContents.Delete</c> in Word: its field and entries go and the text around it stays.
    /// The bookmarks Word put on the headings stay too.
    /// </summary>
    public void Delete()
    {
        FieldChar begin = BeginChar();
        if (begin == null)
        {
            _field.Delete();
            return;
        }
        Paragraph first = begin.Ancestors<Paragraph>().First();
        Paragraph last = EndParagraph();
        Run endRun = _field.Range.EndElement as Run;

        if (ReferenceEquals(first, last))
        {
            _field.Delete();
            return;
        }
        List<Paragraph> entries = EntryParagraphs().ToList();
        endRun.Remove();
        foreach (Paragraph entry in entries)
        {
            entry.Remove();
        }
    }

    /// <summary>
    /// Rebuilds the entries from the headings, as <c>TableOfContents.Update</c> in Word. An entry for a heading the table already had keeps
    /// its page number; Word works the others out when the document is opened.
    /// </summary>
    public void Update()
    {
        TocSettings settings = Settings;
        TabLeader leader = TabLeader;
        Dictionary<string, string> pages = TableOfContentsWrapperCollection.CachedPages(this);

        Paragraph target = _field.Range.EndElement is Run end && end.GetFirstChild<FieldChar>() != null
            ? EndParagraph() : _field.Parent.ParagraphOf();
        Delete();
        //The object goes on standing for the table, as in Word.
        _field = TableOfContentsWrapperCollection.Build(target, settings, leader, pages, true).Field;
    }

    #region Equality

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as TableOfContentsWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(TableOfContentsWrapper other)
    {
        return !(other is null) && _field == other._field;
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _field.GetHashCode();
    }

    public static bool operator ==(TableOfContentsWrapper left, TableOfContentsWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(TableOfContentsWrapper left, TableOfContentsWrapper right)
    {
        return !(left == right);
    }

    #endregion
}

internal static class RangeParagraphExtensions
{
    internal static Paragraph ParagraphOf(this RangeWrapper range)
    {
        return range.StartElement as Paragraph ?? range.StartElement?.Ancestors<Paragraph>().FirstOrDefault();
    }
}

/// <summary>
/// The tables of contents of a range or a document, as <c>TablesOfContents</c> in Word.
/// </summary>
public class TableOfContentsWrapperCollection : IEnumerable<TableOfContentsWrapper>
{
    private readonly RangeWrapper _range;

    internal TableOfContentsWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the tables are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<TableOfContentsWrapper> GetItems()
    {
        return _range.Fields
            .Where(f => f.Type == FieldType.TableOfContents && !Regex.IsMatch(f.CodeText, "\\\\[ac]\\b"))
            .Select(f => new TableOfContentsWrapper(f)).ToList();
    }

    /// <summary>The number of tables of contents.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The table at the index (0-based), as <c>TablesOfContents.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no table at the index.</exception>
    public TableOfContentsWrapper this[int index]
    {
        get
        {
            List<TableOfContentsWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Makes a table of contents, as <c>TablesOfContents.Add</c> in Word. Its entries go in front of the paragraph the range is at the start of.
    /// </summary>
    /// <param name="range">Where to put the table; it must be at the start of a paragraph.</param>
    /// <param name="useHeadingStyles">True to make the entries from the heading styles.</param>
    /// <param name="upperHeadingLevel">The first heading level (1 to 9).</param>
    /// <param name="lowerHeadingLevel">The last heading level (1 to 9).</param>
    /// <param name="useFields">True if the table uses TC fields (they are not read when the entries are made).</param>
    /// <param name="tableID">The identifier of the kind of TC fields.</param>
    /// <param name="rightAlignPageNumbers">True to put the page numbers at the right margin.</param>
    /// <param name="includePageNumbers">True to give the entries page numbers.</param>
    /// <param name="addedStyles">Other styles to make entries from, as pairs of a style name and a level separated by the list separator (<c>Title,1,Subtitle,2</c>).</param>
    /// <param name="useHyperlinks">True to make the entries hyperlinks to the headings.</param>
    /// <param name="hidePageNumbersInWeb">True to leave the page numbers out of a web page.</param>
    /// <param name="useOutlineLevels">True to make entries from paragraphs with an outline level too.</param>
    /// <exception cref="NotSupportedException">The range is not at the start of a paragraph.</exception>
    /// <exception cref="ArgumentException"><paramref name="addedStyles"/> is not pairs of a style name and a level from 1 to 9.</exception>
    public TableOfContentsWrapper Add(RangeWrapper range, bool useHeadingStyles = true, int upperHeadingLevel = 1, int lowerHeadingLevel = 9,
        bool useFields = false, string tableID = null, bool rightAlignPageNumbers = true, bool includePageNumbers = true, string addedStyles = null,
        bool useHyperlinks = false, bool hidePageNumbersInWeb = false, bool useOutlineLevels = false)
    {
        Guard.ThrowIfNull(range);
        if (upperHeadingLevel < 1 || upperHeadingLevel > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(upperHeadingLevel));
        }
        if (lowerHeadingLevel < upperHeadingLevel || lowerHeadingLevel > 9)
        {
            throw new ArgumentOutOfRangeException(nameof(lowerHeadingLevel));
        }

        RangeWrapper at = range.Duplicate();
        at.Collapse(CollapseDirection.Start);
        Paragraph paragraph = at.Paragraphs[0].Paragraph;
        if (at.Start != new RangeWrapper(paragraph).Start)
        {
            throw new NotSupportedException("A table of contents can be added only at the start of a paragraph.");
        }

        TocSettings settings = new TocSettings
        {
            UseHeadingStyles = useHeadingStyles,
            Upper = upperHeadingLevel,
            Lower = lowerHeadingLevel,
            UseFields = useFields,
            TableId = string.IsNullOrEmpty(tableID) ? null : tableID,
            RightAlign = rightAlignPageNumbers,
            IncludePageNumbers = includePageNumbers,
            UseHyperlinks = useHyperlinks,
            HidePageNumbersInWeb = hidePageNumbersInWeb,
            UseOutlineLevels = useOutlineLevels,
            AddedStyles = TocSettings.ParseStyles(addedStyles)
        };
        return Build(paragraph, settings, TabLeader.Dots, new Dictionary<string, string>());
    }

    //Entries Word makes when it updates a table also name the East Asian theme font and language in their paragraph marks.
    private static ParagraphMarkRunProperties MarkProperties(bool updating)
    {
        return updating
            ? new ParagraphMarkRunProperties(new RunFonts { EastAsiaTheme = ThemeFontValues.MinorEastAsia }, new NoProof(), new Languages { EastAsia = CultureInfo.CurrentCulture.Name })
            : new ParagraphMarkRunProperties(new NoProof());
    }

    private static Run Plain(params OpenXmlElement[] content)
    {
        return new Run(content);
    }

    private static Run NoProofRun(params OpenXmlElement[] content)
    {
        Run run = new Run(new RunProperties(new NoProof()));
        foreach (OpenXmlElement element in content)
        {
            run.AppendChild(element);
        }
        return run;
    }

    internal static Dictionary<string, string> CachedPages(TableOfContentsWrapper toc)
    {
        Dictionary<string, string> pages = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (FieldWrapper field in toc.Range.Fields)
        {
            Match match = Regex.Match(field.CodeText, "^PAGEREF\\s+(?<name>\\S+)");
            if (match.Success)
            {
                pages[match.Groups["name"].Value] = field.ResultText;
            }
        }
        return pages;
    }

    private static int LevelOf(Paragraph paragraph, TocSettings settings, Styles styles)
    {
        string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (settings.UseHeadingStyles && styleId != null)
        {
            string name = StyleHelper.GetStyleName(paragraph, styleId);
            Match match = Regex.Match(name ?? string.Empty, "^heading ([1-9])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (match.Success)
            {
                int level = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
                if (level >= settings.Upper && level <= settings.Lower)
                {
                    return level;
                }
            }
        }
        //With the outline levels Word makes no entries from the added styles.
        if (styleId != null && settings.AddedStyles.Count > 0 && !settings.UseOutlineLevels)
        {
            string styleName = StyleHelper.GetStyleName(paragraph, styleId);
            foreach (KeyValuePair<string, int> added in settings.AddedStyles)
            {
                if (string.Equals(added.Key, styleName, StringComparison.OrdinalIgnoreCase))
                {
                    return added.Value;
                }
            }
        }
        if (settings.UseOutlineLevels)
        {
            int? outline = paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
            for (Style style = styles?.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId); outline == null && style != null;
                style = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == style.BasedOn?.Val?.Value))
            {
                outline = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
                if (style.BasedOn?.Val?.Value == null)
                {
                    break;
                }
            }
            if (outline.HasValue && outline.Value >= 0 && outline.Value <= 8 && outline.Value + 1 >= settings.Upper && outline.Value + 1 <= settings.Lower)
            {
                return outline.Value + 1;
            }
        }
        return 0;
    }

    private static string TextOf(Paragraph paragraph)
    {
        return string.Concat(paragraph.Descendants<Text>().Where(t => !t.Ancestors<FieldCode>().Any()).Select(t => t.Text));
    }

    /// <summary>
    /// Makes the field and its entries in front of the paragraph. The end of the field goes at the start of the paragraph, as Word puts it.
    /// </summary>
    internal static TableOfContentsWrapper Build(Paragraph target, TocSettings settings, TabLeader leader, Dictionary<string, string> cachedPages, bool updating = false)
    {
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(target);
        OpenXmlElement root = target.GetRoot();
        Styles styles = main?.StyleDefinitionsPart?.Styles;
        string code = settings.ToCode();

        List<(Paragraph Heading, int Level)> headings = new List<(Paragraph, int)>();
        foreach (Paragraph paragraph in root.Descendants<Paragraph>())
        {
            int level = LevelOf(paragraph, settings, styles);
            if (level > 0 && !paragraph.Ancestors<Hyperlink>().Any())
            {
                headings.Add((paragraph, level));
            }
        }

        if (headings.Count == 0)
        {
            SimpleField none = new SimpleField(new Run(
                new RunProperties(new Bold(), new BoldComplexScript(), new NoProof(), new Languages { Val = "en-US" }),
                OpenXmlExtensions.NewText("No table of contents entries found.")))
            { Instruction = code };
            OpenXmlElement properties = target.ParagraphProperties;
            if (properties != null)
            {
                properties.InsertAfterSelf(none);
            }
            else
            {
                target.PrependChild(none);
            }
            return new TableOfContentsWrapper(new FieldWrapper(none));
        }

        int width = TextWidth(target);
        uint nextBookmark = main == null ? 0U : main.GetStoryRoots().Where(r => r != null).SelectMany(r => r.Descendants<BookmarkStart>())
            .Select(b => uint.TryParse(b.Id?.Value, NumberStyles.None, CultureInfo.InvariantCulture, out uint id) ? id : 0U).DefaultIfEmpty(0U).Max() + 1;
        long nextName = Math.Max(NextTocNumber(root), 0);
        if (nextName == 0)
        {
            nextName = 100000000L + (ListDefaults.NewRandom() & 0x3FFFFFFF) % 800000000L;
        }

        if (settings.UseHyperlinks)
        {
            BuiltInStyles.GetOrCreate(main, "Hyperlink");
        }

        FieldChar beginChar = null;
        bool first = true;
        foreach ((Paragraph heading, int level) in headings)
        {
            BookmarkStart bookmark = heading.Elements<BookmarkStart>().FirstOrDefault(b => b.Name?.Value != null && b.Name.Value.StartsWith("_Toc", StringComparison.Ordinal));
            if (bookmark == null)
            {
                bookmark = new BookmarkStart
                {
                    Id = (nextBookmark++).ToString(CultureInfo.InvariantCulture),
                    Name = "_Toc" + (nextName++).ToString(CultureInfo.InvariantCulture)
                };
                BookmarkEnd bookmarkEnd = new BookmarkEnd { Id = bookmark.Id.Value };
                OpenXmlElement properties = heading.ParagraphProperties;
                if (properties != null)
                {
                    properties.InsertAfterSelf(bookmark);
                }
                else
                {
                    heading.PrependChild(bookmark);
                }
                heading.AppendChild(bookmarkEnd);
            }
            string name = bookmark.Name.Value;

            Style tocStyle = BuiltInStyles.GetOrCreate(main, "toc " + level.ToString(CultureInfo.InvariantCulture));
            Paragraph entry = new Paragraph(new ParagraphProperties(
                new ParagraphStyleId { Val = tocStyle?.StyleId?.Value ?? "TOC" + level.ToString(CultureInfo.InvariantCulture) },
                new Tabs(new TabStop { Val = TabStopValues.Right, Leader = TableOfContentsWrapper.ToLeader(leader), Position = width }),
                MarkProperties(updating)));

            if (first)
            {
                beginChar = new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = true };
                entry.AppendChild(Plain(beginChar));
                entry.AppendChild(Plain(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }));
                entry.AppendChild(Plain(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                first = false;
            }

            OpenXmlCompositeElement container = entry;
            if (settings.UseHyperlinks)
            {
                Hyperlink link = new Hyperlink { Anchor = name };
                link.SetAttribute(new OpenXmlAttribute("w", "history", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "1"));
                entry.AppendChild(link);
                container = link;
            }
            Run text = NoProofRun(OpenXmlExtensions.NewText(TextOf(heading)));
            if (settings.UseHyperlinks)
            {
                text.RunProperties.InsertAt(new RunStyle { Val = "Hyperlink" }, 0);
            }
            container.AppendChild(text);
            if (settings.IncludePageNumbers)
            {
                container.AppendChild(NoProofRun(new TabChar()));
                container.AppendChild(NoProofRun(new FieldChar { FieldCharType = FieldCharValues.Begin }));
                container.AppendChild(NoProofRun(new FieldCode(" PAGEREF " + name + " \\h ") { Space = SpaceProcessingModeValues.Preserve }));
                container.AppendChild(NoProofRun());
                container.AppendChild(NoProofRun(new FieldChar { FieldCharType = FieldCharValues.Separate }));
                container.AppendChild(NoProofRun(OpenXmlExtensions.NewText(cachedPages.TryGetValue(name, out string page) ? page : "1")));
                container.AppendChild(NoProofRun(new FieldChar { FieldCharType = FieldCharValues.End }));
            }
            target.InsertBeforeSelf(entry);
        }

        Run end = Plain(new FieldChar { FieldCharType = FieldCharValues.End });
        OpenXmlElement targetProperties = target.ParagraphProperties;
        if (targetProperties != null)
        {
            targetProperties.InsertAfterSelf(end);
        }
        else
        {
            target.PrependChild(end);
        }
        return new TableOfContentsWrapper(new FieldWrapper(beginChar));
    }

    //The width between the margins, less the 10 twips Word leaves, in twips.
    private static int TextWidth(Paragraph paragraph)
    {
        PageSetupWrapper page = new RangeWrapper(paragraph).Sections[0].PageSetup;
        return (int)Math.Round((page.PageWidth - page.LeftMargin - page.RightMargin) * 20) - 10;
    }

    private static long NextTocNumber(OpenXmlElement root)
    {
        long highest = 0;
        foreach (BookmarkStart bookmark in root.Descendants<BookmarkStart>())
        {
            string name = bookmark.Name?.Value ?? string.Empty;
            if (name.StartsWith("_Toc", StringComparison.Ordinal) && long.TryParse(name.Remove(0, 4), NumberStyles.None, CultureInfo.InvariantCulture, out long number))
            {
                highest = Math.Max(highest, number);
            }
        }
        return highest == 0 ? 0 : highest + 1;
    }

    /// <inheritdoc/>
    public IEnumerator<TableOfContentsWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
