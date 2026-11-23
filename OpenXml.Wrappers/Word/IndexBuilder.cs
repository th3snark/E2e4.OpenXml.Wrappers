using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The switches of an <c>INDEX</c> field, as Word reads them.
/// </summary>
internal sealed class IndexSettings
{
    private static readonly HashSet<char> WithArgument = new HashSet<char> { 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'k', 'l', 'p', 's', 'y', 'z' };

    public readonly Dictionary<char, string> Switches = new Dictionary<char, string>();

    public bool Has(char name)
    {
        return Switches.ContainsKey(name);
    }

    public string Get(char name)
    {
        return Switches.TryGetValue(name, out string value) ? value : null;
    }

    public bool RunIn
    {
        get { return Has('r'); }
    }

    public int Columns
    {
        get { return int.TryParse(Get('c'), NumberStyles.Integer, CultureInfo.InvariantCulture, out int columns) ? columns : 0; }
    }

    public static IndexSettings Parse(string code)
    {
        IndexSettings settings = new IndexSettings();
        List<KeyValuePair<bool, string>> tokens = FieldCodeTokens.Split(code, false);
        for (int i = 1; i < tokens.Count; i++)
        {
            if (!tokens[i].Key)
            {
                continue;
            }
            char name = char.ToLowerInvariant(tokens[i].Value[0]);
            string argument = null;
            if (WithArgument.Contains(name) && i + 1 < tokens.Count && !tokens[i + 1].Key)
            {
                argument = tokens[++i].Value;
            }
            settings.Switches[name] = argument;
        }
        return settings;
    }
}

/// <summary>
/// Splits a field code into its words and switches, as Word reads them: a quoted argument is one word (it may hold <c>\"</c> for a quote and
/// <c>\\</c> for a backslash when <c>unescape</c> is set; other backslashes are kept), and <c>\x</c> is a switch.
/// </summary>
internal static class FieldCodeTokens
{
    //Each token is (is a switch, text).
    public static List<KeyValuePair<bool, string>> Split(string code, bool unescape)
    {
        List<KeyValuePair<bool, string>> tokens = new List<KeyValuePair<bool, string>>();
        string text = code ?? string.Empty;
        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (char.IsWhiteSpace(c))
            {
                i++;
            }
            else if (c == '\\' && i + 1 < text.Length && !char.IsWhiteSpace(text[i + 1]) && text[i + 1] != '"')
            {
                tokens.Add(new KeyValuePair<bool, string>(true, text.Substring(i + 1, 1)));
                i += 2;
            }
            else if (c == '"' || c == '“' || c == '”')
            {
                StringBuilder value = new StringBuilder();
                for (i++; i < text.Length && text[i] != '"' && text[i] != '“' && text[i] != '”'; i++)
                {
                    if (unescape && text[i] == '\\' && i + 1 < text.Length && (text[i + 1] == '"' || text[i + 1] == '\\'))
                    {
                        i++;
                    }
                    value.Append(text[i]);
                }
                i++;
                tokens.Add(new KeyValuePair<bool, string>(false, value.ToString()));
            }
            else
            {
                int start = i;
                while (i < text.Length && !char.IsWhiteSpace(text[i]))
                {
                    i++;
                }
                tokens.Add(new KeyValuePair<bool, string>(false, text.Substring(start, i - start)));
            }
        }
        return tokens;
    }
}

/// <summary>
/// An <c>XE</c> field: its entry (the text of each level), and switches.
/// </summary>
internal sealed class IndexEntryField
{
    public FieldWrapper Field;
    public List<string> Levels;
    public bool Bold;
    public bool Italic;
    public string CrossReference;
    public string RangeBookmark;
    public string Type;

    //Where the field is in the order of the main text: the field itself, or the mark of the note or the anchor of the text box it is in.
    public long Order;
    public OpenXmlElement Anchor;

    public static IndexEntryField Parse(FieldWrapper field)
    {
        List<KeyValuePair<bool, string>> tokens = FieldCodeTokens.Split(field.CodeText, true);
        IndexEntryField entry = new IndexEntryField() { Field = field };
        string text = null;
        for (int i = 1; i < tokens.Count; i++)
        {
            if (!tokens[i].Key)
            {
                text = text ?? tokens[i].Value;
                continue;
            }
            char name = char.ToLowerInvariant(tokens[i].Value[0]);
            string argument = null;
            if ((name == 't' || name == 'r' || name == 'f' || name == 'y') && i + 1 < tokens.Count && !tokens[i + 1].Key)
            {
                argument = tokens[++i].Value;
            }
            switch (name)
            {
                case 'b':
                    entry.Bold = true;
                    break;
                case 'i':
                    entry.Italic = true;
                    break;
                case 't':
                    entry.CrossReference = argument ?? string.Empty;
                    break;
                case 'r':
                    entry.RangeBookmark = argument;
                    break;
                case 'f':
                    entry.Type = argument;
                    break;
            }
        }
        entry.Levels = text == null ? new List<string>() : SplitLevels(text);
        return entry;
    }

    //"Fruit:Banana" is two levels, each without the spaces around it; "\:" is a colon in the text. A colon at the end is kept as text, as in
    //Word ("Trail:").
    private static List<string> SplitLevels(string text)
    {
        List<string> levels = new List<string>();
        StringBuilder level = new StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] == '\\' && i + 1 < text.Length && text[i + 1] == ':')
            {
                level.Append(':');
                i++;
            }
            else if (text[i] == ':')
            {
                levels.Add(level.ToString().Trim());
                level.Length = 0;
            }
            else
            {
                level.Append(text[i]);
            }
        }
        string last = level.ToString().Trim();
        if (last.Length == 0 && levels.Count > 0)
        {
            levels[levels.Count - 1] += ":";
        }
        else
        {
            levels.Add(last);
        }
        return levels.Where(l => l.Length > 0).ToList();
    }
}

/// <summary>
/// The order of index entries, as Word sorts them: letters by their letter, whatever its accent or case (and, between those that are the same,
/// without an accent first, then lower case first); symbols before digits and digits before letters, one character at a time ("b10" before
/// "b2"); a hyphen or an apostrophe is passed over ("applet" before "Apple-tart"), and the word without it comes first ("coop" before
/// "co-op").
/// </summary>
internal sealed class IndexEntryComparer : IComparer<string>
{
    public static readonly IndexEntryComparer Instance = new IndexEntryComparer();

    private static bool Ignored(char c)
    {
        return c == '-' || c == '\'' || c == '’' || c == '­';
    }

    //The letter of a character without its accent, and whether it had one.
    internal static char BaseLetter(char c, out bool accented)
    {
        string decomposed = c.ToString().Normalize(NormalizationForm.FormD);
        accented = decomposed.Length > 1;
        return decomposed[0];
    }

    public int Compare(string x, string y)
    {
        if (ReferenceEquals(x, y))
        {
            return 0;
        }
        string a = new string((x ?? string.Empty).Where(c => !Ignored(c)).ToArray());
        string b = new string((y ?? string.Empty).Where(c => !Ignored(c)).ToArray());

        //The letters, digits and symbols.
        int length = Math.Min(a.Length, b.Length);
        for (int i = 0; i < length; i++)
        {
            int primary = Primary(a[i]).CompareTo(Primary(b[i]));
            if (primary != 0)
            {
                return primary;
            }
        }
        if (a.Length != b.Length)
        {
            return a.Length.CompareTo(b.Length);
        }
        //The accents, then the case, then the hyphens and apostrophes left out.
        for (int i = 0; i < length; i++)
        {
            BaseLetter(a[i], out bool accentA);
            BaseLetter(b[i], out bool accentB);
            if (accentA != accentB)
            {
                return accentA ? 1 : -1;
            }
        }
        for (int i = 0; i < length; i++)
        {
            bool upperA = char.IsUpper(a[i]);
            bool upperB = char.IsUpper(b[i]);
            if (upperA != upperB)
            {
                return upperA ? 1 : -1;
            }
        }
        int ignored = (x ?? string.Empty).Length.CompareTo((y ?? string.Empty).Length);
        return ignored != 0 ? ignored : string.CompareOrdinal(x, y);
    }

    //Symbols (by their code), then digits, then letters (by their letter in lower case, without an accent).
    private static long Primary(char c)
    {
        if (char.IsLetter(c))
        {
            return (2L << 32) + char.ToLowerInvariant(BaseLetter(c, out bool _));
        }
        if (char.IsDigit(c))
        {
            return (1L << 32) + c;
        }
        return c;
    }
}

/// <summary>
/// Works out the result of an <c>INDEX</c> field from the document's <c>XE</c> fields and writes it as Word does: a paragraph for each entry
/// in the "index 1" to "index 9" styles, letter headings in "index heading", and with <c>\c</c> a section of columns.
/// </summary>
internal static class IndexBuilder
{
    private const string Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string NoEntries = "No index entries found.";

    private sealed class Node
    {
        public string Text;
        public readonly Dictionary<string, Node> Children = new Dictionary<string, Node>(StringComparer.Ordinal);
        public readonly List<Item> Items = new List<Item>();

        public Node Child(string text)
        {
            if (!Children.TryGetValue(text, out Node child))
            {
                child = new Node() { Text = text };
                Children[text] = child;
            }
            return child;
        }

        public IEnumerable<Node> Sorted()
        {
            return Children.Values.OrderBy(c => c.Text, IndexEntryComparer.Instance);
        }
    }

    //A page number (or range) or a cross-reference of an entry, in the order of the text.
    private sealed class Item
    {
        public long Order;
        public string Text;
        public bool CrossReference;
        public bool Bold;
        public bool Italic;
    }

    private sealed class Piece
    {
        public string Text;
        public bool Bold;
        public bool Italic;
    }

    private sealed class Line
    {
        public int Level;
        public readonly List<Piece> Pieces = new List<Piece>();
    }

    /// <summary>
    /// Builds the result of the INDEX field again. <paramref name="pageNumber"/> gives the page number shown at a place (an XE field, or the start
    /// or end of a bookmark); when it is null the numbers are estimated from the breaks in the text and the field is marked for Word to update.
    /// </summary>
    public static void Update(FieldWrapper field, Func<RangeWrapper, string> pageNumber)
    {
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(field.Key);
        Body body = main?.Document?.Body;
        if (body == null || !(StoryMap.FindRoot(field.Key) is Body))
        {
            throw new NotSupportedException("An index can be built only in the main text of a document.");
        }
        field.EnsureComplex();
        field.WriteCodeTabs();
        IndexSettings settings = IndexSettings.Parse(field.CodeText);
        TabStopLeaderCharValues leader = Leader(field);

        Func<RangeWrapper, string> pages = pageNumber;
        PageEstimator estimator = null;
        if (pages == null)
        {
            estimator = new PageEstimator(body);
            pages = estimator.PageNumberAt;
        }
        List<Line> lines = BuildLines(field, settings, main, body, pages);
        Write(field, settings, main, body, lines, leader);
        field.SetDirty(estimator != null);
    }

    #region Entries

    private static List<Line> BuildLines(FieldWrapper index, IndexSettings settings, MainDocumentPart main, Body body, Func<RangeWrapper, string> pageNumber)
    {
        Dictionary<OpenXmlElement, long> order = new Dictionary<OpenXmlElement, long>();
        long next = 0;
        foreach (OpenXmlElement element in body.Descendants())
        {
            order[element] = next++;
        }

        List<IndexEntryField> fields = EntryFields(main, body, order);
        string type = settings.Get('f');
        string bookmarkName = settings.Get('b');
        long bookmarkFrom = long.MinValue;
        long bookmarkTo = long.MaxValue;
        if (bookmarkName != null)
        {
            BookmarkStart start = FindBookmark(body, bookmarkName, out BookmarkEnd end);
            bookmarkFrom = start == null ? long.MaxValue : order[start] * 1000;
            bookmarkTo = end == null ? long.MinValue : order[end] * 1000 + 999;
        }
        Dictionary<FieldWrapper, string> chapters = settings.Get('s') == null ? null : Chapters(body, settings.Get('s'), order);
        string chapterSeparator = settings.Get('d') ?? "-";
        string rangeSeparator = settings.Get('g') ?? "–";

        Node root = new Node();
        foreach (IndexEntryField entry in fields)
        {
            bool wanted = type == null
                ? entry.Type == null || string.Equals(entry.Type, "I", StringComparison.OrdinalIgnoreCase)
                : string.Equals(entry.Type, type, StringComparison.OrdinalIgnoreCase);
            if (!wanted || entry.Levels.Count == 0 || entry.Order < bookmarkFrom || entry.Order > bookmarkTo)
            {
                continue;
            }

            Node node = root;
            foreach (string level in entry.Levels)
            {
                node = node.Child(level);
            }
            if (entry.CrossReference != null)
            {
                node.Items.Add(new Item() { Order = entry.Order, Text = entry.CrossReference, CrossReference = true });
                continue;
            }

            string chapter = chapters != null && chapters.TryGetValue(entry.Field, out string value) ? value + chapterSeparator : string.Empty;
            string number = null;
            long itemOrder = entry.Order;
            BookmarkStart rangeStart = entry.RangeBookmark == null ? null : FindBookmark(body, entry.RangeBookmark, out BookmarkEnd rangeEnd);
            if (rangeStart != null)
            {
                BookmarkEnd end = FindEnd(body, rangeStart);
                RangeWrapper span = new BookmarkWrapper(rangeStart).Range;
                RangeWrapper from = span.Duplicate();
                from.Collapse(CollapseDirection.Start);
                RangeWrapper to = span.Duplicate();
                to.Collapse(CollapseDirection.End);
                string first = pageNumber(from);
                string last = end == null ? first : pageNumber(to);
                if (first != null)
                {
                    number = last == null || last == first ? chapter + first : chapter + first + rangeSeparator + chapter + last;
                    itemOrder = order.TryGetValue(rangeStart, out long startOrder) ? startOrder * 1000 : entry.Order;
                }
            }
            else
            {
                string page = pageNumber(entry.Field.Range);
                number = page == null ? null : chapter + page;
            }
            if (number != null)
            {
                node.Items.Add(new Item() { Order = itemOrder, Text = number, Bold = entry.Bold, Italic = entry.Italic });
            }
        }

        return Lines(root, settings);
    }

    //The XE fields of the main text (with its text boxes), the footnotes and the endnotes, in the order of the main text. Those in headers,
    //footers and comments are not in an index, as in Word.
    private static List<IndexEntryField> EntryFields(MainDocumentPart main, Body body, Dictionary<OpenXmlElement, long> order)
    {
        List<IndexEntryField> fields = new List<IndexEntryField>();
        void Collect(OpenXmlElement root, Func<OpenXmlElement, OpenXmlElement> anchorOf)
        {
            int within = 0;
            foreach (FieldWrapper field in new FieldWrapperCollection(new RangeWrapper(root)))
            {
                if (field.Type != FieldType.IndexEntry)
                {
                    continue;
                }
                IndexEntryField entry = IndexEntryField.Parse(field);
                OpenXmlElement anchor = anchorOf(field.Key);
                entry.Anchor = anchor;
                entry.Order = anchor != null && order.TryGetValue(anchor, out long at) ? at * 1000 + Math.Min(within++, 999) : long.MaxValue;
                fields.Add(entry);
            }
        }

        Collect(body, key => key.Ancestors<TextBoxContent>().Any() ? key.Ancestors<TextBoxContent>().Last().Ancestors<Run>().FirstOrDefault() ?? key : key);
        foreach (FootnoteEndnoteType note in NoteOps.GetNotes(main, false).Concat(NoteOps.GetNotes(main, true)))
        {
            FootnoteEndnoteReferenceType reference = NoteOps.FindReference(main, note is Endnote, note);
            Collect(note, key => reference);
        }
        return fields.OrderBy(f => f.Order).ToList();
    }

    private static BookmarkStart FindBookmark(Body body, string name, out BookmarkEnd end)
    {
        BookmarkStart start = body.Descendants<BookmarkStart>().FirstOrDefault(b => string.Equals(b.Name?.Value, name, StringComparison.OrdinalIgnoreCase));
        end = start == null ? null : FindEnd(body, start);
        return start;
    }

    private static BookmarkEnd FindEnd(Body body, BookmarkStart start)
    {
        return body.Descendants<BookmarkEnd>().FirstOrDefault(e => e.Id?.Value == start.Id?.Value);
    }

    //The number of the SEQ identifier at each XE field, as INDEX \s shows it: that of the last SEQ field of the identifier before it, in its
    //number format, and 0 before the first.
    private static Dictionary<FieldWrapper, string> Chapters(Body body, string identifier, Dictionary<OpenXmlElement, long> order)
    {
        Dictionary<FieldWrapper, string> chapters = new Dictionary<FieldWrapper, string>();
        string current = "0";
        foreach (FieldWrapper field in new FieldWrapperCollection(new RangeWrapper(body)))
        {
            if (field.Type == FieldType.IndexEntry)
            {
                chapters[field] = current;
            }
            else if (field.Type == FieldType.Sequence)
            {
                string value = FieldUpdater.SequenceShown(field, identifier);
                if (value != null)
                {
                    current = value;
                }
            }
        }
        return chapters;
    }

    #endregion

    #region Lines

    private static List<Line> Lines(Node root, IndexSettings settings)
    {
        string heading = settings.Get('h');
        bool accented = settings.Has('a');
        KeyValuePair<char, char>? letters = LetterRange(settings.Get('p'));
        List<Line> lines = new List<Line>();
        string group = null;
        foreach (Node entry in root.Sorted())
        {
            string entryGroup = GroupOf(entry.Text, accented);
            if (letters.HasValue)
            {
                char first = entryGroup.Length == 1 ? entryGroup[0] : '\0';
                if (!char.IsLetter(first) || char.ToUpperInvariant(BaseOf(first)) < letters.Value.Key || char.ToUpperInvariant(BaseOf(first)) > letters.Value.Value)
                {
                    continue;
                }
            }
            if (!string.IsNullOrEmpty(heading) && entryGroup != group)
            {
                group = entryGroup;
                Line line = new Line() { Level = 0 };
                line.Pieces.Add(new Piece() { Text = HeadingText(heading, entryGroup) });
                lines.Add(line);
            }
            if (settings.RunIn)
            {
                lines.Add(RunIn(entry, settings));
            }
            else
            {
                AddIndented(entry, 1, lines, settings);
            }
        }
        return lines;
    }

    private static char BaseOf(char c)
    {
        return IndexEntryComparer.BaseLetter(c, out bool _);
    }

    //The letter heading an entry goes under: its first letter in capitals, without its accent unless \a; a digit or a symbol is its own.
    private static string GroupOf(string text, bool accented)
    {
        if (text.Length == 0)
        {
            return string.Empty;
        }
        char first = text[0];
        if (!char.IsLetter(first))
        {
            return first.ToString();
        }
        return char.ToUpperInvariant(accented ? first : BaseOf(first)).ToString();
    }

    //The heading of a group: the \h text with its "A" the letter of the group.
    private static string HeadingText(string template, string group)
    {
        int at = template.IndexOf('A');
        return at < 0 ? template : template.Remove(at, 1).Insert(at, group);
    }

    //\p "a--m": the first and last letters, in capitals.
    private static KeyValuePair<char, char>? LetterRange(string argument)
    {
        if (argument == null)
        {
            return null;
        }
        string[] ends = argument.Split(new[] { "--", "-" }, StringSplitOptions.RemoveEmptyEntries).Select(e => e.Trim()).ToArray();
        return ends.Length == 2 && ends[0].Length == 1 && ends[1].Length == 1
            ? new KeyValuePair<char, char>(char.ToUpperInvariant(ends[0][0]), char.ToUpperInvariant(ends[1][0]))
            : (KeyValuePair<char, char>?)null;
    }

    private static void AddIndented(Node node, int level, List<Line> lines, IndexSettings settings)
    {
        Line line = new Line() { Level = Math.Min(level, 9) };
        line.Pieces.Add(new Piece() { Text = node.Text });
        AddItems(node, line, settings);
        lines.Add(line);
        foreach (Node child in node.Sorted())
        {
            AddIndented(child, level + 1, lines, settings);
        }
    }

    //Run-in (\r): the subentries, at every level, follow the main entry on its line: "Fruit, 2; Banana, 1, 2; Ripe, 2", or "fruit: cherry, 1"
    //when the main entry has no pages of its own.
    private static Line RunIn(Node node, IndexSettings settings)
    {
        Line line = new Line() { Level = 1 };
        line.Pieces.Add(new Piece() { Text = node.Text });
        bool hasItems = AddItems(node, line, settings);
        bool first = true;
        foreach (Node descendant in Descendants(node))
        {
            line.Pieces.Add(new Piece() { Text = first && !hasItems ? ": " : "; " });
            line.Pieces.Add(new Piece() { Text = descendant.Text });
            AddItems(descendant, line, settings);
            first = false;
        }
        return line;
    }

    private static IEnumerable<Node> Descendants(Node node)
    {
        foreach (Node child in node.Sorted())
        {
            yield return child;
            foreach (Node grandchild in Descendants(child))
            {
                yield return grandchild;
            }
        }
    }

    //The page numbers of an entry (in the order of the text, each once, bold or italic as the first XE field on that page asks), then its
    //cross-references. The first comes after \e (", "), or for a cross-reference after \e, \k or ". "; the others after \l (", ").
    private static bool AddItems(Node node, Line line, IndexSettings settings)
    {
        List<Item> pages = node.Items.Where(i => !i.CrossReference).OrderBy(i => i.Order).GroupBy(i => i.Text, StringComparer.Ordinal).Select(g => g.First()).ToList();
        List<Item> crossReferences = node.Items.Where(i => i.CrossReference).OrderBy(i => i.Order).GroupBy(i => i.Text, StringComparer.Ordinal).Select(g => g.First()).ToList();
        string entrySeparator = settings.Get('e');
        string itemSeparator = settings.Get('l') ?? ", ";
        bool first = true;
        foreach (Item item in pages.Concat(crossReferences))
        {
            string separator = !first ? itemSeparator
                : item.CrossReference ? entrySeparator ?? settings.Get('k') ?? ". "
                : entrySeparator ?? ", ";
            line.Pieces.Add(new Piece() { Text = separator });
            line.Pieces.Add(new Piece() { Text = item.Text, Bold = item.Bold, Italic = item.Italic });
            first = false;
        }
        return !first;
    }

    #endregion

    #region Writing

    private static TabStopLeaderCharValues Leader(FieldWrapper field)
    {
        Paragraph paragraph = field.Key.Ancestors<Paragraph>().FirstOrDefault();
        for (Paragraph p = paragraph; p != null; p = p.NextSibling<Paragraph>())
        {
            TabStop tab = p.ParagraphProperties?.Tabs?.Elements<TabStop>().FirstOrDefault();
            if (tab?.Leader != null && tab.Leader.HasValue)
            {
                return tab.Leader.Value;
            }
            if (p.Descendants<FieldChar>().Any(c => c.FieldCharType != null && c.FieldCharType.HasValue && c.FieldCharType.Value == FieldCharValues.End))
            {
                break;
            }
        }
        return TabStopLeaderCharValues.Dot;
    }

    private static void Write(FieldWrapper field, IndexSettings settings, MainDocumentPart main, Body body, List<Line> lines, TabStopLeaderCharValues leader)
    {
        field.ClearResult(out Paragraph beginParagraph, out Paragraph endParagraph, out Run separateRun);

        if (lines.Count == 0)
        {
            //As in Word: the message, in the field's own paragraph.
            if (!ReferenceEquals(beginParagraph, endParagraph))
            {
                field.JoinResultParagraphs();
                field.ClearResult(out beginParagraph, out endParagraph, out separateRun);
            }
            separateRun.InsertAfterSelf(new Run(
                new RunProperties(new Bold(), new BoldComplexScript(), new NoProof(), new Languages() { Val = "en-US" }),
                OpenXmlExtensions.NewText(NoEntries)));
            return;
        }

        foreach (int level in lines.Select(l => l.Level).Distinct())
        {
            BuiltInStyles.GetOrCreate(main, level == 0 ? "index heading" : "index " + level.ToString(CultureInfo.InvariantCulture));
        }

        bool sectioned = settings.Has('c');
        int columns = Math.Max(settings.Columns, 1);
        Paragraph before = endParagraph;
        if (ReferenceEquals(beginParagraph, endParagraph))
        {
            //The field's end, and what follows it, stay in the paragraph, with its mark; the field's start goes in a paragraph of its own before.
            Paragraph first = new Paragraph();
            foreach (OpenXmlElement element in endParagraph.ChildElements.Where(e => !(e is ParagraphProperties)).TakeWhile(e => !e.Descendants<FieldChar>().Any(c => IsEnd(c, field))).ToList())
            {
                element.Remove();
                first.AppendChild(element);
            }
            endParagraph.InsertBeforeSelf(first);
            SpanRegistry.ParagraphSplit(first, endParagraph);
            beginParagraph = first;
        }
        if (sectioned)
        {
            if (beginParagraph.ParagraphProperties?.SectionProperties == null)
            {
                MakeSections(body, beginParagraph, endParagraph, columns);
            }
            before = endParagraph.PreviousSibling<Paragraph>() ?? endParagraph;
        }

        int width = TextWidth(endParagraph, columns);
        bool beginTakesFirst = !sectioned && !beginParagraph.Elements<Run>().TakeWhile(r => r != field.Key.Parent).Any(r => r.Elements<Text>().Any());
        List<Line> rest = lines;
        if (beginTakesFirst)
        {
            Paragraph template = NewParagraph(lines[0], width, leader, settings.RunIn);
            beginParagraph.ParagraphProperties?.Remove();
            beginParagraph.PrependChild(template.ParagraphProperties.CloneNode(true));
            OpenXmlElement after = separateRun;
            foreach (Run run in template.Elements<Run>().ToList())
            {
                run.Remove();
                after.InsertAfterSelf(run);
                after = run;
            }
            rest = lines.Skip(1).ToList();
        }
        foreach (Line line in rest)
        {
            before.InsertBeforeSelf(NewParagraph(line, width, leader, settings.RunIn));
        }
    }

    private static bool IsEnd(FieldChar c, FieldWrapper field)
    {
        return c.FieldCharType != null && c.FieldCharType.HasValue && c.FieldCharType.Value == FieldCharValues.End && ReferenceEquals(c, field.EndChar());
    }

    //\c: the index is a section of its own, between continuous section breaks, as Word makes it: the field's start ends the section before
    //(which keeps its properties), the paragraph before the field's end ends the index's section (the columns), and the section after starts
    //on the same page.
    private static void MakeSections(Body body, Paragraph beginParagraph, Paragraph endParagraph, int columns)
    {
        SectionProperties current = SectionOf(body, endParagraph);
        SectionProperties before = (SectionProperties)current.CloneNode(true);
        beginParagraph.ParagraphProperties?.Remove();
        beginParagraph.PrependChild(new ParagraphProperties(new ParagraphMarkRunProperties(new NoProof()), before));

        SectionProperties index = (SectionProperties)current.CloneNode(true);
        index.RemoveAllChildren<HeaderReference>();
        index.RemoveAllChildren<FooterReference>();
        SetSectionType(index);
        Columns cols = index.GetFirstChild<Columns>();
        if (cols == null)
        {
            cols = new Columns();
            index.AddChild(cols);
        }
        cols.Space = "720";
        cols.ColumnCount = columns > 1 ? (Int16Value)(short)columns : null;
        endParagraph.InsertBeforeSelf(new Paragraph(new ParagraphProperties(new ParagraphMarkRunProperties(new NoProof()), index)));
        SetSectionType(current);
    }

    private static void SetSectionType(SectionProperties section)
    {
        SectionType type = section.GetFirstChild<SectionType>();
        if (type == null)
        {
            type = new SectionType();
            section.AddChild(type);
        }
        type.Val = SectionMarkValues.Continuous;
    }

    //The properties of the section a paragraph is in: those of the paragraph that ends it, or the body's.
    private static SectionProperties SectionOf(Body body, Paragraph paragraph)
    {
        for (Paragraph p = paragraph; p != null; p = p.NextSibling<Paragraph>())
        {
            if (p.ParagraphProperties?.SectionProperties != null)
            {
                return p.ParagraphProperties.SectionProperties;
            }
        }
        SectionProperties last = body.GetFirstChild<SectionProperties>();
        if (last == null)
        {
            last = new SectionProperties();
            body.AppendChild(last);
        }
        return last;
    }

    //The width of the text (of a column), less the 10 twips Word leaves, where the tab stop for the page numbers goes.
    private static int TextWidth(Paragraph paragraph, int columns)
    {
        PageSetupWrapper page = new RangeWrapper(paragraph).Sections[0].PageSetup;
        int width = (int)Math.Round((page.PageWidth - page.LeftMargin - page.RightMargin) * 20);
        return (width - (columns - 1) * 720) / columns - 10;
    }

    private static Paragraph NewParagraph(Line line, int width, TabStopLeaderCharValues leader, bool runIn)
    {
        bool heading = line.Level == 0;
        ParagraphProperties properties = new ParagraphProperties(
            new ParagraphStyleId() { Val = heading ? "IndexHeading" : "Index" + line.Level.ToString(CultureInfo.InvariantCulture) });
        if (heading)
        {
            properties.AppendChild(new KeepNext());
        }
        properties.AppendChild(new Tabs(new TabStop() { Val = TabStopValues.Right, Leader = leader, Position = width }));

        Paragraph paragraph = new Paragraph(properties);
        bool boldSoFar = false;
        bool italicSoFar = false;
        foreach (Piece piece in line.Pieces)
        {
            if (piece.Text.Length == 0)
            {
                continue;
            }
            RunProperties runProperties = new RunProperties();
            if (piece.Bold)
            {
                runProperties.AppendChild(new Bold());
            }
            if (piece.Bold || boldSoFar)
            {
                runProperties.AppendChild(new BoldComplexScript());
            }
            if (piece.Italic)
            {
                runProperties.AppendChild(new Italic());
            }
            if (piece.Italic || italicSoFar)
            {
                runProperties.AppendChild(new ItalicComplexScript());
            }
            runProperties.AppendChild(new NoProof());
            boldSoFar |= piece.Bold;
            italicSoFar |= piece.Italic;

            Run run = new Run(runProperties);
            string[] parts = piece.Text.Split('\t');
            for (int i = 0; i < parts.Length; i++)
            {
                if (i > 0)
                {
                    run.AppendChild(new TabChar());
                }
                if (parts[i].Length > 0)
                {
                    run.AppendChild(OpenXmlExtensions.NewText(parts[i]));
                }
            }
            paragraph.AppendChild(run);
        }

        //The paragraph mark: a heading's as Word writes it, and an entry's with the complex-script bold or italic of its page numbers (not on a
        //run-in line).
        ParagraphMarkRunProperties mark = new ParagraphMarkRunProperties();
        if (heading)
        {
            mark.AppendChild(new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi });
            //Word writes 0 where the SDK would write false.
            Bold notBold = new Bold();
            notBold.SetAttribute(new OpenXmlAttribute("w", "val", Main, "0"));
            BoldComplexScript notBoldComplex = new BoldComplexScript();
            notBoldComplex.SetAttribute(new OpenXmlAttribute("w", "val", Main, "0"));
            mark.AppendChild(notBold);
            mark.AppendChild(notBoldComplex);
        }
        else if (!runIn)
        {
            if (boldSoFar)
            {
                mark.AppendChild(new BoldComplexScript());
            }
            if (italicSoFar)
            {
                mark.AppendChild(new ItalicComplexScript());
            }
        }
        mark.AppendChild(new NoProof());
        properties.AppendChild(mark);
        return paragraph;
    }

    #endregion
}
