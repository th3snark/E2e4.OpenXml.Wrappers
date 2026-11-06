using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Works out which list item a paragraph is and how it is numbered, as Word does: counters for each list template, a level starting its count
/// again when a higher level is used, and the text of the number made from the level's format.
/// </summary>
internal static class ListNumbering
{
    internal sealed class Item
    {
        public Paragraph Paragraph;
        public NumberingInstance Num;
        public AbstractNum Template;
        public int Level;
        public int Value;
        public string Text;
    }

    internal static Numbering GetNumbering(OpenXmlElement context)
    {
        return StyleHelper.GetMainDocumentPart(context)?.NumberingDefinitionsPart?.Numbering;
    }

    /// <summary>The list and level a paragraph is in, from its own properties or its style; false if it is in none.</summary>
    internal static bool TryGetNumbering(Paragraph paragraph, out int numId, out int level)
    {
        numId = 0;
        level = 0;
        NumberingProperties properties = paragraph.ParagraphProperties?.NumberingProperties;
        int? ownLevel = properties?.NumberingLevelReference?.Val?.Value;
        int? ownId = properties?.NumberingId?.Val?.Value;

        string styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        Styles styles = StyleHelper.GetStyles(paragraph);
        int guard = 0;
        while ((ownId == null || ownLevel == null) && styleId != null && styles != null && guard++ < 20)
        {
            Style style = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == styleId);
            NumberingProperties fromStyle = style?.StyleParagraphProperties?.NumberingProperties;
            ownId = ownId ?? fromStyle?.NumberingId?.Val?.Value;
            ownLevel = ownLevel ?? fromStyle?.NumberingLevelReference?.Val?.Value;
            styleId = style?.BasedOn?.Val?.Value;
        }

        if (ownId == null || ownId.Value == 0)
        {
            return false;
        }
        numId = ownId.Value;
        level = Math.Max(0, Math.Min(8, ownLevel ?? 0));
        return true;
    }

    private static NumberingInstance FindNum(Numbering numbering, int numId)
    {
        return numbering?.Elements<NumberingInstance>().FirstOrDefault(n => n.NumberID?.Value == numId);
    }

    private static AbstractNum FindAbstract(Numbering numbering, NumberingInstance num)
    {
        int? id = num?.AbstractNumId?.Val?.Value;
        return id == null ? null : numbering.Elements<AbstractNum>().FirstOrDefault(a => a.AbstractNumberId?.Value == id.Value);
    }

    internal static Level FindLevel(AbstractNum template, int level)
    {
        return template.Elements<Level>().FirstOrDefault(l => l.LevelIndex?.Value == level) ?? template.Elements<Level>().ElementAtOrDefault(level);
    }

    /// <summary>Every numbered paragraph of the story the paragraph is in, in order, with its number worked out.</summary>
    internal static List<Item> GetItems(OpenXmlElement storyRoot)
    {
        List<Item> items = new List<Item>();
        Numbering numbering = GetNumbering(storyRoot);
        if (numbering == null)
        {
            return items;
        }

        Dictionary<AbstractNum, int?[]> counters = new Dictionary<AbstractNum, int?[]>();
        Dictionary<AbstractNum, int?[]> starts = new Dictionary<AbstractNum, int?[]>();
        HashSet<int> seen = new HashSet<int>();
        foreach (Paragraph paragraph in storyRoot.Descendants<Paragraph>())
        {
            if (!TryGetNumbering(paragraph, out int numId, out int level))
            {
                continue;
            }
            NumberingInstance num = FindNum(numbering, numId);
            AbstractNum template = FindAbstract(numbering, num);
            if (template == null)
            {
                continue;
            }
            if (!counters.TryGetValue(template, out int?[] count))
            {
                count = counters[template] = new int?[9];
                starts[template] = new int?[9];
            }
            if (seen.Add(numId))
            {
                //A list that starts again at a number says so for the levels it overrides.
                foreach (LevelOverride levelOverride in num.Elements<LevelOverride>())
                {
                    int overridden = levelOverride.LevelIndex?.Value ?? 0;
                    if (overridden >= 0 && overridden < 9 && levelOverride.StartOverrideNumberingValue?.Val != null)
                    {
                        count[overridden] = null;
                        starts[template][overridden] = levelOverride.StartOverrideNumberingValue.Val.Value;
                    }
                }
            }

            int start(int l) => starts[template][l] ?? FindLevel(template, l)?.StartNumberingValue?.Val?.Value ?? 1;
            count[level] = count[level] == null ? start(level) : count[level].Value + 1;
            for (int deeper = level + 1; deeper < 9; deeper++)
            {
                int restart = FindLevel(template, deeper)?.LevelRestart?.Val?.Value ?? deeper;
                if (restart != 0 && level + 1 <= restart)
                {
                    count[deeper] = null;
                    starts[template][deeper] = null;
                }
            }

            int value = count[level].Value;
            string text = FormatText(template, level, l => count[l] ?? start(l));
            items.Add(new Item() { Paragraph = paragraph, Num = num, Template = template, Level = level, Value = value, Text = text });
        }
        return items;
    }

    private static string FormatText(AbstractNum template, int level, Func<int, int> value)
    {
        Level definition = FindLevel(template, level);
        string pattern = definition?.LevelText?.Val?.Value ?? string.Empty;
        StringBuilder text = new StringBuilder();
        for (int i = 0; i < pattern.Length; i++)
        {
            if (pattern[i] == '%' && i + 1 < pattern.Length && pattern[i + 1] >= '1' && pattern[i + 1] <= '9')
            {
                int number = pattern[i + 1] - '1';
                text.Append(FormatNumber(FindLevel(template, number)?.NumberingFormat?.Val?.Value, value(number)));
                i++;
            }
            else
            {
                text.Append(pattern[i]);
            }
        }
        return text.ToString();
    }

    private static string FormatNumber(NumberFormatValues? format, int number)
    {
        if (format == NumberFormatValues.LowerLetter || format == NumberFormatValues.UpperLetter)
        {
            //Letters repeat once they have run out, as in Word: z, aa, bb.
            char letter = (char)((format == NumberFormatValues.LowerLetter ? 'a' : 'A') + ((number - 1) % 26));
            return new string(letter, ((number - 1) / 26) + 1);
        }
        if (format == NumberFormatValues.LowerRoman)
        {
            return Roman(number).ToLowerInvariant();
        }
        if (format == NumberFormatValues.UpperRoman)
        {
            return Roman(number);
        }
        if (format == NumberFormatValues.DecimalZero)
        {
            return number.ToString("00", CultureInfo.InvariantCulture);
        }
        if (format == NumberFormatValues.Ordinal)
        {
            int tens = number % 100;
            string suffix = tens >= 11 && tens <= 13 ? "th" : (number % 10) == 1 ? "st" : (number % 10) == 2 ? "nd" : (number % 10) == 3 ? "rd" : "th";
            return number.ToString(CultureInfo.InvariantCulture) + suffix;
        }
        if (format == NumberFormatValues.None)
        {
            return string.Empty;
        }
        return number.ToString(CultureInfo.InvariantCulture);
    }

    private static string Roman(int number)
    {
        int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        StringBuilder text = new StringBuilder();
        for (int i = 0; i < values.Length && number > 0; i++)
        {
            while (number >= values[i])
            {
                text.Append(symbols[i]);
                number -= values[i];
            }
        }
        return text.ToString();
    }

    internal static ListType TypeOf(AbstractNum template)
    {
        Level first = template.Elements<Level>().FirstOrDefault();
        if (first?.NumberingFormat?.Val?.Value == NumberFormatValues.Bullet)
        {
            return ListType.Bullet;
        }
        return template.MultiLevelType?.Val?.Value == MultiLevelValues.Multilevel ? ListType.OutlineNumbering : ListType.SimpleNumbering;
    }
}

/// <summary>
/// The list formatting of the paragraphs of a range, as <c>ListFormat</c> in Word. Values that describe one item (<see cref="ListString"/>,
/// <see cref="ListLevelNumber"/>, <see cref="ListValue"/>) are those of the first paragraph.
/// </summary>
public partial class ListFormatWrapper
{
    private readonly RangeWrapper _range;

    internal ListFormatWrapper(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the formatting is of.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<Paragraph> GetParagraphs()
    {
        return _range.Paragraphs.Select(p => p.Paragraph).ToList();
    }

    //The numbered paragraphs of the range, with their numbers.
    internal List<ListNumbering.Item> GetItems(out int paragraphCount)
    {
        List<Paragraph> paragraphs = GetParagraphs();
        paragraphCount = paragraphs.Count;
        if (paragraphs.Count == 0)
        {
            return new List<ListNumbering.Item>();
        }
        HashSet<Paragraph> wanted = new HashSet<Paragraph>(paragraphs);
        return ListNumbering.GetItems(StoryMap.FindRoot(paragraphs[0])).Where(i => wanted.Contains(i.Paragraph)).ToList();
    }

    private List<ListNumbering.Item> GetItems()
    {
        return GetItems(out _);
    }

    /// <summary>
    /// The kind of list the paragraphs are in, as <c>ListFormat.ListType</c> in Word: <see cref="ListType.NoNumbering"/> if none is numbered,
    /// <see cref="ListType.MixedNumbering"/> if they differ.
    /// </summary>
    public ListType ListType
    {
        get
        {
            List<ListNumbering.Item> items = GetItems(out int count);
            if (items.Count == 0)
            {
                return ListType.NoNumbering;
            }
            List<ListType> types = items.Select(i => ListNumbering.TypeOf(i.Template)).Distinct().ToList();
            return types.Count == 1 && items.Count == count ? types[0] : ListType.MixedNumbering;
        }
    }

    /// <summary>
    /// The level of the first paragraph, from 1, as <c>ListFormat.ListLevelNumber</c> in Word; 1 when it is not in a list. Setting it moves
    /// the numbered paragraphs to that level.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The list's template has no such level.</exception>
    public int ListLevelNumber
    {
        get
        {
            ListNumbering.Item first = GetItems().FirstOrDefault();
            return first == null ? 1 : first.Level + 1;
        }
        set { SetListLevelNumber(value); }
    }

    /// <summary>The text of the number or bullet of the first paragraph, as <c>ListFormat.ListString</c> in Word; empty if it has none.</summary>
    public string ListString
    {
        get
        {
            ListNumbering.Item first = GetItems().FirstOrDefault();
            return first == null ? string.Empty : first.Text;
        }
    }

    /// <summary>The number of the first paragraph in its level, as <c>ListFormat.ListValue</c> in Word; 0 when it is not in a list.</summary>
    public int ListValue
    {
        get
        {
            ListNumbering.Item first = GetItems().FirstOrDefault();
            return first == null ? 0 : first.Value;
        }
    }

    /// <summary>The number of numbered paragraphs in the range, as <c>ListFormat.CountNumberedItems</c> in Word.</summary>
    public int CountNumberedItems()
    {
        return GetItems().Count;
    }

    /// <summary>The template of the first numbered paragraph, as <c>ListFormat.ListTemplate</c> in Word; null if none is numbered.</summary>
    public ListTemplateWrapper ListTemplate
    {
        get
        {
            ListNumbering.Item first = GetItems().FirstOrDefault();
            return first == null ? null : new ListTemplateWrapper(first.Template);
        }
    }

    /// <summary>True if the numbered paragraphs all use one template, as <c>ListFormat.SingleListTemplate</c> in Word.</summary>
    public bool SingleListTemplate
    {
        get
        {
            List<ListNumbering.Item> items = GetItems();
            return items.Count > 0 && items.Select(i => i.Template).Distinct().Count() == 1;
        }
    }

    /// <summary>True if the numbered paragraphs are all in one list, as <c>ListFormat.SingleList</c> in Word.</summary>
    public bool SingleList
    {
        get
        {
            List<ListNumbering.Item> items = GetItems();
            return items.Count > 0 && items.Select(i => i.Num).Distinct().Count() == 1;
        }
    }

    /// <summary>The list the first numbered paragraph is in, as <c>ListFormat.List</c> in Word; null if none is.</summary>
    public ListWrapper List
    {
        get
        {
            ListNumbering.Item first = GetItems().FirstOrDefault();
            return first == null ? null : new ListWrapper(first.Num, first.Paragraph);
        }
    }
}

/// <summary>
/// A list, as <c>List</c> in Word: the paragraphs that share one numbering instance (a <c>w:num</c>).
/// </summary>
public class ListWrapper : IEquatable<ListWrapper>
{
    private readonly NumberingInstance _num;
    private readonly Paragraph _anyParagraph;

    internal ListWrapper(NumberingInstance num, Paragraph anyParagraph)
    {
        _num = num;
        _anyParagraph = anyParagraph;
    }

    private List<ListNumbering.Item> GetItems()
    {
        return ListNumbering.GetItems(StoryMap.FindRoot(_anyParagraph)).Where(i => ReferenceEquals(i.Num, _num)).ToList();
    }

    /// <summary>The paragraphs of the list, as <c>List.ListParagraphs</c> in Word.</summary>
    public ListParagraphWrapperCollection ListParagraphs
    {
        get { return new ListParagraphWrapperCollection(GetItems().Select(i => i.Paragraph).ToList()); }
    }

    /// <summary>The number of items in the list, as <c>List.CountNumberedItems</c> in Word.</summary>
    public int CountNumberedItems()
    {
        return GetItems().Count;
    }

    /// <summary>The name of the paragraph style linked to the list, as <c>List.StyleName</c> in Word; empty if there is none.</summary>
    public string StyleName
    {
        get { return string.Empty; }
    }

    /// <summary>True if the whole list uses one template, as <c>List.SingleListTemplate</c> in Word; it always does here.</summary>
    public bool SingleListTemplate
    {
        get { return true; }
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as ListWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(ListWrapper other)
    {
        return !(other is null) && ReferenceEquals(_num, other._num);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _num.GetHashCode();
    }

    public static bool operator ==(ListWrapper left, ListWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(ListWrapper left, ListWrapper right)
    {
        return !(left == right);
    }
}

/// <summary>
/// The lists of the main text of a document, as <c>Lists</c> in Word, in the order they first appear.
/// </summary>
public class ListWrapperCollection : IEnumerable<ListWrapper>
{
    private readonly WordprocessingDocumentWrapper _document;

    internal ListWrapperCollection(WordprocessingDocumentWrapper document)
    {
        _document = document;
    }

    /// <summary>The document the lists are in.</summary>
    public WordprocessingDocumentWrapper Parent
    {
        get { return _document; }
    }

    private List<ListWrapper> GetLists()
    {
        List<ListNumbering.Item> items = ListNumbering.GetItems(_document.WordprocessingDocument.MainDocumentPart.Document.Body);
        return items.GroupBy(i => i.Num).Select(g => new ListWrapper(g.Key, g.First().Paragraph)).ToList();
    }

    /// <summary>The number of lists.</summary>
    public int Count
    {
        get { return GetLists().Count; }
    }

    /// <summary>The list at the index (0-based), as <c>Lists.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no list at the index.</exception>
    public ListWrapper this[int index]
    {
        get
        {
            List<ListWrapper> lists = GetLists();
            if (index < 0 || index >= lists.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return lists[index];
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ListWrapper> GetEnumerator()
    {
        return GetLists().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// The paragraphs in a list or in all the lists of a document, as <c>ListParagraphs</c> in Word.
/// </summary>
public class ListParagraphWrapperCollection : IEnumerable<ParagraphWrapper>
{
    private readonly List<Paragraph> _paragraphs;

    internal ListParagraphWrapperCollection(List<Paragraph> paragraphs)
    {
        _paragraphs = paragraphs;
    }

    /// <summary>The number of paragraphs.</summary>
    public int Count
    {
        get { return _paragraphs.Count; }
    }

    /// <summary>The paragraph at the index (0-based).</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no paragraph at the index.</exception>
    public ParagraphWrapper this[int index]
    {
        get
        {
            if (index < 0 || index >= _paragraphs.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new ParagraphWrapper(_paragraphs[index]);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<ParagraphWrapper> GetEnumerator()
    {
        return _paragraphs.Select(p => new ParagraphWrapper(p)).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
