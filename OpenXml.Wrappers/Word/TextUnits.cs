using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The units that a range can be moved and expanded by, as <c>WdUnits</c> in Word.
/// </summary>
public enum TextUnit
{
    Character = 1,
    Word = 2,
    Sentence = 3,
    Paragraph = 4
}

/// <summary>
/// The end of a range that it is collapsed to, as <c>WdCollapseDirection</c> in Word.
/// </summary>
public enum CollapseDirection
{
    End = 0,
    Start = 1
}

/// <summary>
/// Finds where the characters, words, sentences and paragraphs of a text begin.
/// </summary>
internal static class TextUnits
{
    /// <summary>
    /// The positions at which a unit begins, in order, followed by the length of the text so that each unit runs from one entry to the next.
    /// A word includes the spaces after it; a sentence ends at a full stop, question mark or exclamation mark and the spaces after it, or at
    /// the end of a paragraph. The end of a paragraph and the end of a table row are units of their own in words.
    /// </summary>
    public static List<int> GetBoundaries(char[] text, TextUnit unit)
    {
        if (unit == TextUnit.Character || Array.IndexOf(text, '\0') < 0)
        {
            return GetBoundariesOfVisible(text, unit);
        }

        //Positions that show no text (field codes, pictures) do not make words or sentences of their own: they go with the unit that follows.
        List<int> positions = new List<int>();
        for (int i = 0; i < text.Length; i++)
        {
            if (text[i] != '\0')
            {
                positions.Add(i);
            }
        }
        char[] visible = positions.Select(p => text[p]).ToArray();
        return GetBoundariesOfVisible(visible, unit)
            .Select(b => b >= visible.Length ? text.Length : b == 0 ? 0 : positions[b - 1] + 1)
            .ToList();
    }

    private static List<int> GetBoundariesOfVisible(char[] text, TextUnit unit)
    {
        List<int> starts = new List<int>();
        int n = text.Length;
        switch (unit)
        {
            case TextUnit.Character:
                for (int i = 0; i < n; i++)
                {
                    starts.Add(i);
                }
                break;
            case TextUnit.Word:
                AddWords(text, starts);
                break;
            case TextUnit.Sentence:
                AddSentences(text, starts);
                break;
            case TextUnit.Paragraph:
                AddParagraphs(text, starts);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(unit));
        }
        starts.Add(n);
        return starts;
    }

    private static bool IsMark(char c)
    {
        return c == '\r' || c == '\a';
    }

    private static bool IsSpace(char c)
    {
        return c == ' ' || c == '\t' || c == ' ';
    }

    private static void AddParagraphs(char[] text, List<int> starts)
    {
        int start = 0;
        for (int i = 0; i < text.Length; i++)
        {
            if (IsMark(text[i]))
            {
                starts.Add(start);
                start = i + 1;
            }
        }
        if (start < text.Length)
        {
            starts.Add(start);
        }
    }

    private static void AddWords(char[] text, List<int> starts)
    {
        int i = 0;
        int n = text.Length;
        while (i < n)
        {
            starts.Add(i);
            char c = text[i];
            if (IsMark(c))
            {
                i++;
                continue;
            }
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                i++;
                while (i < n && (char.IsLetterOrDigit(text[i]) || text[i] == '_' || ((text[i] == '\'' || text[i] == '’') && i + 1 < n && char.IsLetterOrDigit(text[i + 1]))))
                {
                    i++;
                }
            }
            else if (IsSpace(c))
            {
                while (i < n && IsSpace(text[i]))
                {
                    i++;
                }
                continue;
            }
            else
            {
                i++;
            }

            //The spaces after a word belong to it.
            while (i < n && IsSpace(text[i]))
            {
                i++;
            }
        }
    }

    private static void AddSentences(char[] text, List<int> starts)
    {
        int i = 0;
        int n = text.Length;
        while (i < n)
        {
            starts.Add(i);
            while (i < n)
            {
                char c = text[i];
                if (IsMark(c))
                {
                    i++;
                    break;
                }
                i++;
                if (c == '.' || c == '!' || c == '?')
                {
                    //Closing punctuation after the end of the sentence belongs to it.
                    while (i < n && (text[i] == '.' || text[i] == '!' || text[i] == '?' || text[i] == '"' || text[i] == ')' || text[i] == '”' || text[i] == '\''))
                    {
                        i++;
                    }
                    if (i >= n || IsSpace(text[i]) || IsMark(text[i]))
                    {
                        while (i < n && IsSpace(text[i]))
                        {
                            i++;
                        }
                        break;
                    }
                }
            }
        }
    }
}

/// <summary>
/// The characters, words or sentences of a range, as <c>Range.Characters</c>, <c>Words</c> and <c>Sentences</c> in Word. Each item is a range.
/// A word or sentence that is only partly inside the range is included whole, as in Word. Indexes are 0-based.
/// </summary>
public class TextUnitCollection : IEnumerable<RangeWrapper>
{
    private readonly RangeWrapper _range;
    private readonly TextUnit _unit;

    internal TextUnitCollection(RangeWrapper range, TextUnit unit)
    {
        _range = range;
        _unit = unit;
    }

    /// <summary>
    /// The range the items are taken from.
    /// </summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    public RangeWrapper this[int index]
    {
        get
        {
            List<(int Start, int End)> units = GetUnits(out StoryMap map);
            if (index < 0 || index >= units.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return RangeWrapper.FromMap(map, units[index].Start, units[index].End);
        }
    }

    //The units that the range touches, as positions in the map that is returned: the map of the paragraph if the range is inside one.
    private List<(int Start, int End)> GetUnits(out StoryMap map)
    {
        _range.GetScoped(false, out map, out int start, out int end);
        if (_unit == TextUnit.Character)
        {
            List<(int, int)> characters = new List<(int, int)>();
            for (int i = start; i < end; i++)
            {
                characters.Add((i, i + 1));
            }
            return characters;
        }

        List<int> boundaries = TextUnits.GetBoundaries(map.ToCharacters(), _unit);
        List<(int, int)> units = new List<(int, int)>();
        for (int i = 0; i + 1 < boundaries.Count; i++)
        {
            int unitStart = boundaries[i];
            int unitEnd = boundaries[i + 1];
            bool overlaps = start == end ? (unitStart <= start && start < unitEnd) : (unitStart < end && unitEnd > start);
            if (overlaps)
            {
                units.Add((unitStart, unitEnd));
            }
        }
        return units;
    }

    public int Count
    {
        get { return GetUnits(out _).Count; }
    }

    public IEnumerator<RangeWrapper> GetEnumerator()
    {
        //The units are worked out once, for the whole enumeration, so that editing the items does not disturb it.
        List<(int Start, int End)> units = GetUnits(out StoryMap map);
        List<RangeWrapper> ranges = units.Select(u => RangeWrapper.FromMap(map, u.Start, u.End)).ToList();
        foreach (RangeWrapper range in ranges)
        {
            yield return range;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
