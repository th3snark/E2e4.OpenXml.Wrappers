using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What <see cref="FindWrapper.Execute(ReplaceMode)"/> does when it finds the text. The values are those of Word's WdReplace.
/// </summary>
public enum ReplaceMode
{
    /// <summary>Only look for the text.</summary>
    None = 0,

    /// <summary>Replace the first match.</summary>
    One = 1,

    /// <summary>Replace every match.</summary>
    All = 2,
}

/// <summary>
/// The replacement text of a find, as <c>Find.Replacement</c> in Word.
/// </summary>
public class ReplacementWrapper
{
    public string Text { get; set; }
}

/// <summary>
/// Finds, and optionally replaces, text in a range, as <c>Range.Find</c> does in Word.
/// </summary>
/// <remarks>
/// <para>When the text is found the range becomes the span of characters that matched (and for <see cref="ReplaceMode.One"/> the replacement
/// text), so it can be formatted or read, and the next <c>Execute</c> carries on from the end of it up to the end of the range that the search
/// started in, as in Word. For <see cref="ReplaceMode.All"/> the range keeps its start and covers the same region with the replacements in it.
/// A range that is collapsed is searched to the end of the story.</para>
/// <para>A match can span paragraphs. A replacement takes the formatting of the first character of the match. The text to find and the
/// replacement can use <c>\t</c> or <c>^t</c> for a tab, <c>\v</c>, <c>\n</c> or <c>^l</c> for a line break, <c>\f</c> or <c>^m</c> for a page break,
/// <c>\r</c> or <c>^p</c> for a paragraph mark and <c>^^</c> for a caret. A replacement containing a paragraph mark starts a new paragraph.</para>
/// </remarks>
public class FindWrapper
{
    private static readonly System.Buffers.SearchValues<char> ControlCharacters = System.Buffers.SearchValues.Create("\t\n\r\v\f");

    private readonly RangeWrapper _range;

    //The state that lets the next Execute carry on from the last match: where that match was, and where the search is to stop.
    private int _lastStart = -1;
    private int _lastEnd = -1;
    private int _limit;

    public FindWrapper(RangeWrapper range)
    {
        ArgumentNullException.ThrowIfNull(range);
        _range = range;
    }

    /// <summary>The text to find.</summary>
    public string Text { get; set; }

    /// <summary>The text to replace the matches with. Null replaces them with nothing.</summary>
    public ReplacementWrapper Replacement { get; } = new ReplacementWrapper();

    /// <summary>True to match the case of the letters of <see cref="Text"/>. The default is a case insensitive match.</summary>
    public bool MatchCase { get; set; }

    /// <summary>True to match only whole words, that is text not directly preceded or followed by a letter or digit.</summary>
    public bool MatchWholeWord { get; set; }

    /// <summary>True if the last <c>Execute</c> found the text.</summary>
    public bool Found { get; private set; }

    /// <summary>
    /// Finds, or replaces, the text using the properties of this object.
    /// </summary>
    /// <returns>True if the text was found.</returns>
    public bool Execute(ReplaceMode replace = ReplaceMode.None)
    {
        Found = false;
        if (String.IsNullOrEmpty(Text))
        {
            return false;
        }

        string find = Unescape(Text);
        string replacement = Unescape(Replacement.Text ?? String.Empty);

        _range.GetPositions(out StoryMap map, out int start, out int end);
        int from;
        int to;
        if (start == _lastStart && end == _lastEnd)
        {
            //Carry on after the last match, as far as the search was meant to go.
            from = end;
            to = Math.Min(_limit, map.Length);
        }
        else
        {
            from = start;
            to = start == end ? map.Length : end;
            _limit = to;
        }

        string haystack = new string(map.ToCharacters());
        StringComparison comparison = MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        List<int> matches = FindMatches(haystack, find, from, to, comparison, replace == ReplaceMode.All);
        if (matches.Count == 0)
        {
            return false;
        }
        Found = true;

        if (replace == ReplaceMode.None)
        {
            Remember(matches[0], matches[0] + find.Length);
            return true;
        }

        if (replace == ReplaceMode.One)
        {
            int length = Replace(map, matches[0], find.Length, replacement);
            Remember(matches[0], matches[0] + length);
            _limit += length - find.Length;
            return true;
        }

        //Replace from the end so that the positions of the matches that are still to be replaced stay valid.
        int delta = 0;
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            int length = Replace(map, matches[i], find.Length, replacement);
            delta += length - find.Length;
        }
        _range.SetRange(start, Math.Max(start, to + delta));
        _lastStart = -1;
        _lastEnd = -1;
        return true;
    }

    /// <summary>
    /// Finds, or replaces, text, setting the properties of this object as Word does.
    /// </summary>
    /// <returns>True if the text was found.</returns>
    public bool Execute(string findText, bool matchCase = false, bool matchWholeWord = false, string replaceWith = null, ReplaceMode replace = ReplaceMode.None)
    {
        Text = findText;
        MatchCase = matchCase;
        MatchWholeWord = matchWholeWord;
        Replacement.Text = replaceWith;
        return Execute(replace);
    }

    //Replaces the match. A match inside one piece of text, replaced by plain text, is done in place; anything else goes through the range.
    //The map is the one made before any replacement; replacing from the end of the story means the text before the match is as it was.
    private int Replace(StoryMap map, int start, int length, string replacement)
    {
        int index = map.IndexAt(start);
        if (index >= 0 && map.Atoms[index].Kind == AtomKind.Text && start + length <= map.Atoms[index].End
            && replacement.AsSpan().IndexOfAny(ControlCharacters) < 0 && replacement.Length > 0)
        {
            Atom atom = map.Atoms[index];
            DocumentFormat.OpenXml.Wordprocessing.Text text = (DocumentFormat.OpenXml.Wordprocessing.Text)atom.Element;
            string current = text.Text ?? String.Empty;
            int offset = start - atom.Start;
            text.Text = String.Concat(current.AsSpan(0, offset), replacement, current.AsSpan(offset + length));
            text.Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve;
            return replacement.Length;
        }
        return _range.CreateSibling(start, start + length).ReplaceSpanText(replacement);
    }

    private void Remember(int start, int end)
    {
        _range.SetRange(start, end);
        _lastStart = start;
        _lastEnd = end;
    }

    //The codes that Word accepts for characters that cannot be typed, and the escapes of C#.
    private static string Unescape(string text)
    {
        return text.Replace("^^", "\u0001").Replace("^p", "\r").Replace("^t", "\t").Replace("^l", "\v").Replace("^m", "\f").Replace("\u0001", "^").Replace('\n', '\v');
    }

    private List<int> FindMatches(string haystack, string find, int from, int to, StringComparison comparison, bool all)
    {
        List<int> matches = new List<int>();
        int position = from;
        while (position <= to - find.Length)
        {
            int index = haystack.IndexOf(find, position, to - position, comparison);
            if (index < 0)
            {
                break;
            }

            if (!MatchWholeWord || IsWholeWord(haystack, index, find.Length))
            {
                matches.Add(index);
                position = index + find.Length;
                if (!all)
                {
                    break;
                }
            }
            else
            {
                position = index + 1;
            }
        }
        return matches;
    }

    private static bool IsWholeWord(string haystack, int index, int length)
    {
        //Only the edges of the match that are themselves part of a word need a word boundary.
        bool startOk = !Char.IsLetterOrDigit(haystack[index]) || index == 0 || !Char.IsLetterOrDigit(haystack[index - 1]);
        int end = index + length;
        bool endOk = !Char.IsLetterOrDigit(haystack[end - 1]) || end == haystack.Length || !Char.IsLetterOrDigit(haystack[end]);
        return startOk && endOk;
    }
}
