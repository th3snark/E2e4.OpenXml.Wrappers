using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

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
/// started in, as in Word. For <see cref="ReplaceMode.All"/> the range keeps covering the same region, with the replacements in it. A range
/// that is collapsed is searched to the end of the story.</para>
/// <para>A match can span paragraphs. A replacement takes the formatting of the first character of the match. The text to find and the
/// replacement can use <c>\t</c> or <c>^t</c> for a tab, <c>\v</c>, <c>\n</c> or <c>^l</c> for a line break, <c>\f</c> or <c>^m</c> for a page break,
/// <c>\r</c> or <c>^p</c> for a paragraph mark and <c>^^</c> for a caret. A replacement containing a paragraph mark starts a new paragraph.</para>
/// <para>A search for text that does not include a paragraph mark goes through the paragraphs one at a time from where it starts, so looking
/// for the next match in a large document costs only as much as the distance to it.</para>
/// </remarks>
public class FindWrapper
{
    private static readonly System.Buffers.SearchValues<char> ControlCharacters = System.Buffers.SearchValues.Create("\t\n\r\v\f");

    private readonly RangeWrapper _range;

    //The state that lets the next Execute carry on from the last match: a copy of the range as it was after that match, and a range at the point
    //where the search is to stop. They are ranges so that they follow the text, as the range itself does, when the text is formatted or changed.
    private RangeWrapper _lastMatch;
    private RangeWrapper _limit;

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
        StringComparison comparison = MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

        //Where the search starts and ends. Looking for the same thing again from the match carries on from its end.
        _range.GetScoped(false, out StoryMap scopeMap, out int scopeStart, out int scopeEnd);
        OpenXmlElement story = StoryMap.FindRoot(_range.AnchorElement);
        Anchor start = _range.IsSpan ? _range.StartAnchor.Clone() : Anchor.At(scopeMap, scopeStart);
        Anchor end = _range.IsSpan ? _range.EndAnchor.Clone() : Anchor.At(scopeMap, scopeEnd);
        bool continuing = _lastMatch != null && start.SameAs(_lastMatch.StartAnchor) && end.SameAs(_lastMatch.EndAnchor);
        Anchor from;
        Anchor to;
        if (continuing)
        {
            from = end;
            to = _limit.StartAnchor.Clone();
        }
        else
        {
            from = start;
            to = scopeStart == scopeEnd ? new Anchor(null, 0) : end;
            _limit = RangeWrapper.CreateAt(story, to);
        }

        bool? found = find.Contains('\r', StringComparison.Ordinal) || find.Contains('\a', StringComparison.Ordinal)
            ? null
            : ExecuteByParagraph(story, scopeMap, scopeStart, scopeEnd, from, to, find, replacement, replace, comparison);
        if (found == null)
        {
            found = ExecuteInStory(story, from, to, find, replacement, replace, comparison, scopeMap, scopeStart, scopeEnd);
        }
        Found = found.Value;
        return Found;
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

    #region By paragraph

    /// <summary>
    /// Searches the paragraphs one at a time. Returns null when the window cannot be worked out that way (it ends in a table row, say), in
    /// which case nothing has been changed.
    /// </summary>
    private bool? ExecuteByParagraph(OpenXmlElement story, StoryMap scopeMap, int scopeStart, int scopeEnd, Anchor from, Anchor to, string find, string replacement, ReplaceMode replace, StringComparison comparison)
    {
        Paragraph first = ParagraphOf(from);
        Paragraph last = to.Element == null ? null : ParagraphOf(to);
        if (first == null || (to.Element != null && last == null))
        {
            return null;
        }

        bool any = false;
        bool madeSpan = _range.IsSpan;
        for (Paragraph paragraph = first; paragraph != null; paragraph = StoryMap.NextParagraph(paragraph, story))
        {
            StoryMap map = StoryMap.Build(paragraph);
            int windowStart = paragraph == first ? RangeWrapper.ResolveLocal(map, from) : 0;
            int windowEnd = last != null && paragraph == last ? RangeWrapper.ResolveLocal(map, to) : map.Length - 1;
            if (windowStart < 0)
            {
                if (any)
                {
                    break;
                }
                return null;
            }

            //The mark that ends the paragraph is not text that can match.
            windowEnd = Math.Min(windowEnd < 0 ? map.Length - 1 : windowEnd, map.Length - 1);
            if (windowEnd > windowStart)
            {
                List<int> matches = FindMatches(new string(map.ToCharacters()), find, windowStart, windowEnd, comparison, replace == ReplaceMode.All);
                if (matches.Count > 0)
                {
                    any = true;
                    if (replace == ReplaceMode.None)
                    {
                        Remember(map, matches[0], matches[0] + find.Length);
                        return true;
                    }
                    if (replace == ReplaceMode.One)
                    {
                        int length = Replace(map, matches[0], find.Length, replacement);
                        Remember(StoryMap.Build(paragraph), matches[0], matches[0] + length);
                        return true;
                    }

                    //A replace all keeps the range as it was, but as a span, which the replacements then keep up to date.
                    if (!madeSpan)
                    {
                        _range.MakeSpan(scopeMap, scopeStart, scopeEnd);
                        madeSpan = true;
                    }
                    for (int i = matches.Count - 1; i >= 0; i--)
                    {
                        Replace(map, matches[i], find.Length, replacement);
                    }
                }
            }

            if (last != null && paragraph == last)
            {
                break;
            }
        }

        _lastMatch = null;
        return any;
    }

    private static Paragraph ParagraphOf(Anchor anchor)
    {
        if (anchor.Element == null || anchor.Element is TableRow)
        {
            return null;
        }
        return anchor.Element as Paragraph ?? anchor.Element.Ancestors<Paragraph>().FirstOrDefault();
    }

    #endregion

    #region By story

    //Searches the whole story at once, which a match across paragraphs needs.
    private bool ExecuteInStory(OpenXmlElement story, Anchor from, Anchor to, string find, string replacement, ReplaceMode replace, StringComparison comparison, StoryMap scopeMap, int scopeStart, int scopeEnd)
    {
        StoryMap map = StoryMap.Build(story);
        int fromPosition = RangeWrapper.ResolveLocal(map, from);
        int toPosition = to.Element == null ? map.Length : RangeWrapper.ResolveLocal(map, to);
        fromPosition = Math.Max(fromPosition, 0);
        toPosition = toPosition < 0 ? map.Length : toPosition;

        List<int> matches = FindMatches(new string(map.ToCharacters()), find, fromPosition, toPosition, comparison, replace == ReplaceMode.All);
        if (matches.Count == 0)
        {
            return false;
        }

        if (replace == ReplaceMode.None)
        {
            Remember(map, matches[0], matches[0] + find.Length);
            return true;
        }

        if (replace == ReplaceMode.One)
        {
            int length = Replace(map, matches[0], find.Length, replacement);
            Remember(StoryMap.Build(story), matches[0], matches[0] + length);
            return true;
        }

        if (!_range.IsSpan)
        {
            _range.MakeSpan(scopeMap, scopeStart, scopeEnd);
        }

        //Replace from the end so that the positions of the matches that are still to be replaced stay valid.
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            Replace(map, matches[i], find.Length, replacement);
        }
        _lastMatch = null;
        return true;
    }

    #endregion

    //Replaces the match. A match inside one piece of text, replaced by plain text, is done in place; anything else goes through the range.
    //The map is the one made before any replacement; replacing from the end means the text before the match is as it was.
    private static int Replace(StoryMap map, int start, int length, string replacement)
    {
        int index = map.IndexAt(start);
        if (index >= 0 && map.Atoms[index].Kind == AtomKind.Text && start + length <= map.Atoms[index].End
            && replacement.AsSpan().IndexOfAny(ControlCharacters) < 0 && replacement.Length > 0)
        {
            Atom atom = map.Atoms[index];
            Text text = (Text)atom.Element;
            string current = text.Text ?? String.Empty;
            int offset = start - atom.Start;
            text.Text = String.Concat(current.AsSpan(0, offset), replacement, current.AsSpan(offset + length));
            text.Space = SpaceProcessingModeValues.Preserve;
            SpanRegistry.TextEdited(text, offset, length, replacement.Length);
            return replacement.Length;
        }
        return RangeWrapper.FromMap(map, start, start + length).ReplaceSpanText(replacement);
    }

    //Makes the range the match and remembers it, so that the next find carries on from it.
    private void Remember(StoryMap map, int start, int end)
    {
        _range.MakeSpan(map, start, end);
        _lastMatch = _range.Duplicate();
    }

    //The codes that Word accepts for characters that cannot be typed, and the escapes of C#.
    private static string Unescape(string text)
    {
        return text.Replace("^^", "\u0001", StringComparison.Ordinal).Replace("^p", "\r", StringComparison.Ordinal).Replace("^t", "\t", StringComparison.Ordinal)
            .Replace("^l", "\v", StringComparison.Ordinal).Replace("^m", "\f", StringComparison.Ordinal).Replace("\u0001", "^", StringComparison.Ordinal)
            .Replace('\n', '\v');
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
