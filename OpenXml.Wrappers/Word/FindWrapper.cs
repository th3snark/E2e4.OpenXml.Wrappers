using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

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
/// What a find does when it reaches the end of the range it searches, as <c>WdFindWrap</c> in Word (with the same values).
/// </summary>
public enum FindWrap
{
    /// <summary>Stop at the end of the range.</summary>
    Stop = 0,

    /// <summary>Carry on through the rest of the story, wrapping round to the start of it.</summary>
    Continue = 1,

    /// <summary>Ask the user. This is not supported, because there is no user to ask.</summary>
    Ask = 2,
}

/// <summary>
/// The replacement text of a find, as <c>Find.Replacement</c> in Word.
/// </summary>
public class ReplacementWrapper
{
    public string Text { get; set; }

    /// <summary>Does nothing: the formatting of the replacement cannot be set, so there is none to clear. It is here so that code written for Word runs.</summary>
    public void ClearFormatting()
    {
    }
}

/// <summary>
/// Finds, and optionally replaces, text in a range, as <c>Range.Find</c> does in Word.
/// </summary>
/// <remarks>
/// <para>When the text is found the range becomes the span of characters that matched (and for <see cref="ReplaceMode.One"/> the replacement
/// text), so it can be formatted or read, and the next <c>Execute</c> carries on from the end of it up to the end of the range that the search
/// started in, as in Word. For <see cref="ReplaceMode.All"/> the range keeps covering the same region, with the replacements in it. A range
/// that is collapsed is searched to the end of the story (or, searching backwards, to the start).</para>
/// <para>A match can span paragraphs. A replacement takes the formatting of the first character of the match. The text to find and the
/// replacement can use <c>\t</c> or <c>^t</c> for a tab, <c>\v</c>, <c>\n</c> or <c>^l</c> for a line break, <c>\f</c> or <c>^m</c> for a page break,
/// <c>\r</c> or <c>^p</c> for a paragraph mark, <c>^n</c> for a column break, <c>^s</c> for a no-break space, <c>^+</c> and <c>^=</c> for an em
/// and en dash, <c>^-</c> for an optional hyphen and <c>^^</c> for a caret. The text to find can use <c>^w</c> for any white space, <c>^#</c>
/// for a digit, <c>^$</c> for a letter and <c>^?</c> for any character; the replacement can use <c>^&amp;</c> for the text that was found. A
/// replacement containing a paragraph mark starts a new paragraph.</para>
/// <para>With <see cref="MatchWildcards"/> the text to find is a Word pattern: <c>?</c> any character, <c>*</c> any string, <c>[abc]</c>,
/// <c>[a-z]</c> and <c>[!abc]</c> sets, <c>&lt;</c> and <c>&gt;</c> the start and end of a word, <c>@</c> one or more of the previous
/// character or group, <c>{n}</c>, <c>{n,}</c>, <c>{n,m}</c> and <c>{,m}</c> repeats, <c>(...)</c> groups that the replacement uses as
/// <c>\1</c> to <c>\9</c> (so can the pattern itself), and <c>\</c> to find a character that has a meaning. Matching the sound or the forms of a
/// word, and searching for formatting, are not supported.</para>
/// <para>A search for text that does not include a paragraph mark goes through the paragraphs one at a time from where it starts, so looking
/// for the next match in a large document costs only as much as the distance to it. A search with <see cref="MatchWildcards"/> looks at the
/// whole story.</para>
/// </remarks>
public class FindWrapper
{
    private static readonly char[] ControlCharacters = new char[] { '\t', '\n', '\r', '\v', '\f' };

    private readonly RangeWrapper _range;

    //The state that lets the next Execute carry on from the last match: a copy of the range as it was after that match, and a range at the point
    //where the search is to stop. They are ranges so that they follow the text, as the range itself does, when the text is formatted or changed.
    private RangeWrapper _lastMatch;
    private RangeWrapper _limit;

    //What the Execute that is running looks for and replaces with.
    private Pattern _pattern;
    private Replacer _replacer;

    public FindWrapper(RangeWrapper range)
    {
        Guard.ThrowIfNull(range);
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

    /// <summary>True to treat <see cref="Text"/> as a pattern, as described for the class.</summary>
    public bool MatchWildcards { get; set; }

    /// <summary>True (the default) to search towards the end of the range, false to search back towards the start.</summary>
    public bool Forward { get; set; } = true;

    /// <summary>What to do at the end of the range: stop, or carry on through the rest of the story.</summary>
    public FindWrap Wrap { get; set; }

    /// <summary>Matching words that sound alike is not supported: setting this to true throws <see cref="NotSupportedException"/>.</summary>
    public bool MatchSoundsLike
    {
        get { return false; }
        set { ThrowIfSet(value, nameof(MatchSoundsLike)); }
    }

    /// <summary>Matching the forms of a word is not supported: setting this to true throws <see cref="NotSupportedException"/>.</summary>
    public bool MatchAllWordForms
    {
        get { return false; }
        set { ThrowIfSet(value, nameof(MatchAllWordForms)); }
    }

    /// <summary>Searching for formatting is not supported: setting this to true throws <see cref="NotSupportedException"/>.</summary>
    public bool Format
    {
        get { return false; }
        set { ThrowIfSet(value, nameof(Format)); }
    }

    /// <summary>True if the last <c>Execute</c> found the text.</summary>
    public bool Found { get; private set; }

    /// <summary>Does nothing: formatting cannot be searched for, so there is none to clear. It is here so that code written for Word runs.</summary>
    public void ClearFormatting()
    {
    }

    private static void ThrowIfSet(bool value, string name)
    {
        if (value)
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Find.{0} is not supported.", name));
        }
    }

    /// <summary>
    /// Finds, or replaces, the text using the properties of this object.
    /// </summary>
    /// <returns>True if the text was found.</returns>
    /// <exception cref="ArgumentException">The text is not a valid pattern.</exception>
    public bool Execute(ReplaceMode replace = ReplaceMode.None)
    {
        Found = false;
        if (String.IsNullOrEmpty(Text))
        {
            return false;
        }
        if (Wrap == FindWrap.Ask)
        {
            throw new NotSupportedException("FindWrap.Ask is not supported: there is no user to ask.");
        }

        _pattern = Pattern.Create(Text, MatchWildcards, MatchCase);
        _replacer = new Replacer(Replacement.Text ?? String.Empty, MatchWildcards);

        //Where the search starts and ends. Looking for the same thing again from the match carries on from its end.
        _range.GetScoped(false, out StoryMap scopeMap, out int scopeStart, out int scopeEnd);
        OpenXmlElement story = StoryMap.FindRoot(_range.AnchorElement);
        Anchor start = _range.IsSpan ? _range.StartAnchor : Anchor.At(scopeMap, scopeStart);
        Anchor end = _range.IsSpan ? _range.EndAnchor : Anchor.At(scopeMap, scopeEnd);
        bool continuing = _lastMatch != null && start.SameAs(_lastMatch.StartAnchor) && end.SameAs(_lastMatch.EndAnchor);
        Anchor from;
        Anchor to;
        if (Forward)
        {
            if (continuing)
            {
                from = end;
                to = _limit.StartAnchor;
            }
            else
            {
                from = start;
                to = scopeStart == scopeEnd ? new Anchor(null, 0) : end;
                _limit = RangeWrapper.CreateAt(story, to);
            }
        }
        else if (continuing)
        {
            from = _limit.StartAnchor;
            to = start;
        }
        else
        {
            from = scopeStart == scopeEnd ? GetStoryStart(story) : start;
            to = end;
            _limit = RangeWrapper.CreateAt(story, from);
        }

        bool found = Search(story, scopeMap, scopeStart, scopeEnd, from, to, replace);
        if (!found && Wrap == FindWrap.Continue)
        {
            //The rest of the story, and then the part of it before the start; what was searched already has nothing to find.
            found = Search(story, scopeMap, scopeStart, scopeEnd, GetStoryStart(story), new Anchor(null, 0), replace);
        }
        Found = found;
        return Found;
    }

    /// <summary>
    /// Finds, or replaces, text, setting the properties of this object as Word does (a parameter that is left out leaves its property as it is). The parameters are in the order of Word's
    /// <c>Find.Execute</c>.
    /// </summary>
    /// <returns>True if the text was found.</returns>
    /// <exception cref="NotSupportedException"><paramref name="matchSoundsLike"/>, <paramref name="matchAllWordForms"/> or <paramref name="format"/> is true.</exception>
    public bool Execute(string findText, bool? matchCase = null, bool? matchWholeWord = null, bool? matchWildcards = null, bool? matchSoundsLike = null,
        bool? matchAllWordForms = null, bool? forward = null, FindWrap? wrap = null, bool? format = null, string replaceWith = null, ReplaceMode replace = ReplaceMode.None)
    {
        ThrowIfSet(matchSoundsLike == true, nameof(MatchSoundsLike));
        ThrowIfSet(matchAllWordForms == true, nameof(MatchAllWordForms));
        ThrowIfSet(format == true, nameof(Format));

        //An argument that is left out leaves the property as it is, as in Word.
        Text = findText;
        MatchCase = matchCase ?? MatchCase;
        MatchWholeWord = matchWholeWord ?? MatchWholeWord;
        MatchWildcards = matchWildcards ?? MatchWildcards;
        Forward = forward ?? Forward;
        Wrap = wrap ?? Wrap;
        if (replaceWith != null)
        {
            Replacement.Text = replaceWith;
        }
        return Execute(replace);
    }

    private static Anchor GetStoryStart(OpenXmlElement story)
    {
        return Anchor.At(StoryMap.Build(story), 0);
    }

    private bool Search(OpenXmlElement story, StoryMap scopeMap, int scopeStart, int scopeEnd, Anchor from, Anchor to, ReplaceMode replace)
    {
        //A pattern can match across paragraphs, so it looks at the whole story. So can text with a cell or row marker.
        bool? found = _pattern.Regex != null || _pattern.Literal.Contains('\a', StringComparison.Ordinal)
            ? null
            : _pattern.Literal.Contains('\r', StringComparison.Ordinal)
                ? ExecuteAcrossParagraphs(story, scopeMap, scopeStart, scopeEnd, from, to, replace)
                : ExecuteByParagraph(story, scopeMap, scopeStart, scopeEnd, from, to, replace);
        if (found == null)
        {
            found = ExecuteInStory(story, from, to, replace, scopeMap, scopeStart, scopeEnd);
        }
        return found.Value;
    }

    //Whether every match in the window is needed, to pick the last of them or to replace them all.
    private bool NeedsAll(ReplaceMode replace)
    {
        return replace == ReplaceMode.All || !Forward;
    }

    private FoundMatch Pick(List<FoundMatch> matches)
    {
        return Forward ? matches[0] : matches[matches.Count - 1];
    }

    //Makes the range the match, or replaces it, according to the mode. The map is the one the match was found in.
    private void Apply(StoryMap map, Paragraph paragraph, FoundMatch match, ReplaceMode replace)
    {
        if (replace == ReplaceMode.None)
        {
            Remember(map, match.Start, match.Start + match.Length);
            return;
        }

        int length = Replace(map, match.Start, match.Length, _replacer.For(match));
        Remember(StoryMap.Build(paragraph), match.Start, match.Start + length);
    }

    #region By paragraph

    /// <summary>
    /// Searches the paragraphs one at a time. Returns null when the window cannot be worked out that way (it ends in a table row, say), in
    /// which case nothing has been changed.
    /// </summary>
    private bool? ExecuteByParagraph(OpenXmlElement story, StoryMap scopeMap, int scopeStart, int scopeEnd, Anchor from, Anchor to, ReplaceMode replace)
    {
        Paragraph first = ParagraphOf(from);
        Paragraph last = to.Element == null ? null : ParagraphOf(to);
        if (first == null || (to.Element != null && last == null))
        {
            return null;
        }

        bool any = false;
        bool madeSpan = _range.IsSpan;
        bool keepLast = !Forward && replace != ReplaceMode.All;
        StoryMap bestMap = null;
        Paragraph bestParagraph = null;
        FoundMatch best = null;
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
                List<FoundMatch> matches = FindMatches(map, windowStart, windowEnd, NeedsAll(replace));
                if (matches.Count > 0)
                {
                    any = true;
                    if (keepLast)
                    {
                        //Searching backwards the last match in the window is wanted, so the paragraphs are all looked at.
                        bestMap = map;
                        bestParagraph = paragraph;
                        best = matches[matches.Count - 1];
                    }
                    else if (replace != ReplaceMode.All)
                    {
                        Apply(map, paragraph, matches[0], replace);
                        return true;
                    }
                    else
                    {
                        //A replace all keeps the range as it was, but as a span, which the replacements then keep up to date.
                        if (!madeSpan)
                        {
                            _range.MakeSpan(scopeMap, scopeStart, scopeEnd);
                            madeSpan = true;
                        }
                        for (int i = matches.Count - 1; i >= 0; i--)
                        {
                            Replace(map, matches[i].Start, matches[i].Length, _replacer.For(matches[i]));
                        }
                    }
                }
            }

            if (last != null && paragraph == last)
            {
                break;
            }
        }

        if (best != null)
        {
            Apply(bestMap, bestParagraph, best, replace);
            return true;
        }

        _lastMatch = null;
        return any;
    }

    /// <summary>
    /// Searches for text that includes paragraph marks. A match starts in one paragraph and runs on into the next few, so each paragraph is
    /// searched together with as many of those after it as the text has marks. Returns null when the window cannot be worked out that way.
    /// </summary>
    private bool? ExecuteAcrossParagraphs(OpenXmlElement story, StoryMap scopeMap, int scopeStart, int scopeEnd, Anchor from, Anchor to, ReplaceMode replace)
    {
        int marks = _pattern.Literal.Count(c => c == '\r');
        Paragraph first = ParagraphOf(from);
        Paragraph last = to.Element == null ? null : ParagraphOf(to);
        if (first == null || (to.Element != null && last == null))
        {
            return null;
        }

        bool any = false;
        bool madeSpan = _range.IsSpan;
        bool keepLast = !Forward && replace != ReplaceMode.All;
        StoryMap bestMap = null;
        FoundMatch best = null;
        Paragraph paragraph = first;
        Anchor startAt = from;
        while (paragraph != null)
        {
            Paragraph reach = paragraph;
            for (int i = 0; i < marks && reach != null; i++)
            {
                reach = StoryMap.NextParagraph(reach, story);
            }
            if (reach == null)
            {
                break;
            }

            StoryMap map = StoryMap.BuildWindowFor(paragraph, reach, story);
            if (map == null)
            {
                return any ? any : (bool?)null;
            }

            int windowStart = startAt == null ? 0 : RangeWrapper.ResolveLocal(map, startAt);
            int windowEnd = to.Element == null ? map.Length : RangeWrapper.ResolveLocal(map, to);
            if (windowEnd < 0)
            {
                windowEnd = map.Length;
            }
            if (windowStart < 0 || !map.TryGetAtom(paragraph, out Atom mark))
            {
                return any ? any : (bool?)null;
            }

            //A match belongs to the paragraph it starts in, so that it is found once.
            List<FoundMatch> matches = FindMatches(map, windowStart, windowEnd, keepLast, mark.End);
            if (matches.Count > 0)
            {
                any = true;
                if (keepLast)
                {
                    bestMap = map;
                    best = matches[matches.Count - 1];
                }
                else
                {
                    FoundMatch match = matches[0];
                    if (replace == ReplaceMode.None)
                    {
                        Remember(map, match.Start, match.Start + match.Length);
                        return true;
                    }

                    RangeWrapper replaced = RangeWrapper.FromMap(map, match.Start, match.Start + match.Length);
                    replaced.ReplaceSpanText(_replacer.For(match));
                    if (replace == ReplaceMode.One)
                    {
                        _range.MakeSpanAnchors(replaced.StartAnchor, replaced.EndAnchor);
                        _lastMatch = _range.Duplicate();
                        return true;
                    }

                    //A replace all goes on from the end of the replacement, which can be in the middle of a paragraph now.
                    if (!madeSpan)
                    {
                        _range.MakeSpan(scopeMap, scopeStart, scopeEnd);
                        madeSpan = true;
                    }
                    paragraph = ParagraphOf(replaced.EndAnchor);
                    startAt = replaced.EndAnchor;
                    continue;
                }
            }

            if (last != null && paragraph == last)
            {
                break;
            }
            paragraph = StoryMap.NextParagraph(paragraph, story);
            startAt = null;
        }

        if (best != null)
        {
            if (replace == ReplaceMode.None)
            {
                Remember(bestMap, best.Start, best.Start + best.Length);
            }
            else
            {
                RangeWrapper replaced = RangeWrapper.FromMap(bestMap, best.Start, best.Start + best.Length);
                replaced.ReplaceSpanText(_replacer.For(best));
                _range.MakeSpanAnchors(replaced.StartAnchor, replaced.EndAnchor);
                _lastMatch = _range.Duplicate();
            }
            return true;
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
    private bool ExecuteInStory(OpenXmlElement story, Anchor from, Anchor to, ReplaceMode replace, StoryMap scopeMap, int scopeStart, int scopeEnd)
    {
        StoryMap map = StoryMap.Build(story);
        int fromPosition = RangeWrapper.ResolveLocal(map, from);
        int toPosition = to.Element == null ? map.Length : RangeWrapper.ResolveLocal(map, to);
        fromPosition = Math.Max(fromPosition, 0);
        toPosition = toPosition < 0 ? map.Length : toPosition;

        List<FoundMatch> matches = FindMatches(map, fromPosition, toPosition, NeedsAll(replace));
        if (matches.Count == 0)
        {
            return false;
        }

        if (replace == ReplaceMode.None)
        {
            FoundMatch match = Pick(matches);
            Remember(map, match.Start, match.Start + match.Length);
            return true;
        }

        if (replace == ReplaceMode.One)
        {
            FoundMatch match = Pick(matches);
            int length = Replace(map, match.Start, match.Length, _replacer.For(match));
            Remember(StoryMap.Build(story), match.Start, match.Start + length);
            return true;
        }

        if (!_range.IsSpan)
        {
            _range.MakeSpan(scopeMap, scopeStart, scopeEnd);
        }

        //Replace from the end so that the positions of the matches that are still to be replaced stay valid.
        for (int i = matches.Count - 1; i >= 0; i--)
        {
            Replace(map, matches[i].Start, matches[i].Length, _replacer.For(matches[i]));
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
            && replacement.IndexOfAny(ControlCharacters) < 0 && replacement.Length > 0)
        {
            Atom atom = map.Atoms[index];
            Text text = (Text)atom.Element;
            string current = text.Text ?? String.Empty;
            int offset = start - atom.Start;
            text.Text = String.Concat(current.Substring(0, offset), replacement, current.Substring(offset + length));
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

    //Finds the matches in the positions from windowStart up to windowEnd. Positions that show no text (field codes, pictures) are skipped over,
    //so a match can run across them; each match is given as the position it starts at and the number of positions it takes up.
    private List<FoundMatch> FindMatches(StoryMap map, int windowStart, int windowEnd, bool all, int maxStart = int.MaxValue)
    {
        char[] characters = map.ToCharacters();
        List<int> positions = null;
        string haystack;
        if (Array.IndexOf(characters, '\0') < 0)
        {
            haystack = new string(characters);
        }
        else
        {
            positions = new List<int>(characters.Length);
            StringBuilder visible = new StringBuilder(characters.Length);
            for (int i = 0; i < characters.Length; i++)
            {
                if (characters[i] != '\0')
                {
                    positions.Add(i);
                    visible.Append(characters[i]);
                }
            }
            haystack = visible.ToString();
        }

        //The window in the characters that show.
        int from = positions == null ? windowStart : FirstIndexAtOrAfter(positions, windowStart);
        int to = positions == null ? windowEnd : FirstIndexAtOrAfter(positions, windowEnd);
        int startLimit = positions == null ? maxStart : FirstIndexAtOrAfter(positions, maxStart);

        List<FoundMatch> matches = new List<FoundMatch>();
        int position = from;
        if (_pattern.Regex != null)
        {
            //The pattern sees the text before the window (for the edges of words) but nothing after it.
            string capped = to < haystack.Length ? haystack.Substring(0, to) : haystack;
            while (position <= capped.Length)
            {
                Match match = _pattern.Regex.Match(capped, position);
                if (!match.Success || match.Index >= startLimit)
                {
                    break;
                }
                if (match.Length == 0 || (MatchWholeWord && !IsWholeWord(haystack, match.Index, match.Length)))
                {
                    position = match.Index + 1;
                    continue;
                }

                matches.Add(CreateMatch(positions, match.Index, match.Length, match.Value, match.Groups.Cast<Group>().Select(g => g.Value).ToArray()));
                position = match.Index + match.Length;
                if (!all)
                {
                    break;
                }
            }
            return matches;
        }

        string find = _pattern.Literal;
        StringComparison comparison = _pattern.Comparison;
        while (position <= to - find.Length)
        {
            int index = haystack.IndexOf(find, position, to - position, comparison);
            if (index < 0 || index >= startLimit)
            {
                break;
            }

            if (!MatchWholeWord || IsWholeWord(haystack, index, find.Length))
            {
                matches.Add(CreateMatch(positions, index, find.Length, haystack.Substring(index, find.Length), null));
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

    private static FoundMatch CreateMatch(List<int> positions, int index, int length, string text, string[] groups)
    {
        int start = positions == null ? index : positions[index];
        int end = positions == null ? index + length : positions[index + length - 1] + 1;
        return new FoundMatch(start, end - start, text, groups);
    }

    private static int FirstIndexAtOrAfter(List<int> positions, int position)
    {
        int index = positions.BinarySearch(position);
        return index >= 0 ? index : ~index;
    }

    private static bool IsWholeWord(string haystack, int index, int length)
    {
        //Only the edges of the match that are themselves part of a word need a word boundary.
        bool startOk = !Char.IsLetterOrDigit(haystack[index]) || index == 0 || !Char.IsLetterOrDigit(haystack[index - 1]);
        int end = index + length;
        bool endOk = !Char.IsLetterOrDigit(haystack[end - 1]) || end == haystack.Length || !Char.IsLetterOrDigit(haystack[end]);
        return startOk && endOk;
    }

    #region Matching and replacing text

    //A match: where it is, the text that matched and, for a pattern, what its groups matched (the first is the whole match).
    private sealed class FoundMatch
    {
        public FoundMatch(int start, int length, string text, string[] groups)
        {
            Start = start;
            Length = length;
            Text = text;
            Groups = groups;
        }

        public int Start { get; }

        public int Length { get; }

        public string Text { get; }

        public string[] Groups { get; }
    }

    //The text of a ^ code (as in ^p), or null if there is no such code. ^^ is a caret.
    private static string GetCode(char code)
    {
        switch (code)
        {
            case 'p':
                return "\r";
            case 't':
                return "\t";
            case 'l':
                return "\v";
            case 'm':
                return "\f";
            case 'n':
                return "\u000e";
            case 's':
                return " ";
            case '+':
                return "—";
            case '=':
                return "–";
            case '-':
                return "­";
            case '^':
                return "^";
            default:
                return null;
        }
    }

    //What is looked for: plain text, or a regular expression made from a pattern or from the codes that stand for a kind of character.
    private sealed class Pattern
    {
        public string Literal { get; private set; }

        public Regex Regex { get; private set; }

        public StringComparison Comparison { get; private set; }

        public static Pattern Create(string text, bool wildcards, bool matchCase)
        {
            text = text.Replace('\n', '\v');
            Pattern pattern = new Pattern() { Comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase };

            bool kinds = false;
            for (int i = 0; i + 1 < text.Length; i++)
            {
                if (text[i] == '^')
                {
                    char code = text[i + 1];
                    kinds |= code == 'w' || code == '#' || code == '$' || code == '?';
                    i++;
                }
            }

            if (!wildcards && !kinds)
            {
                pattern.Literal = Unescape(text);
                return pattern;
            }

            try
            {
                pattern.Regex = new Regex(ToRegex(text, wildcards), RegexOptions.CultureInvariant | (matchCase ? RegexOptions.None : RegexOptions.IgnoreCase), TimeSpan.FromSeconds(30));
            }
            catch (ArgumentException ex)
            {
                throw new ArgumentException("The text to find is not a valid pattern: " + ex.Message, nameof(text), ex);
            }
            pattern.Literal = text;
            return pattern;
        }

        private static string Unescape(string text)
        {
            StringBuilder sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                string code = text[i] == '^' && i + 1 < text.Length ? GetCode(text[i + 1]) : null;
                if (code != null)
                {
                    sb.Append(code);
                    i++;
                }
                else
                {
                    sb.Append(text[i]);
                }
            }
            return sb.ToString();
        }

        private static string ToRegex(string text, bool wildcards)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '^' && i + 1 < text.Length)
                {
                    char d = text[i + 1];
                    string code = GetCode(d);
                    if (code != null)
                    {
                        sb.Append(Regex.Escape(code));
                        i++;
                        continue;
                    }
                    if (d == 'w' || d == '#' || d == '$' || d == '?')
                    {
                        sb.Append(d == 'w' ? "[ \\t\\u00A0]+" : d == '#' ? "[0-9]" : d == '$' ? "\\p{L}" : "[\\s\\S]");
                        i++;
                        continue;
                    }
                }

                if (!wildcards)
                {
                    sb.Append(Regex.Escape(c.ToString()));
                    continue;
                }

                switch (c)
                {
                    case '?':
                        sb.Append("[\\s\\S]");
                        break;
                    case '*':
                        sb.Append("[\\s\\S]*?");
                        break;
                    case '@':
                        sb.Append('+');
                        break;
                    case '<':
                        sb.Append("\\b(?=\\w)");
                        break;
                    case '>':
                        sb.Append("(?<=\\w)\\b");
                        break;
                    case '(':
                    case ')':
                        sb.Append(c);
                        break;
                    case '[':
                        int close = text.IndexOf(']', i + 1);
                        if (close < 0)
                        {
                            throw new ArgumentException("A [ has no matching ].");
                        }
                        string set = text.Substring(i + 1, close - i - 1);
                        bool negated = set.Length > 0 && set[0] == '!';
                        if (negated)
                        {
                            set = set.Substring(1);
                        }
                        sb.Append(negated ? "[^" : "[").Append(set.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("^", "\\^", StringComparison.Ordinal).Replace("[", "\\[", StringComparison.Ordinal)).Append(']');
                        i = close;
                        break;
                    case '{':
                        int end = text.IndexOf('}', i + 1);
                        if (end < 0)
                        {
                            throw new ArgumentException("A { has no matching }.");
                        }
                        string repeat = text.Substring(i + 1, end - i - 1);
                        sb.Append('{').Append(repeat.Length > 0 && repeat[0] == ',' ? "0" + repeat : repeat).Append('}');
                        i = end;
                        break;
                    case '\\':
                        if (i + 1 >= text.Length)
                        {
                            throw new ArgumentException("A \\ has nothing after it.");
                        }
                        char next = text[i + 1];
                        sb.Append(next >= '1' && next <= '9' ? "\\" + next : Regex.Escape(next.ToString()));
                        i++;
                        break;
                    default:
                        sb.Append(Regex.Escape(c.ToString()));
                        break;
                }
            }
            return sb.ToString();
        }
    }

    //Works out the replacement for a match: the codes are turned into the characters they stand for, ^& into the text that was found and,
    //with patterns, \1 to \9 into the text of the groups.
    private sealed class Replacer
    {
        private readonly string _template;
        private readonly bool _wildcards;
        private readonly bool _dynamic;
        private readonly string _constant;

        public Replacer(string template, bool wildcards)
        {
            _template = template;
            _wildcards = wildcards;
            _dynamic = template.Contains("^&", StringComparison.Ordinal) || (wildcards && Regex.IsMatch(template, "\\\\[1-9]"));
            _constant = _dynamic ? null : Expand(null);
        }

        public string For(FoundMatch match)
        {
            return _dynamic ? Expand(match) : _constant;
        }

        private string Expand(FoundMatch match)
        {
            StringBuilder sb = new StringBuilder(_template.Length);
            for (int i = 0; i < _template.Length; i++)
            {
                char c = _template[i];
                if (c == '^' && i + 1 < _template.Length)
                {
                    if (_template[i + 1] == '&')
                    {
                        sb.Append(match?.Text);
                        i++;
                        continue;
                    }
                    string code = GetCode(_template[i + 1]);
                    if (code != null)
                    {
                        sb.Append(code);
                        i++;
                        continue;
                    }
                }
                else if (c == '\\' && _wildcards && i + 1 < _template.Length)
                {
                    char next = _template[i + 1];
                    if (next >= '1' && next <= '9')
                    {
                        int group = next - '0';
                        if (match?.Groups != null && group < match.Groups.Length)
                        {
                            sb.Append(match.Groups[group]);
                        }
                    }
                    else
                    {
                        sb.Append(next);
                    }
                    i++;
                    continue;
                }

                sb.Append(c == '\n' ? '\v' : c);
            }
            return sb.ToString();
        }
    }

    #endregion
}
