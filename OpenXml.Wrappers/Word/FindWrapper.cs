using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

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
/// <para>A match can span several runs but not several paragraphs. A replacement takes the formatting of the first run of the match.</para>
/// <para>Unlike in Word, a find never changes the range it was called on: ranges are made of whole elements so they cannot be narrowed to
/// the match. <see cref="Execute(ReplaceMode)"/> returns whether the text was found and <see cref="Found"/> remembers that.</para>
/// <para>Tabs and line breaks in the text to find match tab and break characters (<c>\t</c>, <c>\v</c> and <c>\f</c> for a page break). The
/// replacement text must not contain them (or <c>\r</c> or <c>\n</c>) because it replaces text within a run.</para>
/// </remarks>
public class FindWrapper
{
    private static readonly System.Buffers.SearchValues<char> ControlCharacters = System.Buffers.SearchValues.Create("\t\n\r\v\f");

    private readonly RangeWrapper _range;

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

        string replacement = Replacement.Text ?? String.Empty;
        if (replace != ReplaceMode.None && replacement.AsSpan().IndexOfAny(ControlCharacters) >= 0)
        {
            throw new NotSupportedException("Replacement text containing tabs, line breaks or paragraph marks is not currently supported.");
        }

        StringComparison comparison = MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        foreach (List<Atom> atoms in GetParagraphAtoms())
        {
            string haystack = String.Concat(atoms.Select(a => a.Content));
            List<int> matches = FindMatches(haystack, comparison);
            if (matches.Count == 0)
            {
                continue;
            }

            Found = true;
            if (replace == ReplaceMode.None)
            {
                return true;
            }

            if (replace == ReplaceMode.One)
            {
                ReplaceMatch(atoms, matches[0], Text.Length, replacement);
                return true;
            }

            //Replace from the end so that the offsets of the matches that are still to be replaced stay valid.
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                ReplaceMatch(atoms, matches[i], Text.Length, replacement);
            }
        }
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

    private List<int> FindMatches(string haystack, StringComparison comparison)
    {
        List<int> matches = new List<int>();
        int position = 0;
        while (position <= haystack.Length - Text.Length)
        {
            int index = haystack.IndexOf(Text, position, comparison);
            if (index < 0)
            {
                break;
            }

            if (!MatchWholeWord || IsWholeWord(haystack, index, Text.Length))
            {
                matches.Add(index);
                position = index + Text.Length;
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

    /// <summary>
    /// A piece of the text of a paragraph and the element that holds it: a text element, or a tab or break character.
    /// </summary>
    private sealed class Atom
    {
        public OpenXmlElement Element { get; set; }

        public string Content { get; set; }
    }

    private IEnumerable<List<Atom>> GetParagraphAtoms()
    {
        foreach (List<Run> runs in _range.GetRunsByParagraph())
        {
            List<Atom> atoms = new List<Atom>();
            foreach (Run run in runs)
            {
                foreach (OpenXmlElement element in run.ChildElements)
                {
                    if (element is Text)
                    {
                        atoms.Add(new Atom { Element = element, Content = ((Text)element).Text ?? String.Empty });
                    }
                    else if (element is TabChar)
                    {
                        atoms.Add(new Atom { Element = element, Content = "\t" });
                    }
                    else if (element is Break)
                    {
                        Break lineBreak = (Break)element;
                        bool isPageBreak = lineBreak.Type != null && lineBreak.Type.HasValue && lineBreak.Type.Value == BreakValues.Page;
                        atoms.Add(new Atom { Element = element, Content = isPageBreak ? "\f" : "\v" });
                    }
                }
            }
            yield return atoms;
        }
    }

    private static void ReplaceMatch(List<Atom> atoms, int matchStart, int matchLength, string replacement)
    {
        int matchEnd = matchStart + matchLength;
        List<Run> touchedRuns = new List<Run>();
        bool replaced = false;

        int atomStart = 0;
        foreach (Atom atom in atoms)
        {
            int atomEnd = atomStart + atom.Content.Length;
            if (atomEnd > matchStart && atomStart < matchEnd)
            {
                Run run = atom.Element.Parent as Run;
                if (run != null)
                {
                    touchedRuns.Add(run);
                }

                //The part of this atom before and after the match is kept.
                string before = atom.Content.Substring(0, Math.Max(0, matchStart - atomStart));
                string after = atom.Content.Substring(Math.Min(atom.Content.Length, matchEnd - atomStart));
                string kept = (replaced ? String.Empty : before + replacement) + after;

                if (atom.Element is Text)
                {
                    Text text = (Text)atom.Element;
                    text.Text = kept;
                    text.Space = SpaceProcessingModeValues.Preserve;
                    if (kept.Length == 0)
                    {
                        text.Remove();
                    }
                }
                else if (kept.Length > 0)
                {
                    //A tab or break at the edge of the match is replaced by the text.
                    Text newText = new Text(kept) { Space = SpaceProcessingModeValues.Preserve };
                    atom.Element.InsertBeforeSelf(newText);
                    atom.Element.Remove();
                    atom.Element = newText;
                }
                else
                {
                    atom.Element.Remove();
                }

                //Keep the atom in step with the document so that the matches still to be replaced (all of which are earlier) see the new text.
                atom.Content = kept;
                replaced = true;
            }
            atomStart = atomEnd;
        }

        //Runs left with nothing but their properties are removed.
        foreach (Run run in touchedRuns.Distinct())
        {
            if (run.Parent != null && !run.ChildElements.Any(e => !(e is RunProperties)))
            {
                run.Remove();
            }
        }
    }
}
