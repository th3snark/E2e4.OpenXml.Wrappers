using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The case of text, as <c>WdCharacterCase</c> in Word.
/// </summary>
public enum CharacterCase
{
    /// <summary>All lower case.</summary>
    LowerCase = 0,

    /// <summary>All upper case.</summary>
    UpperCase = 1,

    /// <summary>The first letter of each word in upper case.</summary>
    TitleWord = 2,

    /// <summary>The first letter of each sentence in upper case. Reading it, text that is neither of the others is <see cref="ToggleCase"/>.</summary>
    TitleSentence = 4,

    /// <summary>Each letter in the opposite case. It is also what a range of text in mixed case reads as.</summary>
    ToggleCase = 5
}

/// <summary>
/// Which character the text of an alignment tab lines up on, as <c>WdAlignmentTabAlignment</c> in Word.
/// </summary>
public enum AlignmentTabAlignment
{
    Left = 0,
    Center = 1,
    Right = 2
}

/// <summary>
/// What the position of an alignment tab is measured from, as <c>WdAlignmentTabRelative</c> in Word.
/// </summary>
public enum AlignmentTabRelative
{
    /// <summary>The margins.</summary>
    Margin = 0,

    /// <summary>The indents of the paragraph.</summary>
    Indent = 1
}

/// <summary>
/// Inserting symbols, alignment tabs and the date or time, and changing the case of text.
/// </summary>
public partial class RangeWrapper
{
    /// <summary>
    /// The case of the text of the range, as <c>Range.Case</c> in Word: <see cref="CharacterCase.LowerCase"/> if there is no capital letter,
    /// <see cref="CharacterCase.UpperCase"/> if there is no small letter, <see cref="CharacterCase.TitleWord"/> or
    /// <see cref="CharacterCase.TitleSentence"/> if the text is already like that, otherwise <see cref="CharacterCase.ToggleCase"/>. Setting
    /// it changes the letters in place, so the formatting of the text is kept. Text in capitals that is mixed with other text, such as an
    /// acronym, is kept in capitals, as in Word.
    /// </summary>
    public CharacterCase Case
    {
        get
        {
            GetScoped(false, out StoryMap map, out int start, out int end);
            return CaseRules.Read(new string(map.ToCharacters().Skip(start).Take(end - start).ToArray()));
        }
        set
        {
            if (value != CharacterCase.LowerCase && value != CharacterCase.UpperCase && value != CharacterCase.TitleWord && value != CharacterCase.TitleSentence && value != CharacterCase.ToggleCase)
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }

            GetPositions(out StoryMap map, out int start, out int end);
            char[] characters = map.ToCharacters().Skip(start).Take(end - start).ToArray();
            string changed = CaseRules.Apply(new string(characters), value);
            foreach (Atom atom in map.Atoms)
            {
                if (atom.Kind != AtomKind.Text || !(atom.Element is Text text) || atom.End <= start || atom.Start >= end)
                {
                    continue;
                }
                int from = Math.Max(start, atom.Start);
                int to = Math.Min(end, atom.End);
                string current = text.Text;
                text.Text = new StringBuilder().Append(current, 0, from - atom.Start).Append(changed, from - start, to - from).Append(current, to - atom.Start, current.Length - (to - atom.Start)).ToString();
            }
        }
    }

    /// <summary>
    /// Replaces the range with a symbol, as <c>Range.InsertSymbol</c> in Word. Collapse the range first to keep its text.
    /// </summary>
    /// <param name="characterNumber">The code of the character: a Unicode character if <paramref name="unicode"/> is true, otherwise the
    /// number the character has in the font's table of symbols (31 and the place in the table, as in Word).</param>
    /// <param name="font">The name of the font that has the symbol. Without one the character is inserted in the font of the text around it.</param>
    /// <param name="unicode">True when <paramref name="characterNumber"/> is a Unicode character.</param>
    /// <remarks>
    /// A symbol of a Unicode character is written as the character in a run with that font, and the other kind as a <c>w:sym</c> element,
    /// which is how Word writes each. A <c>w:sym</c> takes one position and shows no text.
    /// </remarks>
    public void InsertSymbol(int characterNumber, string font = null, bool unicode = false)
    {
        if (characterNumber < 0 || characterNumber > 0x10FFFF || (characterNumber >= 0xD800 && characterNumber <= 0xDFFF))
        {
            throw new ArgumentOutOfRangeException(nameof(characterNumber));
        }
        bool asText = unicode || string.IsNullOrEmpty(font);
        if (!asText && characterNumber > 0xFF)
        {
            throw new ArgumentOutOfRangeException(nameof(characterNumber), "A symbol that is not Unicode has a number from 0 to 255.");
        }

        ReplaceWithRuns((paragraph, formatting) =>
        {
            Run run = new Run();
            RunProperties properties = formatting == null ? null : (RunProperties)formatting.CloneNode(true);
            if (asText)
            {
                if (!string.IsNullOrEmpty(font))
                {
                    //The complex script font is left out when the text has that font anyway, which Word does too.
                    RunFonts fonts = new RunFonts() { Ascii = font, HighAnsi = font };
                    if (!string.Equals(EffectiveComplexScriptFont(paragraph, properties), font, StringComparison.Ordinal))
                    {
                        fonts.ComplexScript = font;
                    }
                    properties = properties ?? new RunProperties();
                    properties.RunFonts = fonts;
                }
                if (properties != null && properties.HasChildren)
                {
                    run.AppendChild(properties);
                }
                run.AppendChild(OpenXmlExtensions.NewText(char.ConvertFromUtf32(unicode ? characterNumber : characterNumber)));
            }
            else
            {
                if (properties != null && properties.HasChildren)
                {
                    run.AppendChild(properties);
                }
                run.AppendChild(new SymbolChar() { Font = font, Char = (0xF000 + characterNumber).ToString("X4", CultureInfo.InvariantCulture) });
            }
            return new List<Run> { run };
        });
    }

    /// <summary>
    /// Replaces the range with an alignment tab, an absolute tab that is always in the same place, as <c>Range.InsertAlignmentTab</c> in Word.
    /// It takes one position and shows no text here.
    /// </summary>
    /// <param name="alignment">Left, centre or right.</param>
    /// <param name="relativeTo">Whether the place is measured from the margins or the indents.</param>
    public void InsertAlignmentTab(AlignmentTabAlignment alignment, AlignmentTabRelative relativeTo = AlignmentTabRelative.Margin)
    {
        if (!EnumCompat.IsDefined(alignment))
        {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }
        if (!EnumCompat.IsDefined(relativeTo))
        {
            throw new ArgumentOutOfRangeException(nameof(relativeTo));
        }

        PositionalTab tab = new PositionalTab()
        {
            RelativeTo = relativeTo == AlignmentTabRelative.Margin ? AbsolutePositionTabPositioningBaseValues.Margin : AbsolutePositionTabPositioningBaseValues.Indent,
            Alignment = alignment == AlignmentTabAlignment.Left ? AbsolutePositionTabAlignmentValues.Left : alignment == AlignmentTabAlignment.Center ? AbsolutePositionTabAlignmentValues.Center : AbsolutePositionTabAlignmentValues.Right,
            Leader = AbsolutePositionTabLeaderCharValues.None
        };
        ReplaceWithRuns((paragraph, formatting) => new List<Run> { WithFormatting(new Run((OpenXmlElement)tab), formatting) });
    }

    /// <summary>
    /// Replaces the range with the current date or time, as <c>Range.InsertDateTime</c> in Word: as a <c>DATE</c> field (which Word writes
    /// for it) or as text.
    /// </summary>
    /// <param name="format">The format, in Word's letters (<c>dd/MM/yyyy</c>, <c>d MMMM yyyy</c>, <c>HH:mm</c>, <c>h:mm am/pm</c>). The short date format of the current culture when omitted.</param>
    /// <param name="insertAsField">True (the default, as in Word) for a field that shows the date when it was inserted and is updated by Word.</param>
    public void InsertDateTime(string format = null, bool insertAsField = true)
    {
        string picture = string.IsNullOrEmpty(format) ? CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern : format;
        string shown = DateTime.Now.ToString(ToNetFormat(picture), CultureInfo.CurrentCulture);
        if (!insertAsField)
        {
            Text = shown;
            return;
        }

        ReplaceWithRuns((paragraph, formatting) =>
        {
            RunProperties resultProperties = formatting == null ? new RunProperties() : (RunProperties)formatting.CloneNode(true);
            resultProperties.NoProof = new NoProof();
            return new List<Run>
            {
                WithFormatting(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }), formatting),
                WithFormatting(new Run(new FieldCode(" DATE \\@ \"" + picture + "\" ") { Space = SpaceProcessingModeValues.Preserve }), formatting),
                WithFormatting(new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }), formatting),
                new Run(resultProperties, OpenXmlExtensions.NewText(shown)),
                WithFormatting(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }), formatting)
            };
        });
    }

    //Puts the mark of a footnote or endnote at the place of a collapsed range.
    internal void InsertNoteReference(Func<Paragraph, List<Run>> build)
    {
        ReplaceWithRuns((paragraph, formatting) => build(paragraph));
    }

    //Replaces the text of a span with runs made for the place, which take the formatting of the first character replaced (or of the character
    //before an insertion point); the formatting is passed to the builder, which puts it in the run. A range of elements is not replaced.
    private void ReplaceWithRuns(Func<Paragraph, RunProperties, List<Run>> build)
    {
        if (!_isSpan)
        {
            InsertRuns(paragraph => build(paragraph, null), true, null);
            return;
        }

        GetScoped(false, out StoryMap map, out int start, out int end);
        RunProperties formatting = FindFormatting(map, start, end);
        formatting = formatting == null ? null : (RunProperties)formatting.CloneNode(true);
        if (end > start)
        {
            Delete();
        }
        InsertRuns(paragraph => build(paragraph, formatting), true, null);
    }

    //The font of complex script text in a place: the one the text has directly, else the one its paragraph style or the defaults give.
    private static string EffectiveComplexScriptFont(Paragraph paragraph, RunProperties direct)
    {
        FormattingResolver resolver = new FormattingResolver(paragraph);
        RunFonts fonts = direct?.RunFonts;
        if (fonts?.ComplexScript?.Value != null && fonts.ComplexScriptTheme == null)
        {
            return fonts.ComplexScript.Value;
        }
        if (fonts?.ComplexScriptTheme != null && fonts.ComplexScriptTheme.HasValue)
        {
            return resolver.GetThemeComplexScriptFontName(fonts.ComplexScriptTheme.Value);
        }

        RunFonts resolved = resolver.ResolveParagraphMark(paragraph)?.RunFonts;
        if (resolved?.ComplexScriptTheme != null && resolved.ComplexScriptTheme.HasValue)
        {
            return resolver.GetThemeComplexScriptFontName(resolved.ComplexScriptTheme.Value);
        }
        return resolved?.ComplexScript?.Value;
    }

    //Word's letters for a date or time picture as .NET's: am/pm is tt, and a single letter is a custom format rather than a standard one.
    private static string ToNetFormat(string picture)
    {
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < picture.Length; i++)
        {
            char c = picture[i];
            if (c == '\'' || c == '"')
            {
                int close = picture.IndexOf(c, i + 1);
                int stop = close < 0 ? picture.Length - 1 : close;
                result.Append('\'').Append(picture, i + 1, stop - i - 1).Append('\'');
                i = stop;
            }
            else if (string.Compare(picture, i, "am/pm", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
            {
                result.Append("tt");
                i += 4;
            }
            else if (string.Compare(picture, i, "a/p", 0, 3, StringComparison.OrdinalIgnoreCase) == 0)
            {
                result.Append('t');
                i += 2;
            }
            else
            {
                result.Append(c);
            }
        }
        return result.Length == 1 ? "%" + result : result.ToString();
    }
}

/// <summary>
/// Word's rules for changing the case of text. A word is a run of letters and apostrophes. Text with no small letter is first made small. A
/// word in capitals (of two letters or more) or a capital followed by small letters is kept as it is, and any other word (<c>wORLD</c>,
/// <c>McDonald</c>) is made small, before the first letter of each word or sentence is made a capital.
/// </summary>
internal static class CaseRules
{
    public static string Apply(string text, CharacterCase value)
    {
        if (value == CharacterCase.LowerCase)
        {
            return Map(text, char.ToLowerInvariant);
        }
        if (value == CharacterCase.UpperCase)
        {
            return Map(text, char.ToUpperInvariant);
        }
        if (value == CharacterCase.ToggleCase)
        {
            return Map(text, c => char.IsUpper(c) ? char.ToLowerInvariant(c) : char.IsLower(c) ? char.ToUpperInvariant(c) : c);
        }

        char[] characters = (text.Any(char.IsLower) ? text : Map(text, char.ToLowerInvariant)).ToCharArray();
        foreach ((int start, int length) word in Words(characters))
        {
            if (!Keeps(characters, word.start, word.length))
            {
                for (int i = word.start; i < word.start + word.length; i++)
                {
                    characters[i] = char.ToLowerInvariant(characters[i]);
                }
            }
        }

        if (value == CharacterCase.TitleWord)
        {
            foreach ((int start, int length) word in Words(characters))
            {
                Capitalize(characters, word.start);
            }
        }
        else
        {
            bool atStart = true;
            for (int i = 0; i < characters.Length; i++)
            {
                if (atStart && char.IsLetter(characters[i]))
                {
                    Capitalize(characters, i);
                    atStart = false;
                }
                else if (characters[i] == '\r' || characters[i] == '\a')
                {
                    atStart = true;
                }
                else if ((characters[i] == '.' || characters[i] == '!' || characters[i] == '?') && (i + 1 >= characters.Length || char.IsWhiteSpace(characters[i + 1]) || characters[i + 1] == '\a'))
                {
                    atStart = true;
                }
            }
        }
        return new string(characters);
    }

    /// <summary>
    /// The case a text is in: lower if it has no capital, upper if it has no small letter, a title if each word is capitalised or the way
    /// Word would make sentences of it, and otherwise toggled (mixed).
    /// </summary>
    public static CharacterCase Read(string text)
    {
        if (!text.Any(char.IsUpper))
        {
            return CharacterCase.LowerCase;
        }
        if (!text.Any(char.IsLower))
        {
            return CharacterCase.UpperCase;
        }

        List<(int start, int length)> words = Words(text.ToCharArray()).ToList();
        if (words.All(w => char.IsUpper(text[w.start])) && Apply(text, CharacterCase.TitleWord) == text)
        {
            return CharacterCase.TitleWord;
        }

        //A sentence has a capital at its start and none after it.
        if (Apply(text, CharacterCase.TitleSentence) == text)
        {
            bool atStart = true;
            bool sentences = true;
            foreach ((int start, int length) word in words)
            {
                bool capital = char.IsUpper(text[word.start]);
                bool moreCapitals = text.Skip(word.start + 1).Take(word.length - 1).Any(char.IsUpper);
                if (atStart ? !capital : capital || moreCapitals)
                {
                    sentences = false;
                    break;
                }
                atStart = StartsSentenceAfter(text, word.start + word.length);
            }
            if (sentences)
            {
                return CharacterCase.TitleSentence;
            }
        }
        return CharacterCase.ToggleCase;
    }

    private static bool StartsSentenceAfter(string text, int position)
    {
        for (int i = position; i < text.Length; i++)
        {
            if (text[i] == '\r' || text[i] == '\a')
            {
                return true;
            }
            if (text[i] == '.' || text[i] == '!' || text[i] == '?')
            {
                return i + 1 >= text.Length || char.IsWhiteSpace(text[i + 1]) || text[i + 1] == '\a';
            }
            if (char.IsLetter(text[i]))
            {
                return false;
            }
        }
        return false;
    }

    private static string Map(string text, Func<char, char> change)
    {
        char[] characters = text.ToCharArray();
        for (int i = 0; i < characters.Length; i++)
        {
            characters[i] = change(characters[i]);
        }
        return new string(characters);
    }

    private static void Capitalize(char[] characters, int index)
    {
        characters[index] = char.ToUpperInvariant(characters[index]);
    }

    private static IEnumerable<(int start, int length)> Words(char[] characters)
    {
        int i = 0;
        while (i < characters.Length)
        {
            if (char.IsLetter(characters[i]))
            {
                int start = i;
                while (i < characters.Length && (char.IsLetter(characters[i]) || ((characters[i] == '\'' || characters[i] == '’') && i + 1 < characters.Length && char.IsLetter(characters[i + 1]))))
                {
                    i++;
                }
                yield return (start, i - start);
            }
            else
            {
                i++;
            }
        }
    }

    //Whether a word is left as it is: one capital letter or more followed by none, a capital followed by small letters, or a single letter.
    private static bool Keeps(char[] characters, int start, int length)
    {
        IEnumerable<char> letters = characters.Skip(start).Take(length).Where(char.IsLetter).ToList();
        if (letters.Count() < 2)
        {
            return true;
        }
        return letters.All(char.IsUpper) || (char.IsUpper(letters.First()) && letters.Skip(1).All(char.IsLower)) || letters.All(char.IsLower);
    }
}
