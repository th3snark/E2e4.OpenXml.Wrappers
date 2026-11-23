using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// SECTION and NOTEREF fields give what Word gives (#91). The document is that of tools/WordReference/Measure-NoteRefs.ps1, and
/// TestFiles/NoteRefs.txt holds what Word showed for each field in it, with each set of options of the notes.
/// </summary>
public class FieldNoteReferenceTest
{
    private const string Results = @"TestFiles\NoteRefs.txt";

    //The paragraphs and the note at the end of each (F a footnote, E an endnote, * a footnote with the custom mark "*"). A section break
    //(next page) follows HostB and a continuous one follows HostD.
    private static readonly (string Text, string Note)[] Paragraphs =
    {
        ("HostA", ""), ("Alpha", "F"), ("Beta", "F"), ("Gamma", "E"), ("HostB", ""), ("Delta", "F"), ("HostC", ""), ("Epsilon", "E"),
        ("Plain", ""), ("HostD", ""), ("Zeta", "*"), ("Eta", "F"), ("HostE", ""),
    };

    //Footnotes then endnotes: number style, starting number, numbering rule, and for endnotes the location.
    private static readonly Dictionary<string, (NoteNumberStyle Style, int Start, NumberingRule Rule, NoteNumberStyle EndStyle, int EndStart, NumberingRule EndRule, EndnoteLocation EndLocation)> Options =
        new Dictionary<string, (NoteNumberStyle, int, NumberingRule, NoteNumberStyle, int, NumberingRule, EndnoteLocation)>
        {
            ["default"] = (NoteNumberStyle.Arabic, 1, NumberingRule.Continuous, NoteNumberStyle.LowercaseRoman, 1, NumberingRule.Continuous, EndnoteLocation.EndOfDocument),
            ["roman"] = (NoteNumberStyle.UppercaseRoman, 3, NumberingRule.Continuous, NoteNumberStyle.Arabic, 1, NumberingRule.Continuous, EndnoteLocation.EndOfDocument),
            ["symbol"] = (NoteNumberStyle.Symbol, 1, NumberingRule.Continuous, NoteNumberStyle.LowercaseLetter, 1, NumberingRule.RestartSection, EndnoteLocation.EndOfSection),
            ["restart"] = (NoteNumberStyle.Arabic, 1, NumberingRule.RestartSection, NoteNumberStyle.LowercaseRoman, 5, NumberingRule.Continuous, EndnoteLocation.EndOfDocument),
            ["page"] = (NoteNumberStyle.Arabic, 1, NumberingRule.RestartPage, NoteNumberStyle.LowercaseRoman, 1, NumberingRule.Continuous, EndnoteLocation.EndOfDocument),
        };

    [Fact]
    public void SectionAndNoteReferencesGiveWhatWordGives()
    {
        List<string[]> rows = File.ReadAllLines(Results).Where(l => l.Length > 0 && l[0] != '#').Select(l => l.Split('\t')).ToList();
        List<string> differences = new List<string>();
        foreach (IGrouping<string, string[]> options in rows.GroupBy(r => r[0]))
        {
            using WordprocessingDocumentWrapper document = Build(options.Key, out FootnoteWrapper delta);
            foreach (string[] row in options)
            {
                RangeWrapper at;
                if (row[1] == "Note")
                {
                    at = delta.Range;
                }
                else
                {
                    at = TextOf(document, row[1]);
                }
                at.Collapse(CollapseDirection.End);
                FieldWrapper field = document.Fields.Add(at, FieldType.Empty, row[2], false);
                field.Result.Text = "?";
                bool updated = field.Update();

                //Footnotes numbered again on each page depend on the layout, and are left for Word; a custom mark does not.
                bool left = options.Key == "page" && row[2].StartsWith("NOTEREF Fn", StringComparison.Ordinal) && !row[2].StartsWith("NOTEREF FnStar", StringComparison.Ordinal);
                string expected = left ? "?" : row[3];
                bool error = row[3].StartsWith("Error!", StringComparison.Ordinal) && !row[3].Contains("represented");
                string style = ExpectedStyle(row[2], row[3]);
                string actualStyle = field.Result.Text.Length == 0 ? null : field.Result.Range.Characters[0].CharacterStyle;
                if (field.Result.Text != expected || updated != (!left && !error) || (!left && style != null && actualStyle != style))
                {
                    differences.Add(string.Format(CultureInfo.InvariantCulture, "{0} from {1} {2}: expected '{3}' ({4}), library '{5}' ({6}, updated {7})",
                        options.Key, row[1], row[2], expected, style, field.Result.Text, actualStyle, updated));
                }
                field.Delete();
            }
        }

        Assert.True(rows.Count > 900);
        Assert.True(differences.Count == 0, differences.Count.ToString(CultureInfo.InvariantCulture) + " differences:" + Environment.NewLine + string.Join(Environment.NewLine, differences.Take(60)));
    }

    [Fact]
    public void SectionAndNoteReferenceOfTheIssueAreWorkedOut()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        ParagraphWrapper p = document.Paragraphs[0];
        p.Range.Text = "A claim.";
        FootnoteWrapper note = document.Footnotes.Add(p.Range, null, "A note.");
        document.Bookmarks.Add("NoteMark", note.Reference);
        List<FieldWrapper> fields = new List<FieldWrapper>();
        foreach (string code in new[] { "SECTION", "NOTEREF NoteMark" })
        {
            ParagraphWrapper q = document.Paragraphs.Add();
            q.Range.Text = code + ": ";
            FieldWrapper field = document.Fields.Add(q.Range, code);
            field.Result.Text = "?";
            fields.Add(field);
        }

        document.Fields.Update();

        Assert.Equal(new[] { "1", "1" }, fields.Select(f => f.Result.Text));
    }

    //Measured in Word 16: the number formats of \* (and \#) see an endnote's number as negative, whatever the number style of the endnotes.
    [Theory]
    [InlineData("NOTEREF En1 \\* Arabic", "-1")]
    [InlineData("NOTEREF En2 \\* Arabic", "-2")]
    [InlineData("NOTEREF En2 \\* Ordinal", "-2nd")]
    [InlineData("NOTEREF En2 \\* Hex", "Error! Number cannot be represented in specified format.")]
    [InlineData("NOTEREF En2 \\* Upper \\* Arabic", "-2")]
    [InlineData("NOTEREF En2 \\* Arabic \\* Upper", "-2")]
    [InlineData("NOTEREF En2 \\* Upper", "II")]
    [InlineData("NOTEREF En2", "ii")]
    [InlineData("NOTEREF Fn2 \\* Arabic", "2")]
    [InlineData("NOTEREF Fn2 \\* Ordinal", "2nd")]
    public void NumberFormatsOfANoteReferenceGiveWhatWordGives(string code, string expected)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = TwoOfEach(stream);

        FieldWrapper field = document.Fields.Add(document.Paragraphs[2].Range, FieldType.Empty, code, false);

        Assert.True(field.Update());
        Assert.Equal(expected, field.Result.Text);
    }

    //What Word writes for \f: the number in the style of a note's mark, and the position after it in the formatting of the field.
    [Fact]
    public void NoteReferenceWithFGivesTheNumberTheStyleOfAMark()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = TwoOfEach(stream);

        FieldWrapper field = document.Fields.Add(document.Paragraphs[2].Range, FieldType.Empty, "NOTEREF Fn1 \\p \\f", false);
        Assert.True(field.Update());

        Assert.Equal("1 above", field.Result.Text);
        Assert.Equal("Footnote Reference", document.Range(field.Result.Range.Start, field.Result.Range.Start + 1).CharacterStyle);
        Assert.Equal("Default Paragraph Font", document.Range(field.Result.Range.Start + 1, field.Result.Range.End).CharacterStyle);
        Assert.NotNull(document.Styles["Footnote Reference"]);
    }

    //Left for Word: \# (not measured), \p with \*, SECTION in a header (which depends on the page), and a number style other than Word's six.
    [Theory]
    [InlineData("NOTEREF Fn2 \\# \"0.0\"")]
    [InlineData("NOTEREF Fn2 \\p \\* Upper")]
    [InlineData("NOTEREF")]
    public void NoteReferencesNotWorkedOutAreLeftForWord(string code)
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = TwoOfEach(stream);
        FieldWrapper field = document.Fields.Add(document.Paragraphs[2].Range, FieldType.Empty, code, false);
        field.Result.Text = "?";

        Assert.False(field.Update());
        Assert.Equal("?", field.Result.Text);
    }

    [Fact]
    public void SectionInAHeaderIsLeftForWord()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        HeaderFooterWrapper header = document.Sections[0].Headers[HeaderFooterIndex.Primary];
        header.Range.Text = "Section ";
        FieldWrapper field = document.Fields.Add(header.Range, FieldType.Empty, "SECTION", false);
        field.Result.Text = "?";

        Assert.False(field.Update());
        Assert.Equal("?", field.Result.Text);
    }

    //Paragraphs A, B and Host; A and B each have a footnote (Fn1, Fn2) and an endnote (En1, En2) bookmarked.
    private static WordprocessingDocumentWrapper TwoOfEach(Stream stream)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "A\rB\rHost";
        for (int i = 0; i < 2; i++)
        {
            string n = (i + 1).ToString(CultureInfo.InvariantCulture);
            RangeWrapper text = document.Range(document.Paragraphs[i].Range.Start, document.Paragraphs[i].Range.End - 1);
            EndnoteWrapper endnote = document.Endnotes.Add(text, null, "E" + n);
            document.Bookmarks.Add("En" + n, endnote.Reference);
            text = document.Range(document.Paragraphs[i].Range.Start, document.Paragraphs[i].Range.End - 1);
            FootnoteWrapper footnote = document.Footnotes.Add(text, null, "F" + n);
            document.Bookmarks.Add("Fn" + n, footnote.Reference);
        }
        return document;
    }

    //What Word showed for \f: the style of a note's mark on the number (of a footnote or an endnote, as the bookmark's name says), unless a
    //number format of \* writes the number again.
    private static string ExpectedStyle(string code, string result)
    {
        if (!code.Contains("\\f") || code.Contains("\\*") || result.StartsWith("Error!", StringComparison.Ordinal))
        {
            return null;
        }
        return code.Contains("NOTEREF En") ? "Endnote Reference" : "Footnote Reference";
    }

    private static RangeWrapper TextOf(WordprocessingDocumentWrapper document, string text)
    {
        ParagraphWrapper paragraph = document.Paragraphs.First(p => p.Range.Text.StartsWith(text, StringComparison.Ordinal));
        return document.Range(paragraph.Range.Start, paragraph.Range.End - 1);
    }

    private static WordprocessingDocumentWrapper Build(string optionsName, out FootnoteWrapper delta)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(new MemoryStream(), true, false);
        document.Content.Text = string.Join("\r", Paragraphs.Select(p => p.Text));

        Dictionary<string, object> notes = new Dictionary<string, object>();
        foreach ((string text, string note) in Paragraphs.Where(p => p.Note.Length > 0))
        {
            RangeWrapper at = TextOf(document, text);
            string body = "Note " + text;
            notes[text] = note == "E" ? document.Endnotes.Add(at, null, body) : document.Footnotes.Add(at, note == "*" ? "*" : null, body);
        }
        foreach (string host in new[] { "HostB", "HostD" })
        {
            RangeWrapper at = TextOf(document, host);
            at.Collapse(CollapseDirection.End);
            at.InsertBreak(host == "HostB" ? BreakType.SectionBreakNextPage : BreakType.SectionBreakContinuous);
        }

        RangeWrapper Mark(string text) => notes[text] is EndnoteWrapper endnote ? endnote.Reference : ((FootnoteWrapper)notes[text]).Reference;
        document.Bookmarks.Add("FnA", Mark("Alpha"));
        document.Bookmarks.Add("FnB", TextOf(document, "Beta"));
        document.Bookmarks.Add("EnA", Mark("Gamma"));
        document.Bookmarks.Add("FnC", Mark("Delta"));
        document.Bookmarks.Add("EnB", Mark("Epsilon"));
        document.Bookmarks.Add("NoNote", TextOf(document, "Plain"));
        document.Bookmarks.Add("FnStar", Mark("Zeta"));
        document.Bookmarks.Add("FnD", Mark("Eta"));
        document.Bookmarks.Add("Two", document.Range(TextOf(document, "Zeta").Start, TextOf(document, "Eta").End));

        (NoteNumberStyle Style, int Start, NumberingRule Rule, NoteNumberStyle EndStyle, int EndStart, NumberingRule EndRule, EndnoteLocation EndLocation) o = Options[optionsName];
        document.Footnotes.NumberingRule = o.Rule;
        document.Footnotes.NumberStyle = o.Style;
        document.Footnotes.StartingNumber = o.Start;
        document.Endnotes.Location = o.EndLocation;
        document.Endnotes.NumberingRule = o.EndRule;
        document.Endnotes.NumberStyle = o.EndStyle;
        document.Endnotes.StartingNumber = o.EndStart;

        delta = (FootnoteWrapper)notes["Delta"];
        return document;
    }
}
