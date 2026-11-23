using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Indexes (#96): <c>Indexes.Add</c>, <c>MarkEntry</c> and <c>Index.Update</c>, with page numbers given by the caller or estimated. The markup
/// of an index is compared with Word's in MarkupConformanceTest (the index-* scenarios); these check the rest against what Word 16 gave,
/// measured over COM.
/// </summary>
public class IndexTest
{
    //A document of paragraphs, each with the XE fields given, a page break after each paragraph but the last, and an empty paragraph at the end.
    private static WordprocessingDocumentWrapper Create(Stream stream, params string[][] pages)
    {
        WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        for (int page = 0; page < pages.Length; page++)
        {
            document.Content.InsertAfter("Page ");
            foreach (string code in pages[page])
            {
                document.Fields.Add(End(document), FieldType.Empty, code, false);
                document.Content.InsertAfter("w ");
            }
            if (page < pages.Length - 1)
            {
                End(document).InsertBreak(BreakType.PageBreak);
            }
        }
        document.Content.InsertParagraphAfter();
        return document;
    }

    private static RangeWrapper End(WordprocessingDocumentWrapper document)
    {
        return document.Range(document.Content.End - 1, document.Content.End - 1);
    }

    private static string[] Lines(IndexWrapper index)
    {
        return index.Range.Fields[0].Result.Text.TrimEnd('\r').Split('\r');
    }

    private static FieldChar Begin(IndexWrapper index)
    {
        return index.Range.Fields[0].Range.StartElement.Descendants<FieldChar>().First();
    }

    //Reflow's way: the pages come from the caller, for the XE fields and for the two ends of a bookmark, and the field is not marked for Word
    //to update. A null page number leaves that page out.
    [Fact]
    public void PageNumbersComeFromTheCaller()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream,
            new[] { "XE \"Apple\"", "XE \"Unknown\"" }, new[] { "XE \"Apple\" \\b", "XE \"Ranged\" \\r Span" }, new[] { "XE \"Pear\"" });
        int Start(string code) => document.Fields.Where(f => f.Code.Text == code).Select(f => f.Range.Start).First();
        RangeWrapper span = document.Range(Start("XE \"Ranged\" \\r Span"), Start("XE \"Pear\"") - 1);
        document.Bookmarks.Add("Span", span);
        IndexWrapper index = document.Indexes.Add(End(document), HeadingSeparator.None);
        Dictionary<int, string> pages = new Dictionary<int, string>
        {
            [Start("XE \"Apple\"")] = "i",
            [Start("XE \"Apple\" \\b")] = "ii",
            [Start("XE \"Pear\"")] = "iv",
            [span.Start] = "iii",
            [span.End] = "v",
        };

        List<int> asked = new List<int>();
        index.Update(range =>
        {
            asked.Add(range.Start);
            return pages.TryGetValue(range.Start, out string page) ? page : null;
        });

        Assert.Equal(new[] { "Apple, i, ii", "Pear, iv", "Ranged, iii–v", "Unknown" }, Lines(index));
        Assert.Null(Begin(index).Dirty);
        Assert.Contains(span.Start, asked);
        Assert.Contains(span.End, asked);
    }

    //Without the pages laid out, the numbers are estimated from the breaks, and the field is marked for Word to update.
    [Fact]
    public void PageNumbersAreEstimatedFromTheBreaks()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[] { "XE \"One\"" }, new[] { "XE \"Two\"" }, new[] { "XE \"Three\"" });

        IndexWrapper index = document.Indexes.Add(End(document), HeadingSeparator.None);

        Assert.Equal(new[] { "One, 1", "Three, 3", "Two, 2" }, Lines(index));
        Assert.True(Begin(index).Dirty.Value);
    }

    //Where Word last broke a page (w:lastRenderedPageBreak) counts as a page break, and a section's own numbering (start and format) is
    //followed.
    [Fact]
    public void EstimatedPageNumbersFollowRenderedBreaksAndSectionNumbering()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "First \rLater \r";
        document.Fields.Add(document.Range(6, 6), FieldType.Empty, "XE \"First\"", false);
        RangeWrapper later = document.Paragraphs[1].Range;
        document.Fields.Add(document.Range(later.End - 1, later.End - 1), FieldType.Empty, "XE \"Second\"", false);
        Run run = document.Paragraphs[1].Paragraph.Elements<Run>().First();
        run.InsertAt(new LastRenderedPageBreak(), run.RunProperties == null ? 0 : 1);
        document.Content.InsertParagraphAfter();
        SectionProperties section = document.WordprocessingDocument.MainDocumentPart.Document.Body.GetFirstChild<SectionProperties>();
        section.AppendChild(new PageNumberType() { Start = 5, Format = NumberFormatValues.LowerRoman });

        IndexWrapper index = document.Indexes.Add(End(document), HeadingSeparator.None);

        Assert.Equal(new[] { "First, v", "Second, vi" }, Lines(index));
    }

    //Word's order and groups: symbols, then digits, then letters; lower case and no accent first; a hyphen or apostrophe passed over; a digit or
    //symbol its own group, a letter with an accent under its letter. Spaces round an entry go; a colon at the end stays; an empty entry is left out.
    [Fact]
    public void EntriesAreSortedAndGroupedAsWordDoes()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[]
        {
            "XE \"zz\"", "XE \"Zürich\"", "XE \"it's\"", "XE \"its\"", "XE \"co-op\"", "XE \"coop\"", "XE \"ångström\"", "XE \"A\"", "XE \"a\"",
            "XE \"2020\"", "XE \"1999\"", "XE \"(paren)\"", "XE \"#hash\"", "XE \"  padded  \"", "XE \"Trail:\"", "XE \"\"", "XE \"Spaces  inside\"",
        });

        IndexWrapper index = document.Indexes.Add(End(document));

        Assert.Equal(new[]
        {
            "#", "#hash, 1", "(", "(paren), 1", "1", "1999, 1", "2", "2020, 1", "A", "a, 1", "A, 1", "ångström, 1", "C", "coop, 1", "co-op, 1",
            "I", "its, 1", "it's, 1", "P", "padded, 1", "S", "Spaces  inside, 1", "T", "Trail:, 1", "Z", "Zürich, 1", "zz, 1",
        }, Lines(index));
    }

    //Run-in: the subentries at every level on the main entry's line; ": " before the first when the main entry has no pages, "; " otherwise.
    [Fact]
    public void RunInPutsEveryLevelOnTheMainEntrysLine()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream,
            new[] { "XE \"Fruit:Banana\"", "XE \"fruit:cherry\"" }, new[] { "XE \"Fruit\"", "XE \"Fruit:Banana\"", "XE \"Fruit:Banana:Ripe\"" });

        IndexWrapper index = document.Indexes.Add(End(document), HeadingSeparator.None, false, IndexType.RunIn);

        Assert.Equal(new[] { "fruit: cherry, 1", "Fruit, 2; Banana, 1, 2; Ripe, 2" }, Lines(index));
    }

    //Word: XE fields in a footnote and in a text box are in the index, at the page of the note's mark and of the text box's anchor; those in a
    //header are not.
    [Fact]
    public void EntriesInNotesAndTextBoxesCountButNotThoseInHeaders()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[] { "XE \"Main\"" }, new[] { "XE \"Second\"" });
        FootnoteWrapper note = document.Footnotes.Add(document.Range(document.Paragraphs[2].Range.Start, document.Paragraphs[2].Range.Start + 4), null, "Note ");
        RangeWrapper noteEnd = note.Range;
        noteEnd.Collapse(CollapseDirection.End);
        document.Fields.Add(noteEnd, FieldType.Empty, "XE \"InFootnote\"", false);
        ShapeWrapper box = document.Shapes.AddTextbox(TextOrientation.Horizontal, 20, 30, 120, 60, document.Range(0, 0));
        box.TextFrame.TextRange.Text = "Box ";
        RangeWrapper boxEnd = box.TextFrame.TextRange;
        boxEnd.Collapse(CollapseDirection.End);
        boxEnd.Move(TextUnit.Character, -1);
        document.Fields.Add(boxEnd, FieldType.Empty, "XE \"InTextBox\"", false);
        HeaderFooterWrapper header = document.Sections[0].Headers[HeaderFooterIndex.Primary];
        header.Range.Text = "Head ";
        document.Fields.Add(header.Range, FieldType.Empty, "XE \"InHeader\"", false);

        IndexWrapper index = document.Indexes.Add(End(document), HeadingSeparator.None);

        Assert.Equal(new[] { "InFootnote, 2", "InTextBox, 1", "Main, 1", "Second, 2" }, Lines(index));
    }

    //The properties are the switches of the field, changed as Word changes them (measured: \e and \r go first, \c before \z).
    [Fact]
    public void PropertiesChangeTheSwitchesAsWordDoes()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[] { "XE \"Alpha\"" });
        IndexWrapper index = document.Indexes.Add(End(document));
        FieldWrapper field = index.Range.Fields[0];
        Assert.Equal("INDEX \\h \"A\" \\z \"2057\"", field.Code.Text);
        Assert.Equal(HeadingSeparator.Letter, index.HeadingSeparator);
        Assert.Equal(2057, index.IndexLanguage);

        index.RightAlignPageNumbers = true;
        Assert.Equal("INDEX \\e \"\t\" \\h \"A\" \\z \"2057\"", field.Code.Text);
        index.Type = IndexType.RunIn;
        Assert.Equal("INDEX \\r \\e \"\t\" \\h \"A\" \\z \"2057\"", field.Code.Text);
        index.NumberOfColumns = 3;
        Assert.Equal("INDEX \\r \\e \"\t\" \\h \"A\" \\c \"3\" \\z \"2057\"", field.Code.Text);

        Assert.True(index.RightAlignPageNumbers);
        Assert.Equal(IndexType.RunIn, index.Type);
        Assert.Equal(3, index.NumberOfColumns);
        index.HeadingSeparator = HeadingSeparator.BlankLine;
        Assert.Equal(HeadingSeparator.BlankLine, index.HeadingSeparator);
        index.HeadingSeparator = HeadingSeparator.None;
        index.AccentedLetters = true;
        Assert.Equal("INDEX \\r \\e \"\t\" \\c \"3\" \\a \\z \"2057\"", field.Code.Text);
        //The tab is kept in the code as Word keeps it, a w:tab.
        Assert.Contains(field.Range.StartElement.Parent.Descendants<TabChar>(), t => t.Parent.Elements<FieldCode>().Any());
    }

    [Fact]
    public void IndexesAreListedAndUpdatingAgainGivesTheSameEntries()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[] { "XE \"Alpha\"", "XE \"Beta:Sub\"" }, new[] { "XE \"Alpha\" \\i" });
        IndexWrapper index = document.Indexes.Add(End(document));
        string[] first = Lines(index);
        int paragraphs = document.Paragraphs.Count;

        index.Update();
        document.Fields.Update();

        Assert.Equal(first, Lines(index));
        Assert.Equal(paragraphs, document.Paragraphs.Count);
        Assert.Equal(1, document.Indexes.Count);
        Assert.Equal(index, document.Indexes[0]);
        Assert.Single(document.Content.Indexes);
    }

    [Fact]
    public void DeleteTakesTheIndexAwayAndLeavesTheText()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[] { "XE \"Alpha\"", "XE \"Beta\"" });
        int paragraphs = document.Paragraphs.Count;
        IndexWrapper index = document.Indexes.Add(End(document));

        index.Delete();

        Assert.Equal(0, document.Indexes.Count);
        Assert.Equal(paragraphs, document.Paragraphs.Count);
        Assert.Equal(2, document.Fields.Count);
    }

    //Word writes MarkEntry's switches in this order; the bookmark of a page range goes as \r.
    [Fact]
    public void MarkEntryWritesAnXEField()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false);
        document.Content.Text = "Some words";

        FieldWrapper field = document.Indexes.MarkEntry(document.Range(0, 4), crossReference: "See other", bookmarkName: "Mark", bold: true);
        RangeWrapper words = document.Content;
        words.Find.Execute("words");
        FieldWrapper plain = document.Indexes.MarkEntry(words);

        Assert.Equal("XE \"Some\" \\t \"See other\" \\r Mark \\b", field.Code.Text);
        Assert.Equal("XE \"words\"", plain.Code.Text);
        Assert.Throws<NotSupportedException>(() => document.Indexes.MarkEntry(document.Range(0, 4), entryAutoText: "Auto"));
    }

    //The chapter number of \s is the number of the last SEQ field before the entry, even one that \h hides (measured in Word).
    [Fact]
    public void ChapterNumbersCountHiddenSequenceFields()
    {
        using MemoryStream stream = new MemoryStream();
        using WordprocessingDocumentWrapper document = Create(stream, new[] { "SEQ chapter \\h", "XE \"A\"" }, new[] { "SEQ chapter \\h", "XE \"B\"" });

        FieldWrapper field = document.Fields.Add(End(document), FieldType.Empty, "INDEX \\s chapter", false);

        Assert.Equal("A, 1-1\rB, 2-2\r", field.Result.Text);
    }
}
