using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Searching for formatting and replacing with it, with the font, paragraph format, style and highlighting of <c>Find</c> and its
/// <c>Replacement</c>. The markup is what Word wrote (the find-* scenarios of MarkupConformanceTest).
/// </summary>
public class RangeWrapperFindFormatTest
{
    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Run Bold(string text)
    {
        return new Run(new RunProperties(new Bold()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Body CreateBody(params Paragraph[] paragraphs)
    {
        return new Body(paragraphs);
    }

    //Each run as its properties and then its text.
    private static string Describe(Body body)
    {
        return string.Join("|", body.Descendants<Run>().Select(r => (r.RunProperties == null ? string.Empty : string.Concat(r.RunProperties.ChildElements.Select(e => e.LocalName + " "))) + r.InnerText));
    }

    private static void WithDocument(Action<WordprocessingDocumentWrapper> test)
    {
        string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
            {
                test(document);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void TheFormattingOfTheFindAndItsReplacementIsThereToSet()
    {
        FindWrapper find = new RangeWrapper(CreateBody(new Paragraph(Plain("text")))).Find;

        Assert.False(find.Format);
        find.Font.Bold = true;
        find.Replacement.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        Assert.True(find.Format);
        Assert.True(find.Font.Bold);
        Assert.Equal(ParagraphAlignment.Center, find.Replacement.ParagraphFormat.Alignment);

        find.ClearFormatting();
        Assert.False(find.Font.Bold);
        Assert.True(find.Format);
        find.Replacement.ClearFormatting();
        Assert.False(find.Format);
    }

    [Fact]
    public void TextIsFoundOnlyWhereItHasTheFormat()
    {
        Body body = CreateBody(new Paragraph(Plain("one two "), Bold("two"), Plain(" two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Font.Bold = true;

        Assert.True(range.Find.Execute("two"));
        Assert.Equal("two", range.Text);
        Assert.Equal("one two ".Length, range.Start);
        Assert.False(range.Find.Execute("two"));
    }

    [Fact]
    public void TextThatIsNotBoldCanBeFoundToo()
    {
        Body body = CreateBody(new Paragraph(Bold("two"), Plain(" two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Font.Bold = false;

        Assert.True(range.Find.Execute("two"));
        Assert.Equal(4, range.Start);
    }

    [Fact]
    public void FormattingIsIgnoredWhenFormatIsFalse()
    {
        Body body = CreateBody(new Paragraph(Plain("two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Font.Bold = true;
        range.Find.Format = false;

        Assert.True(range.Find.Execute("two"));
    }

    [Fact]
    public void ReplacingWithAFontFormatsTheReplacement()
    {
        Body body = CreateBody(new Paragraph(Plain("One two three two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Replacement.Font.Bold = true;
        range.Find.Replacement.Font.Size = 14;

        Assert.True(range.Find.Execute("two", replaceWith: "2", replace: ReplaceMode.All));

        Assert.Equal("One |b sz 2| three |b sz 2", Describe(body));
    }

    [Fact]
    public void ReplacingWithNoTextFormatsWhatWasFound()
    {
        Body body = CreateBody(new Paragraph(Plain("One two three two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Replacement.Font.Italic = true;

        Assert.True(range.Find.Execute("two", replace: ReplaceMode.All));

        Assert.Equal("One |i two| three |i two", Describe(body));
    }

    [Fact]
    public void FormattingWithoutTextFindsTheRunsThatHaveIt()
    {
        Body body = CreateBody(new Paragraph(Plain("plain "), Bold("bold"), Plain(" plain "), Bold("bold")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Font.Bold = true;
        range.Find.Replacement.Font.Italic = true;
        range.Find.Replacement.Font.Bold = false;

        Assert.True(range.Find.Execute(string.Empty, replace: ReplaceMode.All));

        Assert.Equal("plain |i bold| plain |i bold", Describe(body));
    }

    [Fact]
    public void FormattingWithoutTextCanBeFoundOneAtATime()
    {
        Body body = CreateBody(new Paragraph(Plain("a "), Bold("b"), Plain(" c "), Bold("d")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Font.Bold = true;

        Assert.True(range.Find.Execute(string.Empty));
        Assert.Equal("b", range.Text);
        Assert.True(range.Find.Execute(string.Empty));
        Assert.Equal("d", range.Text);
        Assert.False(range.Find.Execute(string.Empty));
    }

    [Fact]
    public void TheParagraphFormatIsGivenToTheParagraphsOfTheReplacement()
    {
        Body body = CreateBody(new Paragraph(Plain("One")), new Paragraph(Plain("Two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Replacement.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        Assert.True(range.Find.Execute("Two", replaceWith: "Two", replace: ReplaceMode.All));

        Assert.Null(body.Elements<Paragraph>().First().ParagraphProperties);
        Assert.Equal(JustificationValues.Center, body.Elements<Paragraph>().Last().ParagraphProperties.Justification.Val.Value);
    }

    [Fact]
    public void AParagraphFormatIsFoundOnTheParagraphsThatHaveIt()
    {
        Body body = CreateBody(
            new Paragraph(Plain("One two")),
            new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), Plain("Three two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        Assert.True(range.Find.Execute("two"));
        Assert.StartsWith("Three", range.Paragraphs[0].Text, StringComparison.Ordinal);
        Assert.False(range.Find.Execute("two"));
    }

    [Fact]
    public void AParagraphFormatWithNoTextFindsWholeParagraphs()
    {
        Body body = CreateBody(
            new Paragraph(Plain("One")),
            new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }), Plain("Two")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        Assert.True(range.Find.Execute(string.Empty));
        Assert.Equal("Two\r", range.Text);
    }

    [Fact]
    public void HighlightingIsReplacedAndFound()
    {
        Body body = CreateBody(new Paragraph(Plain("One two three")));
        RangeWrapper range = new RangeWrapper(body);
        range.Find.Replacement.Highlight = true;
        Assert.True(range.Find.Execute("two", replaceWith: "^&", replace: ReplaceMode.All));
        Assert.Equal(HighlightColorValues.Yellow, body.Descendants<Highlight>().Single().Val.Value);

        RangeWrapper again = new RangeWrapper(body);
        again.Find.Highlight = true;
        Assert.True(again.Find.Execute("two"));
        Assert.Equal("two", again.Text);

        RangeWrapper not = new RangeWrapper(body);
        not.Find.Highlight = false;
        Assert.True(not.Find.Execute("three"));
    }

    [Fact]
    public void AStyleThatTheDocumentDoesNotHaveIsAnErrorAndNothingIsChanged()
    {
        WithDocument(document =>
        {
            document.Content.InsertAfter("One two");

            RangeWrapper find = document.Content;
            find.Find.Style = "No Such Style";
            Assert.Throws<ArgumentException>(() => find.Find.Execute("two"));

            RangeWrapper replace = document.Content;
            replace.Find.Replacement.Style = "No Such Style";
            Assert.Throws<ArgumentException>(() => replace.Find.Execute("two", replaceWith: "2", replace: ReplaceMode.All));
            Assert.Equal("One two", document.Content.Text.TrimEnd('\r'));
        });
    }

    [Fact]
    public void AStyleIsFoundAndGiven()
    {
        WithDocument(document =>
        {
            document.Content.InsertAfter("One");
            document.Content.InsertParagraphAfter();
            document.Content.InsertAfter("Two");
            document.Paragraphs[0].Style = "Heading 1";

            RangeWrapper range = document.Content;
            range.Find.Style = "Heading 1";
            Assert.True(range.Find.Execute("One"));
            RangeWrapper two = document.Content;
            two.Find.Style = "Heading 1";
            Assert.False(two.Find.Execute("Two"));

            RangeWrapper replace = document.Content;
            replace.Find.Replacement.Style = "Heading 2";
            Assert.True(replace.Find.Execute("Two", replaceWith: "Dos", replace: ReplaceMode.All));
            Assert.Equal("Heading 2", document.Paragraphs[1].Style);
            Assert.Equal("Dos", document.Paragraphs[1].Text.TrimEnd('\r'));
        });
    }
}
