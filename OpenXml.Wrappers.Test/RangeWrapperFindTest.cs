using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Test;

public class RangeWrapperFindTest
{
    private static Paragraph CreateParagraph(params string[] runTexts)
    {
        return new Paragraph(runTexts.Select(t => new Run(new Text(t))));
    }

    private static string TextOf(Paragraph paragraph)
    {
        return new ParagraphWrapper(paragraph).Text.TrimEnd('\r');
    }

    [Fact]
    public void FindReportsAMatchWithoutChangingTheDocument()
    {
        Paragraph paragraph = CreateParagraph("Hello world");
        string before = paragraph.OuterXml;
        FindWrapper find = new RangeWrapper(paragraph).Find;
        find.Text = "world";

        Assert.True(find.Execute());
        Assert.True(find.Found);
        Assert.Equal(before, paragraph.OuterXml);
    }

    [Fact]
    public void FindReportsNoMatch()
    {
        FindWrapper find = new RangeWrapper(CreateParagraph("Hello")).Find;
        find.Text = "xyz";

        Assert.False(find.Execute());
        Assert.False(find.Found);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void FindWithNoTextFindsNothing(string text)
    {
        FindWrapper find = new RangeWrapper(CreateParagraph("Hello")).Find;
        find.Text = text;

        Assert.False(find.Execute());
    }

    [Fact]
    public void FoundIsResetByTheNextExecute()
    {
        FindWrapper find = new RangeWrapper(CreateParagraph("Hello")).Find;
        find.Text = "Hello";
        Assert.True(find.Execute());

        find.Text = "xyz";
        Assert.False(find.Execute());
        Assert.False(find.Found);
    }

    [Fact]
    public void TheFindObjectIsTheSameEachTimeItIsRequested()
    {
        RangeWrapper range = new RangeWrapper(CreateParagraph("a"));

        Assert.Same(range.Find, range.Find);
    }

    [Fact]
    public void FindIsCaseInsensitiveUnlessMatchCaseIsSet()
    {
        Assert.True(new RangeWrapper(CreateParagraph("Hello")).Find.Execute("hello"));
        Assert.False(new RangeWrapper(CreateParagraph("Hello")).Find.Execute("hello", matchCase: true));
        Assert.True(new RangeWrapper(CreateParagraph("Hello")).Find.Execute("Hello", matchCase: true));
    }

    [Fact]
    public void FindCanMatchWholeWordsOnly()
    {
        FindWrapper find = new RangeWrapper(CreateParagraph("concatenate")).Find;

        Assert.True(find.Execute("cat"));
        Assert.False(find.Execute("cat", matchWholeWord: true));
        Assert.True(new RangeWrapper(CreateParagraph("a cat.")).Find.Execute("cat", matchWholeWord: true));
    }

    [Fact]
    public void ReplaceAllReplacesEveryMatchInEveryParagraph()
    {
        Paragraph first = CreateParagraph("cat and cat");
        Paragraph second = CreateParagraph("a cat");
        Body body = new Body(first, second);

        Assert.True(new RangeWrapper(body).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All));

        Assert.Equal("dog and dog", TextOf(first));
        Assert.Equal("a dog", TextOf(second));
    }

    [Fact]
    public void ReplaceOneReplacesOnlyTheFirstMatch()
    {
        Paragraph first = CreateParagraph("cat and cat");
        Paragraph second = CreateParagraph("a cat");
        Body body = new Body(first, second);

        Assert.True(new RangeWrapper(body).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.One));

        Assert.Equal("dog and cat", TextOf(first));
        Assert.Equal("a cat", TextOf(second));
    }

    [Fact]
    public void ReplaceReturnsFalseWhenThereIsNothingToReplace()
    {
        Paragraph paragraph = CreateParagraph("cat");

        Assert.False(new RangeWrapper(paragraph).Find.Execute("bird", replaceWith: "dog", replace: ReplaceMode.All));
        Assert.Equal("cat", TextOf(paragraph));
    }

    [Fact]
    public void ReplaceUsesTheReplacementPropertyAndNullDeletesTheMatch()
    {
        Paragraph paragraph = CreateParagraph("a cat here");
        FindWrapper find = new RangeWrapper(paragraph).Find;
        find.Text = "cat ";
        find.Replacement.Text = null;

        Assert.True(find.Execute(ReplaceMode.All));
        Assert.Equal("a here", TextOf(paragraph));
    }

    [Fact]
    public void ReplaceTakesTheTextOfTheReplacementNotTheCaseOfTheMatch()
    {
        Paragraph paragraph = CreateParagraph("CAT");
        new RangeWrapper(paragraph).Find.Execute("cat", replaceWith: "Dog", replace: ReplaceMode.All);

        Assert.Equal("Dog", TextOf(paragraph));
    }

    [Fact]
    public void ReplaceMatchesDoNotOverlap()
    {
        Paragraph paragraph = CreateParagraph("aaa");
        new RangeWrapper(paragraph).Find.Execute("aa", replaceWith: "b", replace: ReplaceMode.All);

        Assert.Equal("ba", TextOf(paragraph));
    }

    [Fact]
    public void ReplaceCanMatchTextSpanningSeveralRunsAndKeepsTheFirstRunsFormatting()
    {
        Run first = new Run(new RunProperties(new Bold()), new Text("Hel"));
        Run second = new Run(new Text("lo wo"));
        Run third = new Run(new Text("rld"));
        Paragraph paragraph = new Paragraph(first, second, third);

        new RangeWrapper(paragraph).Find.Execute("llo wor", replaceWith: "X", replace: ReplaceMode.All);

        Assert.Equal("HeXld", TextOf(paragraph));
        Assert.Null(second.Parent);
        Assert.Single(first.RunProperties.Elements<Bold>());
    }

    [Fact]
    public void ReplaceKeepsTextOnBothSidesInASingleRun()
    {
        Paragraph paragraph = CreateParagraph(" a cat  here ");
        new RangeWrapper(paragraph).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All);

        Assert.Equal(" a dog  here ", TextOf(paragraph));
        Assert.Equal(SpaceProcessingModeValues.Preserve, paragraph.Descendants<Text>().Single().Space.Value);
    }

    [Fact]
    public void FindMatchesTabsAndLineBreaks()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a"), new TabChar(), new Text("b"), new Break(), new Text("c")));
        Assert.True(new RangeWrapper(paragraph).Find.Execute("a\tb"));
        Assert.True(new RangeWrapper(paragraph).Find.Execute("b\vc"));
        Assert.False(new RangeWrapper(paragraph).Find.Execute("a b"));
    }

    [Fact]
    public void ReplaceCanReplaceATab()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a"), new TabChar(), new Text("b")));
        new RangeWrapper(paragraph).Find.Execute("a\tb", replaceWith: "X", replace: ReplaceMode.All);

        Assert.Equal("X", TextOf(paragraph));
        Assert.Empty(paragraph.Descendants<TabChar>());
    }

    [Fact]
    public void ReplaceCanReplaceOnlyATab()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a"), new TabChar(), new Text("b")));
        new RangeWrapper(paragraph).Find.Execute("\t", replaceWith: " ", replace: ReplaceMode.All);

        Assert.Equal("a b", TextOf(paragraph));
    }

    [Theory]
    [InlineData("a\tb", typeof(TabChar))]
    [InlineData("a\nb", typeof(Break))]
    [InlineData("a\vb", typeof(Break))]
    [InlineData("a\fb", typeof(Break))]
    [InlineData("a^tb", typeof(TabChar))]
    [InlineData("a^lb", typeof(Break))]
    public void ReplacementTextCanContainTabsAndBreaks(string replacement, Type expected)
    {
        Paragraph paragraph = CreateParagraph("cat");
        Body body = new Body(paragraph);

        new RangeWrapper(body).Find.Execute("cat", replaceWith: replacement, replace: ReplaceMode.All);

        Assert.Single(paragraph.Descendants(), e => e.GetType() == expected);
        Assert.Equal("a", paragraph.Descendants<Text>().First().Text);
        Assert.Equal("b", paragraph.Descendants<Text>().Last().Text);
    }

    [Theory]
    [InlineData("a\rb")]
    [InlineData("a^pb")]
    public void AParagraphMarkInTheReplacementStartsAParagraph(string replacement)
    {
        Body body = new Body(CreateParagraph("x cat y"));

        new RangeWrapper(body).Find.Execute("cat", replaceWith: replacement, replace: ReplaceMode.All);

        Assert.Equal(2, body.Elements<Paragraph>().Count());
        Assert.Equal("x a\rb y\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void AMatchCanSpanParagraphs()
    {
        Body body = new Body(CreateParagraph("ca"), CreateParagraph("t"));

        Assert.True(new RangeWrapper(body).Find.Execute("ca^pt"));
        Assert.True(new RangeWrapper(body).Find.Execute("ca\rt", replaceWith: "X", replace: ReplaceMode.All));
        Assert.Equal("X\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void AFoundRangeIsNarrowedToTheMatch()
    {
        Paragraph paragraph = CreateParagraph("one two three");
        RangeWrapper range = new RangeWrapper(paragraph);

        Assert.True(range.Find.Execute("two"));

        Assert.Equal("two", range.Text);
        Assert.Equal(4, range.Start);
        Assert.Equal(7, range.End);
        range.Font.Bold = true;
        Assert.True(range.Font.Bold);
        Assert.Equal(3, paragraph.Descendants<Run>().Count());
    }

    [Fact]
    public void ExecuteCarriesOnAfterTheLastMatchUntilTheEndOfTheOriginalRange()
    {
        Body body = new Body(CreateParagraph("a b a b a"));
        RangeWrapper range = new RangeWrapper(body);
        List<int> starts = new List<int>();

        while (range.Find.Execute("a"))
        {
            starts.Add(range.Start);
        }

        Assert.Equal(new[] { 0, 4, 8 }, starts);
    }

    [Fact]
    public void ReplaceOneLeavesTheRangeOnTheReplacement()
    {
        Body body = new Body(CreateParagraph("a cat sat"));
        RangeWrapper range = new RangeWrapper(body);

        Assert.True(range.Find.Execute("cat", replaceWith: "kitten", replace: ReplaceMode.One));

        Assert.Equal("kitten", range.Text);
        Assert.Equal("a kitten sat\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void ReplacingPartOfARunKeepsItsFormattingAndSplitsNothing()
    {
        Run run = new Run(new RunProperties(new Italic()), new Text("a cat sat"));
        Body body = new Body(new Paragraph(run));

        new RangeWrapper(body).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All);

        Assert.Single(body.Descendants<Run>());
        Assert.Equal("a dog sat", body.Descendants<Text>().Single().Text);
        Assert.NotNull(run.RunProperties.Italic);
    }

    [Fact]
    public void MatchesDoNotSpanParagraphs()
    {
        Body body = new Body(CreateParagraph("ca"), CreateParagraph("t"));

        Assert.False(new RangeWrapper(body).Find.Execute("cat"));
    }

    [Fact]
    public void RunsInsideHyperlinksAreSearched()
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("a ")), new Hyperlink(new Run(new Text("cat"))));
        new RangeWrapper(paragraph).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All);

        Assert.Equal("a dog", TextOf(paragraph));
    }

    [Fact]
    public void AFindInABookmarkRangeOnlyAffectsTheBookmark()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        BookmarkEnd end = new BookmarkEnd() { Id = "1" };
        Paragraph paragraph = new Paragraph(new Run(new Text("cat ")), start, new Run(new Text("cat")), end, new Run(new Text(" cat")));
        Body body = new Body(paragraph);
        RangeWrapper range = new RangeWrapper(start, end);

        Assert.True(range.Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All));

        Assert.Equal("cat dog cat", TextOf(paragraph));
    }

    [Fact]
    public void AFindInATableCellOnlyAffectsTheCell()
    {
        TableCell cell = new TableCell(CreateParagraph("cat"));
        Paragraph outside = CreateParagraph("cat");
        Body body = new Body(new Table(new TableRow(cell)), outside);

        new RangeWrapper(cell).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All);

        Assert.Equal("dog", new RangeWrapper(cell).Text.TrimEnd('\r', '\a'));
        Assert.Equal("cat", TextOf(outside));
    }

    [Fact]
    public void AFindInATableSearchesEveryCell()
    {
        Table table = new Table(
            new TableRow(new TableCell(CreateParagraph("cat")), new TableCell(CreateParagraph("a cat"))),
            new TableRow(new TableCell(CreateParagraph("cat"))));

        new RangeWrapper(table).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All);

        Assert.DoesNotContain("cat", new RangeWrapper(table).Text);
        Assert.Equal(3, new RangeWrapper(table).Text.Split(new[] { "dog" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void AFindInADocumentSearchesTheWholeBody()
    {
        Document document = new Document(new Body(CreateParagraph("cat"), new Table(new TableRow(new TableCell(CreateParagraph("cat")))), new SectionProperties()));

        Assert.True(new RangeWrapper(document).Find.Execute("cat", replaceWith: "dog", replace: ReplaceMode.All));
        Assert.DoesNotContain(document.Descendants<Text>(), t => t.Text == "cat");
    }

    [Fact]
    public void ReplacingEverythingInARunRemovesTheEmptyRun()
    {
        Run run = new Run(new RunProperties(new Bold()), new Text("cat"));
        Paragraph paragraph = new Paragraph(run, new Run(new Text("x")));

        new RangeWrapper(paragraph).Find.Execute("cat", replaceWith: null, replace: ReplaceMode.All);

        Assert.Null(run.Parent);
        Assert.Equal("x", TextOf(paragraph));
    }

    [Fact]
    public void AFindInARunWithNoParagraphFindsNothing()
    {
        Assert.False(new RangeWrapper(new Run(new Text("a"))).Find.Execute("a"));
    }
}
