using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Ranges made of a span of characters: positions, reading, replacing, inserting, deleting and formatting part of a run.
/// </summary>
public class RangeWrapperSpanTest
{
    private static Paragraph Paragraph(params string[] runs)
    {
        return new Paragraph(runs.Select(text => new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })).ToArray());
    }

    private static string StoryText(Body body)
    {
        return new RangeWrapper(body).Text;
    }

    private static RangeWrapper Range(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    #region Positions

    [Fact]
    public void ElementRangesHavePositions()
    {
        Paragraph first = Paragraph("Hello");
        Paragraph second = Paragraph("World");
        Body body = new Body(first, second);

        Assert.Equal(0, new ParagraphWrapper(first).Range.Start);
        Assert.Equal(6, new ParagraphWrapper(first).Range.End);
        Assert.Equal(6, new ParagraphWrapper(second).Range.Start);
        Assert.Equal(12, new ParagraphWrapper(second).Range.End);
        Assert.Equal(12, new RangeWrapper(body).StoryLength);
    }

    [Fact]
    public void ABookmarkRangeHasThePositionsOfItsText()
    {
        Paragraph paragraph = new Paragraph(
            new Run(new Text("ab")),
            new BookmarkStart() { Id = "1", Name = "m" },
            new Run(new Text("cde")),
            new BookmarkEnd() { Id = "1" },
            new Run(new Text("f")));
        Body body = new Body(paragraph);
        RangeWrapper range = new RangeWrapper(paragraph).Bookmarks[0].Range;

        Assert.Equal(2, range.Start);
        Assert.Equal(5, range.End);
    }

    [Fact]
    public void TableCellsAndRowsCountTheirMarkers()
    {
        TableCell first = new TableCell(Paragraph("a"));
        TableCell second = new TableCell(Paragraph("b"));
        Table table = new Table(new TableRow(first, second));
        Body body = new Body(table, Paragraph("after"));

        //Each cell is its text and the end of the cell, then the end of the row, then the paragraph.
        Assert.Equal(2 + 2 + 1 + 6, new RangeWrapper(body).StoryLength);
        Assert.Equal(2, new TableCellWrapper(second).Range.Start);
    }

    [Fact]
    public void SettingTheRangeGivesTheTextBetweenThePositions()
    {
        Body body = new Body(Paragraph("Hello"), Paragraph("World"));

        Assert.Equal("Hello", Range(body, 0, 5).Text);
        Assert.Equal("lo\rWo", Range(body, 3, 8).Text);
        Assert.Equal(string.Empty, Range(body, 4, 4).Text);
        Assert.Throws<System.ArgumentOutOfRangeException>(() => Range(body, 5, 2));
        Assert.Throws<System.ArgumentOutOfRangeException>(() => Range(body, 0, 99));
    }

    #endregion

    #region Editing

    [Fact]
    public void ReplacingTextInTheMiddleOfARunKeepsTheFormatting()
    {
        Run run = new Run(new RunProperties(new Bold()), new Text("Hello world") { Space = SpaceProcessingModeValues.Preserve });
        Body body = new Body(new Paragraph(run));

        Range(body, 6, 11).Text = "there";

        Assert.Equal("Hello there\r", StoryText(body));
        Assert.All(body.Descendants<Run>(), r => Assert.NotNull(r.RunProperties?.Bold));
    }

    [Fact]
    public void TheRangeCoversTheNewText()
    {
        Body body = new Body(Paragraph("Hello world"));
        RangeWrapper range = Range(body, 6, 11);

        range.Text = "everyone";

        Assert.Equal(6, range.Start);
        Assert.Equal(14, range.End);
        Assert.Equal("everyone", range.Text);
    }

    [Fact]
    public void InsertingAtACollapsedRange()
    {
        Body body = new Body(Paragraph("Helloworld"));
        RangeWrapper range = Range(body, 5, 5);

        range.InsertAfter(" ");

        Assert.Equal("Hello world\r", StoryText(body));
        Assert.Equal(5, range.Start);
        Assert.Equal(6, range.End);

        range.InsertBefore(">");
        Assert.Equal("Hello> world\r", StoryText(body));
    }

    [Fact]
    public void TextInsertedAtTheStartOfAParagraphAndInAnEmptyParagraph()
    {
        Body body = new Body(Paragraph("b"), new Paragraph());

        Range(body, 0, 0).InsertBefore("a");
        Range(body, 3, 3).InsertAfter("x");

        Assert.Equal("ab\rx\r", StoryText(body));
    }

    [Fact]
    public void ACarriageReturnInTheTextStartsAParagraph()
    {
        Body body = new Body(Paragraph("onetwo"));

        Range(body, 3, 3).InsertAfter("\r");

        Assert.Equal(2, body.Elements<Paragraph>().Count());
        Assert.Equal("one\rtwo\r", StoryText(body));
    }

    [Fact]
    public void DeletingPartOfARunSplitsIt()
    {
        Body body = new Body(Paragraph("Hello cruel", " world"));

        Range(body, 5, 11).Delete();

        Assert.Equal("Hello world\r", StoryText(body));
    }

    [Fact]
    public void DeletingAcrossParagraphsJoinsThem()
    {
        Body body = new Body(Paragraph("Hello"), Paragraph("middle"), Paragraph("World"));

        Range(body, 3, 15).Delete();

        Assert.Equal("Helrld\r", StoryText(body));
        Assert.Single(body.Elements<Paragraph>());
    }

    [Fact]
    public void DeletingWholeParagraphsRemovesThem()
    {
        Body body = new Body(Paragraph("one"), Paragraph("two"), Paragraph("three"));

        Range(body, 4, 8).Delete();

        Assert.Equal("one\rthree\r", StoryText(body));
    }

    [Fact]
    public void DeletingTheContentOfACellLeavesTheCell()
    {
        TableCell cell = new TableCell(Paragraph("abc"));
        Body body = new Body(new Table(new TableRow(cell)), Paragraph("x"));

        Range(body, 0, 3).Delete();

        Assert.Single(cell.Elements<Paragraph>());
        Assert.Empty(cell.Descendants<Run>());
        Assert.Single(body.Descendants<TableCell>());
    }

    [Fact]
    public void DeletingAWholeTableAloneClearsItAsInWord()
    {
        Body body = new Body(Paragraph("before"), new Table(new TableRow(new TableCell(Paragraph("a")))), Paragraph("after"));

        //The table is the cell's text and end of cell, and the end of the row. Word clears a table that is the whole range and keeps it.
        Range(body, 7, 7 + 2 + 1).Delete();

        Assert.Single(body.Descendants<Table>());
        Assert.Equal("before\r\r\a\r\aafter\r", StoryText(body));
    }

    [Fact]
    public void DeletingAWholeTableWithTextAroundItRemovesIt()
    {
        Body body = new Body(Paragraph("before"), new Table(new TableRow(new TableCell(Paragraph("a")))), Paragraph("after"));

        Range(body, 0, 7 + 2 + 1).Delete();

        Assert.Empty(body.Descendants<Table>());
        Assert.Equal("after\r", StoryText(body));
    }

    [Fact]
    public void DeletingPartOfAFieldRemovesTheWholeField()
    {
        Run Begin() => new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
        Run Code() => new Run(new FieldCode(" PAGE "));
        Run Separate() => new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
        Run End() => new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
        Body body = new Body(new Paragraph(new Run(new Text("ab")), Begin(), Code(), Separate(), new Run(new Text("12")), End(), new Run(new Text("cd"))));

        Range(body, 2, 3).Delete();

        Assert.Equal("abcd\r", StoryText(body));
        Assert.Empty(body.Descendants<FieldChar>());
    }

    #endregion

    #region Formatting

    [Fact]
    public void FormattingPartOfARunSplitsTheRun()
    {
        Body body = new Body(Paragraph("Hello world"));

        Range(body, 6, 11).Font.Bold = true;

        List<Run> runs = body.Descendants<Run>().ToList();
        Assert.Equal(2, runs.Count);
        Assert.Null(runs[0].RunProperties?.Bold);
        Assert.NotNull(runs[1].RunProperties.Bold);
        Assert.Equal("Hello world\r", StoryText(body));
        Assert.True(Range(body, 6, 11).Font.Bold);
        Assert.False(Range(body, 0, 5).Font.Bold);
        Assert.Null(Range(body, 0, 11).Font.Bold);
    }

    [Fact]
    public void ParagraphFormatAppliesToEveryParagraphTouched()
    {
        Body body = new Body(Paragraph("one"), Paragraph("two"), Paragraph("three"));

        Range(body, 2, 6).ParagraphFormat.Alignment = ParagraphAlignment.Center;

        Assert.All(body.Elements<Paragraph>().Take(2), p => Assert.Equal(JustificationValues.Center, p.ParagraphProperties.Justification.Val.Value));
        Assert.Null(body.Elements<Paragraph>().Last().ParagraphProperties?.Justification);
    }

    #endregion

    #region Moving

    [Fact]
    public void CollapseMovesToTheStartOrEnd()
    {
        Body body = new Body(Paragraph("Hello"));
        RangeWrapper range = Range(body, 1, 4);

        range.Collapse(CollapseDirection.End);
        Assert.Equal(4, range.Start);
        Assert.Equal(4, range.End);

        range = Range(body, 1, 4);
        range.Collapse();
        Assert.Equal(1, range.End);
    }

    [Fact]
    public void MovingByCharactersAndWords()
    {
        Body body = new Body(Paragraph("The quick brown fox"));
        RangeWrapper range = Range(body, 0, 3);

        Assert.Equal(1, range.MoveEnd(TextUnit.Word));
        Assert.Equal("The ", range.Text);
        Assert.Equal(1, range.MoveEnd(TextUnit.Word));
        Assert.Equal("The quick ", range.Text);

        Assert.Equal(1, range.MoveStart(TextUnit.Word));
        Assert.Equal("quick ", range.Text);

        Assert.Equal(-1, range.MoveStart(TextUnit.Character, -1));
        Assert.Equal(" quick ", range.Text);

        Assert.Equal(1, range.Move(TextUnit.Word));
        Assert.Equal(range.Start, range.End);
        Assert.Equal(16, range.Start);
    }

    [Fact]
    public void MovingStopsAtTheEndOfTheStory()
    {
        Body body = new Body(Paragraph("ab"));
        RangeWrapper range = Range(body, 0, 0);

        Assert.Equal(3, range.MoveEnd(TextUnit.Character, 10));
        Assert.Equal(3, range.End);
        Assert.Equal(0, range.MoveStart(TextUnit.Character, -10));
        Assert.Equal(0, range.Start);
    }

    [Fact]
    public void ExpandingToAWordAndASentenceAndAParagraph()
    {
        Body body = new Body(Paragraph("One two. Three four."), Paragraph("Next"));

        RangeWrapper range = Range(body, 5, 5);
        range.Expand(TextUnit.Word);
        Assert.Equal("two", range.Text);

        range = Range(body, 5, 5);
        range.Expand(TextUnit.Sentence);
        Assert.Equal("One two. ", range.Text);

        range = Range(body, 12, 12);
        range.Expand(TextUnit.Paragraph);
        Assert.Equal("One two. Three four.\r", range.Text);
    }

    [Fact]
    public void CharactersWordsAndSentencesAreRanges()
    {
        Body body = new Body(Paragraph("Hi there. Bye now."));
        RangeWrapper range = Range(body, 0, 18);

        Assert.Equal(18, range.Characters.Count);
        Assert.Equal("H", range.Characters[0].Text);
        Assert.Equal(new[] { "Hi ", "there", ". ", "Bye ", "now", "." }, range.Words.Take(6).Select(w => w.Text).ToArray());
        Assert.Equal(new[] { "Hi there. ", "Bye now." }, range.Sentences.Select(s => s.Text).ToArray());

        range.Words[1].Font.Italic = true;
        Assert.True(Range(body, 3, 8).Font.Italic);
        Assert.False(Range(body, 0, 3).Font.Italic);
    }

    [Fact]
    public void ADuplicateIsIndependent()
    {
        Body body = new Body(Paragraph("abcdef"));
        RangeWrapper range = Range(body, 1, 3);
        RangeWrapper copy = range.Duplicate();

        copy.Collapse(CollapseDirection.End);

        Assert.Equal(1, range.Start);
        Assert.Equal(3, copy.Start);
        Assert.NotEqual(range, copy);
        Assert.Equal(range, range.Duplicate());
    }

    #endregion
}
