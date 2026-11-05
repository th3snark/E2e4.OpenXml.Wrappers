using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Moving to the start or end of a unit, to the next or previous unit and until a character, comparing ranges and counting. Every number is
/// the one Word gives for the same text (found by running the same calls in Word 16). The text is 74 positions long, with the last paragraph
/// mark at 73:
/// "The quick brown fox. Second sentence here.\rThird para,  with commas.\rLast\r".
/// </summary>
public class RangeWrapperNavigationTest
{
    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Body NewBody()
    {
        return new Body(
            new Paragraph(Plain("The quick brown fox. Second sentence here.")),
            new Paragraph(Plain("Third para,  with commas.")),
            new Paragraph(Plain("Last")));
    }

    private static RangeWrapper At(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    private static string Show(RangeWrapper range)
    {
        return range == null ? "null" : string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0}-{1}", range.Start, range.End);
    }

    [Theory]
    [InlineData(6, 12, TextUnit.Word, Movement.Move, -3, "4-4")]
    [InlineData(6, 12, TextUnit.Word, Movement.Extend, -2, "4-12")]
    [InlineData(4, 9, TextUnit.Word, Movement.Move, -1, "4-4")]
    [InlineData(6, 6, TextUnit.Word, Movement.Move, -2, "4-4")]
    [InlineData(30, 33, TextUnit.Sentence, Movement.Move, -10, "21-21")]
    [InlineData(30, 33, TextUnit.Sentence, Movement.Extend, -9, "21-33")]
    [InlineData(50, 55, TextUnit.Paragraph, Movement.Move, -8, "43-43")]
    [InlineData(50, 55, TextUnit.Paragraph, Movement.Extend, -7, "43-55")]
    [InlineData(50, 55, TextUnit.Story, Movement.Move, -51, "0-0")]
    [InlineData(50, 55, TextUnit.Story, Movement.Extend, -50, "0-55")]
    [InlineData(5, 8, TextUnit.Character, Movement.Move, -1, "5-5")]
    [InlineData(5, 8, TextUnit.Character, Movement.Extend, 0, "5-8")]
    [InlineData(73, 73, TextUnit.Word, Movement.Move, 0, "73-73")]
    [InlineData(73, 74, TextUnit.Word, Movement.Move, -1, "73-73")]
    [InlineData(73, 74, TextUnit.Paragraph, Movement.Move, -5, "69-69")]
    public void StartOfMovesToTheStartOfTheUnit(int start, int end, TextUnit unit, Movement extend, int returned, string expected)
    {
        RangeWrapper range = At(NewBody(), start, end);

        Assert.Equal(returned, range.StartOf(unit, extend));
        Assert.Equal(expected, Show(range));
    }

    [Theory]
    [InlineData(6, 12, TextUnit.Word, Movement.Move, 5, "16-16")]
    [InlineData(6, 12, TextUnit.Word, Movement.Extend, 4, "6-16")]
    [InlineData(4, 10, TextUnit.Word, Movement.Move, 1, "10-10")]
    [InlineData(6, 6, TextUnit.Word, Movement.Move, 4, "10-10")]
    [InlineData(30, 33, TextUnit.Sentence, Movement.Move, 11, "43-43")]
    [InlineData(30, 33, TextUnit.Sentence, Movement.Extend, 10, "30-43")]
    [InlineData(50, 55, TextUnit.Paragraph, Movement.Move, 15, "69-69")]
    [InlineData(50, 55, TextUnit.Paragraph, Movement.Extend, 14, "50-69")]
    [InlineData(5, 8, TextUnit.Story, Movement.Move, 66, "73-73")]
    [InlineData(5, 8, TextUnit.Story, Movement.Extend, 66, "5-74")]
    [InlineData(5, 8, TextUnit.Character, Movement.Move, 1, "8-8")]
    [InlineData(5, 8, TextUnit.Character, Movement.Extend, 0, "5-8")]
    [InlineData(73, 73, TextUnit.Word, Movement.Move, 0, "73-73")]
    [InlineData(73, 73, TextUnit.Word, Movement.Extend, 1, "73-74")]
    [InlineData(70, 72, TextUnit.Paragraph, Movement.Move, 2, "73-73")]
    [InlineData(70, 72, TextUnit.Paragraph, Movement.Extend, 2, "70-74")]
    [InlineData(70, 72, TextUnit.Word, Movement.Move, 2, "73-73")]
    [InlineData(70, 72, TextUnit.Word, Movement.Extend, 1, "70-73")]
    [InlineData(69, 73, TextUnit.Word, Movement.Extend, 0, "69-73")]
    [InlineData(69, 73, TextUnit.Sentence, Movement.Move, 1, "73-73")]
    [InlineData(0, 74, TextUnit.Paragraph, Movement.Extend, 0, "0-74")]
    public void EndOfMovesToTheEndOfTheUnit(int start, int end, TextUnit unit, Movement extend, int returned, string expected)
    {
        RangeWrapper range = At(NewBody(), start, end);

        Assert.Equal(returned, range.EndOf(unit, extend));
        Assert.Equal(expected, Show(range));
    }

    [Fact]
    public void StartOfAndEndOfDefaultToAWordMoved()
    {
        RangeWrapper start = At(NewBody(), 6, 12);
        RangeWrapper end = At(NewBody(), 6, 12);

        Assert.Equal(-3, start.StartOf());
        Assert.Equal("4-4", Show(start));
        Assert.Equal(5, end.EndOf());
        Assert.Equal("16-16", Show(end));
    }

    [Theory]
    [InlineData(4, 9, TextUnit.Word, 1, "10-16")]
    [InlineData(4, 8, TextUnit.Word, 1, "10-16")]
    [InlineData(6, 6, TextUnit.Word, 1, "10-16")]
    [InlineData(4, 4, TextUnit.Word, 1, "10-16")]
    [InlineData(4, 9, TextUnit.Word, 3, "19-21")]
    [InlineData(4, 9, TextUnit.Character, 1, "9-10")]
    [InlineData(4, 9, TextUnit.Paragraph, 1, "43-69")]
    [InlineData(4, 9, TextUnit.Sentence, 1, "21-43")]
    [InlineData(70, 72, TextUnit.Paragraph, 1, "null")]
    [InlineData(70, 74, TextUnit.Word, 1, "null")]
    [InlineData(4, 10, TextUnit.Word, 1, "10-16")]
    [InlineData(10, 10, TextUnit.Word, 1, "16-19")]
    [InlineData(21, 21, TextUnit.Sentence, 1, "43-69")]
    [InlineData(43, 43, TextUnit.Paragraph, 1, "69-74")]
    [InlineData(4, 4, TextUnit.Character, 1, "5-6")]
    [InlineData(0, 74, TextUnit.Paragraph, 1, "null")]
    [InlineData(43, 69, TextUnit.Paragraph, 1, "69-74")]
    [InlineData(4, 9, TextUnit.Word, 0, "null")]
    [InlineData(10, 15, TextUnit.Word, -1, "4-10")]
    [InlineData(4, 9, TextUnit.Word, 99, "null")]
    public void NextIsTheUnitAfterTheRange(int start, int end, TextUnit unit, int count, string expected)
    {
        Assert.Equal(expected, Show(At(NewBody(), start, end).Next(unit, count)));
    }

    [Theory]
    [InlineData(10, 15, TextUnit.Word, 1, "4-10")]
    [InlineData(10, 10, TextUnit.Word, 1, "4-10")]
    [InlineData(4, 4, TextUnit.Word, 1, "0-4")]
    [InlineData(10, 15, TextUnit.Word, 2, "0-4")]
    [InlineData(10, 15, TextUnit.Character, 1, "9-10")]
    [InlineData(50, 55, TextUnit.Paragraph, 1, "0-43")]
    [InlineData(0, 3, TextUnit.Word, 1, "null")]
    [InlineData(4, 9, TextUnit.Word, 1, "0-4")]
    [InlineData(6, 6, TextUnit.Word, 1, "0-4")]
    [InlineData(43, 50, TextUnit.Paragraph, 1, "0-43")]
    [InlineData(4, 4, TextUnit.Character, 1, "3-4")]
    [InlineData(4, 9, TextUnit.Word, 0, "null")]
    public void PreviousIsTheUnitBeforeTheRange(int start, int end, TextUnit unit, int count, string expected)
    {
        Assert.Equal(expected, Show(At(NewBody(), start, end).Previous(unit, count)));
    }

    [Fact]
    public void NextAndPreviousDefaultToTheCharacterAfterAndBefore()
    {
        Assert.Equal("12-13", Show(At(NewBody(), 6, 12).Next()));
        Assert.Equal("5-6", Show(At(NewBody(), 6, 12).Previous()));
    }

    [Fact]
    public void NextAndPreviousDoNotCountStories()
    {
        RangeWrapper range = At(NewBody(), 4, 9);

        Assert.Throws<ArgumentException>(() => range.Next(TextUnit.Story));
        Assert.Throws<ArgumentException>(() => range.Previous(TextUnit.Story));
    }

    [Theory]
    [InlineData("MoveUntil", "q", 100, 0, 3, 2, "4-4")]
    [InlineData("MoveUntil", "q", 100, 0, 0, 5, "4-4")]
    [InlineData("MoveUntil", "oq", 100, 0, 3, 2, "4-4")]
    [InlineData("MoveUntil", ".", 5, 0, 3, 0, "0-3")]
    [InlineData("MoveUntil", "z", 100, 0, 3, 0, "0-3")]
    [InlineData("MoveUntil", "T", -100, 10, 15, -10, "1-1")]
    [InlineData("MoveUntil", "q", -100, 0, 3, 0, "0-3")]
    [InlineData("MoveUntil", "t", 100, 0, 0, 32, "31-31")]
    [InlineData("MoveUntil", "q", 100, 4, 4, 1, "4-4")]
    [InlineData("MoveUntil", "r", 100, 40, 45, 2, "46-46")]
    [InlineData("MoveWhile", "The ", 100, 0, 0, 4, "4-4")]
    [InlineData("MoveWhile", "Theq ", 3, 0, 0, 3, "3-3")]
    [InlineData("MoveWhile", "qwertyuiopasdfghjklzxcvbnm", -100, 9, 9, -5, "4-4")]
    [InlineData("MoveWhile", "z", 100, 0, 3, 0, "0-3")]
    [InlineData("MoveWhile", "The ", 100, 0, 3, 1, "4-4")]
    [InlineData("MoveStartUntil", "q", 100, 0, 12, 5, "4-12")]
    [InlineData("MoveStartUntil", "q", 100, 0, 3, 5, "4-4")]
    [InlineData("MoveStartUntil", "T", -100, 10, 15, -10, "1-15")]
    [InlineData("MoveStartWhile", "The ", 100, 0, 12, 4, "4-12")]
    [InlineData("MoveEndUntil", ".", 100, 0, 3, 17, "0-19")]
    [InlineData("MoveEndUntil", "T", -100, 10, 15, -15, "1-1")]
    [InlineData("MoveEndUntil", "q", -100, 10, 15, -11, "5-5")]
    [InlineData("MoveEndWhile", "qui", 100, 0, 3, 0, "0-3")]
    [InlineData("MoveEndWhile", "The qui", 100, 0, 1, 6, "0-7")]
    public void MoveUntilAndWhileMoveOverCharacters(string method, string characters, int count, int start, int end, int returned, string expected)
    {
        RangeWrapper range = At(NewBody(), start, end);
        Func<string, int, int> move = (Func<string, int, int>)Delegate.CreateDelegate(typeof(Func<string, int, int>), range, typeof(RangeWrapper).GetMethod(method));

        Assert.Equal(returned, move(characters, count));
        Assert.Equal(expected, Show(range));
    }

    [Theory]
    [InlineData("q", 3, 0, "0-0")]
    [InlineData("q", 4, 0, "0-0")]
    [InlineData("q", 5, 5, "4-4")]
    [InlineData("q", 6, 5, "4-4")]
    [InlineData("T", -5, -5, "1-1")]
    [InlineData("T", -4, 0, "5-5")]
    [InlineData("T", -6, -5, "1-1")]
    public void TheCountIsHowManyCharactersAreLookedAt(string characters, int count, int returned, string expected)
    {
        RangeWrapper range = At(NewBody(), count < 0 ? 5 : 0, count < 0 ? 5 : 0);

        Assert.Equal(returned, range.MoveUntil(characters, count));
        Assert.Equal(expected, Show(range));
    }

    [Fact]
    public void MoveWhileStopsBeforeTheLastParagraphMark()
    {
        RangeWrapper range = At(NewBody(), 70, 70);

        Assert.Equal(3, range.MoveWhile("Last\r"));
        Assert.Equal("73-73", Show(range));
        Assert.Equal(0, range.MoveUntil("z"));
    }

    [Fact]
    public void MoveUntilCanStopAtAParagraphMark()
    {
        RangeWrapper range = At(NewBody(), 30, 30);

        Assert.Equal(13, range.MoveUntil("\r"));
        Assert.Equal("42-42", Show(range));
    }

    [Fact]
    public void MoveUntilWithoutACountLooksToTheEndOfTheStory()
    {
        RangeWrapper range = At(NewBody(), 6, 12);

        Assert.Equal(0, range.MoveUntil("q"));
        Assert.Equal("6-12", Show(range));
    }

    [Theory]
    [InlineData(10, 15, TextUnit.Story, 1, 1, "73-73")]
    [InlineData(10, 15, TextUnit.Story, -1, -1, "0-0")]
    [InlineData(10, 15, TextUnit.Story, 2, 1, "73-73")]
    [InlineData(10, 15, TextUnit.Story, 0, 0, "10-15")]
    [InlineData(73, 73, TextUnit.Story, 1, 0, "73-73")]
    [InlineData(0, 0, TextUnit.Story, -1, 0, "0-0")]
    [InlineData(10, 74, TextUnit.Story, 1, 1, "73-73")]
    [InlineData(5, 9, TextUnit.Story, -1, -1, "0-0")]
    public void MoveByAStoryGoesToItsStartOrEnd(int start, int end, TextUnit unit, int count, int returned, string expected)
    {
        RangeWrapper range = At(NewBody(), start, end);

        Assert.Equal(returned, range.Move(unit, count));
        Assert.Equal(expected, Show(range));
    }

    [Theory]
    [InlineData("MoveStart", 10, 15, 1, 1, "73-73")]
    [InlineData("MoveStart", 10, 15, 2, 1, "73-73")]
    [InlineData("MoveStart", 10, 15, -1, -1, "0-15")]
    [InlineData("MoveStart", 0, 5, -1, 0, "0-5")]
    [InlineData("MoveStart", 73, 74, 1, 0, "73-74")]
    [InlineData("MoveEnd", 10, 15, 1, 1, "10-74")]
    [InlineData("MoveEnd", 10, 15, 2, 1, "10-74")]
    [InlineData("MoveEnd", 10, 15, -1, -1, "0-0")]
    [InlineData("MoveEnd", 10, 74, 1, 0, "10-74")]
    public void MoveStartAndMoveEndByAStory(string method, int start, int end, int count, int returned, string expected)
    {
        RangeWrapper range = At(NewBody(), start, end);

        Assert.Equal(returned, method == "MoveStart" ? range.MoveStart(TextUnit.Story, count) : range.MoveEnd(TextUnit.Story, count));
        Assert.Equal(expected, Show(range));
    }

    [Fact]
    public void ExpandByAStoryAndWholeStoryTakeTheWholeStoryWithItsLastMark()
    {
        RangeWrapper expanded = At(NewBody(), 10, 15);
        RangeWrapper whole = At(NewBody(), 10, 12);

        expanded.Expand(TextUnit.Story);
        whole.WholeStory();

        Assert.Equal("0-74", Show(expanded));
        Assert.Equal("0-74", Show(whole));
    }

    [Fact]
    public void RangesAreComparedByStoryAndPlace()
    {
        Body body = NewBody();
        RangeWrapper quick = At(body, 4, 9);

        Assert.True(quick.IsEqual(At(body, 4, 9)));
        Assert.False(quick.IsEqual(At(body, 4, 10)));
        Assert.True(At(body, 6, 8).InRange(quick));
        Assert.False(quick.InRange(At(body, 6, 8)));
        Assert.True(quick.InRange(At(body, 4, 9)));
        Assert.False(quick.InRange(At(body, 7, 20)));
        Assert.True(quick.InStory(At(body, 0, 3)));
        Assert.False(quick.InStory(new RangeWrapper(NewBody())));
        Assert.False(quick.IsEqual(At(NewBody(), 4, 9)));
    }

    [Fact]
    public void ComparingWithNothingThrows()
    {
        RangeWrapper range = At(NewBody(), 0, 1);

        Assert.Throws<ArgumentNullException>(() => range.IsEqual(null));
        Assert.Throws<ArgumentNullException>(() => range.InRange(null));
        Assert.Throws<ArgumentNullException>(() => range.InStory(null));
    }

    [Theory]
    [InlineData(0, 74, 12, 61, 71, 2)]
    [InlineData(4, 9, 1, 5, 5, 0)]
    [InlineData(4, 10, 1, 5, 6, 0)]
    [InlineData(0, 3, 1, 3, 3, 0)]
    [InlineData(5, 5, 0, 0, 0, 0)]
    [InlineData(6, 15, 2, 8, 9, 0)]
    [InlineData(0, 45, 8, 38, 44, 1)]
    [InlineData(40, 55, 3, 12, 14, 1)]
    [InlineData(19, 20, 1, 1, 1, 0)]
    [InlineData(52, 70, 4, 14, 17, 1)]
    [InlineData(0, 43, 7, 36, 42, 1)]
    [InlineData(43, 69, 4, 21, 25, 1)]
    [InlineData(69, 74, 1, 4, 4, 0)]
    [InlineData(0, 69, 11, 57, 67, 2)]
    [InlineData(42, 44, 1, 1, 1, 0)]
    [InlineData(42, 43, 0, 0, 0, 0)]
    [InlineData(1, 74, 12, 60, 70, 2)]
    public void ComputeStatisticsCountsWhatWordCounts(int start, int end, int words, int characters, int charactersWithSpaces, int paragraphs)
    {
        RangeWrapper range = At(NewBody(), start, end);

        Assert.Equal(words, range.ComputeStatistics(Statistic.Words));
        Assert.Equal(characters, range.ComputeStatistics(Statistic.Characters));
        Assert.Equal(charactersWithSpaces, range.ComputeStatistics(Statistic.CharactersWithSpaces));
        Assert.Equal(paragraphs, range.ComputeStatistics(Statistic.Paragraphs));
    }

    [Fact]
    public void ComputeStatisticsDoesNotCountLayout()
    {
        RangeWrapper range = At(NewBody(), 0, 10);

        Assert.Throws<NotSupportedException>(() => range.ComputeStatistics(Statistic.Lines));
        Assert.Throws<NotSupportedException>(() => range.ComputeStatistics(Statistic.Pages));
        Assert.Throws<ArgumentOutOfRangeException>(() => range.ComputeStatistics((Statistic)99));
    }

    [Fact]
    public void TextUnitCollectionsHaveAFirstAndALast()
    {
        RangeWrapper range = At(NewBody(), 0, 20);

        Assert.Equal("The ", range.Words.First.Text);
        Assert.Equal(". ", range.Words.Last.Text);
        Assert.Equal("T", range.Characters.First.Text);
        Assert.Equal(".", range.Characters.Last.Text);
        Assert.Equal("The quick brown fox. ", range.Sentences.First.Text);
        Assert.Equal("The quick brown fox. ", range.Sentences.Last.Text);
    }


    [Fact]
    public void SentencesTakeTheParagraphMarksThatFollowThem()
    {
        //The sentences Word gives for the same paragraphs.
        Body body = new Body(
            new Paragraph(Plain("No period")),
            new Paragraph(Plain("Next one.")),
            new Paragraph(),
            new Paragraph(Plain("After empty. Yes!")),
            new Paragraph(Plain("End?")));

        Assert.Equal(
            new[] { "No period\r", "Next one.\r\r", "After empty. ", "Yes!\r", "End?\r" },
            At(body, 0, 44).Sentences.Select(s => s.Text).ToArray());
        Assert.Equal(new[] { "0-10", "10-21", "21-34", "34-39", "39-44" }, At(body, 0, 44).Sentences.Select(Show).ToArray());

        Body empty = new Body(new Paragraph(), new Paragraph(Plain("A")));
        Assert.Equal(new[] { "0-1", "1-3" }, At(empty, 0, 3).Sentences.Select(Show).ToArray());
        Body twice = new Body(new Paragraph(Plain("A")), new Paragraph(), new Paragraph(Plain("B")));
        Assert.Equal(new[] { "0-3", "3-5" }, At(twice, 0, 5).Sentences.Select(Show).ToArray());
    }

    [Fact]
    public void InsertParagraphReplacesTheRangeWithAParagraphMark()
    {
        Body body = NewBody();
        RangeWrapper range = At(body, 4, 9);

        range.InsertParagraph();

        Assert.Equal(new[] { "The ", " brown fox. Second sentence here.", "Third para,  with commas.", "Last" }, body.Elements<Paragraph>().Select(p => p.InnerText).ToArray());
        Assert.Equal("4-5", Show(range));
    }

    [Fact]
    public void InsertParagraphAtAnInsertionPointSplitsThere()
    {
        Body body = NewBody();
        RangeWrapper range = At(body, 4, 4);

        range.InsertParagraph();

        Assert.Equal(new[] { "The ", "quick brown fox. Second sentence here.", "Third para,  with commas.", "Last" }, body.Elements<Paragraph>().Select(p => p.InnerText).ToArray());
        Assert.Equal("4-5", Show(range));
    }
}
