using System;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// A span of characters follows its text when the document changes around it, and costs about as much as the paragraph it is in.
/// </summary>
public class RangeWrapperSpanAnchorTest
{
    private static Paragraph Paragraph(params string[] runs)
    {
        return new Paragraph(runs.Select(text => new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })).ToArray());
    }

    private static RangeWrapper Range(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    private static string StoryText(Body body)
    {
        return new RangeWrapper(body).Text;
    }

    #region Following the text

    [Fact]
    public void ARangeMovesWhenTextIsInsertedBeforeIt()
    {
        Body body = new Body(Paragraph("Hello world"));
        RangeWrapper world = Range(body, 6, 11);

        Range(body, 0, 0).InsertBefore(">>> ");

        Assert.Equal("world", world.Text);
        Assert.Equal(10, world.Start);
        Assert.Equal(15, world.End);
    }

    [Fact]
    public void ARangeGrowsWhenTextIsInsertedInsideIt()
    {
        Body body = new Body(Paragraph("Hello world"));
        RangeWrapper all = Range(body, 0, 11);

        Range(body, 6, 6).InsertAfter("big ");

        Assert.Equal("Hello big world", all.Text);
    }

    [Fact]
    public void ARangeIsNotAffectedByTextInsertedAfterIt()
    {
        Body body = new Body(Paragraph("Hello world"));
        RangeWrapper hello = Range(body, 0, 5);

        Range(body, 11, 11).InsertAfter("!");

        Assert.Equal("Hello", hello.Text);
        Assert.Equal(0, hello.Start);
    }

    [Fact]
    public void ARangeFollowsTextWhenARunIsSplitByFormattingAnotherRange()
    {
        Body body = new Body(Paragraph("Hello world today"));
        RangeWrapper world = Range(body, 6, 11);
        RangeWrapper today = Range(body, 12, 17);

        Range(body, 8, 9).Font.Bold = true;
        Range(body, 1, 2).Font.Italic = true;

        Assert.Equal("world", world.Text);
        Assert.Equal(6, world.Start);
        Assert.Equal("today", today.Text);
        Assert.Equal(12, today.Start);
        Assert.Null(world.Font.Bold);
    }

    [Fact]
    public void ARangeFollowsTextWhenAnotherParagraphIsAddedOrRemovedWithoutTheLibrary()
    {
        Paragraph second = Paragraph("World");
        Body body = new Body(Paragraph("Hello"), second);
        RangeWrapper world = Range(body, 6, 11);

        //Edits made directly through the Open XML SDK, or through the other wrappers.
        body.PrependChild(Paragraph("New one"));
        Assert.Equal("World", world.Text);
        Assert.Equal(14, world.Start);

        new ParagraphWrapper(body.Elements<Paragraph>().First()).Delete();
        Assert.Equal("World", world.Text);
        Assert.Equal(6, world.Start);
    }

    [Fact]
    public void ARangeFollowsTextWhenAReplaceAllChangesTheLengthBeforeIt()
    {
        Body body = new Body(Paragraph("a cat and a cat sat"));
        RangeWrapper sat = Range(body, 16, 19);

        new RangeWrapper(body).Find.Execute("cat", replaceWith: "kitten", replace: ReplaceMode.All);

        Assert.Equal("sat", sat.Text);
        Assert.Equal("a kitten and a kitten sat\r", StoryText(body));
    }

    [Fact]
    public void ARangeFollowsTextWhenAParagraphIsSplitBeforeIt()
    {
        Body body = new Body(Paragraph("onetwo three"));
        RangeWrapper three = Range(body, 7, 12);
        RangeWrapper end = Range(body, 12, 12);

        Range(body, 3, 3).InsertAfter("\r");

        Assert.Equal("three", three.Text);
        Assert.Equal(8, three.Start);
        Assert.Equal(13, end.Start);
        Assert.Equal(2, body.Elements<Paragraph>().Count());
    }

    [Fact]
    public void ARangeWhoseTextIsDeletedByAnotherRangeCollapsesToWhereItWas()
    {
        Body body = new Body(Paragraph("Hello big world"));
        RangeWrapper big = Range(body, 6, 9);
        RangeWrapper world = Range(body, 10, 15);

        Range(body, 5, 9).Delete();

        Assert.Equal(string.Empty, big.Text);
        Assert.Equal(5, big.Start);
        Assert.Equal(5, big.End);
        Assert.Equal("world", world.Text);
        Assert.Equal("Hello world\r", StoryText(body));
    }

    [Fact]
    public void ARangeStaysWhereItIsWhenParagraphsAreJoinedAfterIt()
    {
        Body body = new Body(Paragraph("one"), Paragraph("two"), Paragraph("three"));
        RangeWrapper one = Range(body, 0, 3);
        RangeWrapper three = Range(body, 8, 13);

        Range(body, 3, 8).Delete();

        Assert.Equal("one", one.Text);
        Assert.Equal("three", three.Text);
        Assert.Equal(3, three.Start);
    }

    [Fact]
    public void FindingFromARangeThatFollowsTheTextWorksAfterEdits()
    {
        Body body = new Body(Paragraph("x a x a x"));
        RangeWrapper range = new RangeWrapper(body);
        Assert.True(range.Find.Execute("a"));
        Assert.Equal(2, range.Start);

        Range(body, 0, 0).InsertBefore("....");

        Assert.Equal("a", range.Text);
        Assert.Equal(6, range.Start);
    }

    [Fact]
    public void ARangeWhoseParagraphIsRemovedWithTheSdkCollapsesToWhereItWas()
    {
        Paragraph doomed = Paragraph("middle");
        Body body = new Body(Paragraph("first"), doomed, Paragraph("last"));
        RangeWrapper inDoomed = Range(body, 6, 12);
        RangeWrapper last = Range(body, 13, 17);

        //Removed behind the library's back.
        doomed.Remove();

        Assert.Equal(6, inDoomed.Start);
        Assert.Equal(6, inDoomed.End);
        Assert.Equal(string.Empty, inDoomed.Text);
        Assert.Equal("last", last.Text);
        Assert.Equal(6, last.Start);

        //The range goes on following the text from there.
        new ParagraphWrapper(body.Elements<Paragraph>().First()).Range.InsertBefore(">>");
        Assert.Equal(8, inDoomed.Start);
    }

    [Fact]
    public void ARangeWhoseRunIsRemovedWithTheSdkCollapsesToWhereItWas()
    {
        Paragraph paragraph = Paragraph("Hello ", "big", " world");
        Body body = new Body(paragraph);
        RangeWrapper big = Range(body, 6, 9);
        RangeWrapper world = Range(body, 10, 15);

        paragraph.Elements<Run>().ElementAt(1).Remove();

        Assert.Equal(6, big.Start);
        Assert.Equal(6, big.End);
        Assert.Equal("world", world.Text);
        Assert.Equal(7, world.Start);
    }

    #endregion

    [Fact]
    public void ARangeWhoseTextIsRemovedBehindTheLibrarysBackLandsExactlyWhenLittleWentAndNearWhenMuchDid()
    {
        Body body = new Body();
        for (int i = 0; i < 60; i++)
        {
            body.AppendChild(Paragraph("para" + i.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }
        Paragraph[] paragraphs = body.Elements<Paragraph>().ToArray();

        //"para0" to "para9" are six positions with their paragraph marks and the rest are seven. The ranges start in the text of two paragraphs.
        RangeWrapper small = new RangeWrapper(body);
        small.SetRange(60, 63);
        RangeWrapper large = new RangeWrapper(body);
        large.SetRange(270, 273);

        //Three paragraphs (six pieces of text) go from around the small one and twenty from around the large one.
        for (int i = 9; i <= 11; i++)
        {
            paragraphs[i].Remove();
        }
        for (int i = 30; i <= 49; i++)
        {
            paragraphs[i].Remove();
        }

        //Little went, so the range is exactly where the text was: the start of "para9".
        Assert.Equal(string.Empty, small.Text);
        Assert.Equal(54, small.Start);
        Assert.Equal(54, small.End);

        //A lot went, so the range is near where it was, which is the start of the old "para30" (now where "para50" is), 200 positions in:
        //between the end of the last piece of text remembered before it and the start of the first remembered after it.
        Assert.Equal(string.Empty, large.Text);
        Assert.InRange(large.Start, 100, 242);
        Assert.Equal(large.Start, large.End);
    }

    [Fact]
    public void RangesCanBeMadeAndEditedFromSeveralThreadsOnSeveralDocuments()
    {
        System.Threading.Tasks.Parallel.For(0, 8, document =>
        {
            Body body = new Body(Paragraph("one two three four five six seven eight nine ten"), Paragraph("second paragraph here"));
            RangeWrapper content = new RangeWrapper(body);
            for (int i = 0; i < 40; i++)
            {
                RangeWrapper word = content.Words[i % 10];
                word.Font.Bold = true;
                word.InsertAfter("!");
                Assert.EndsWith("!", word.Text, StringComparison.Ordinal);
            }
            Assert.Equal(40, new RangeWrapper(body).Text.Count(c => c == '!'));
        });
    }
    [Fact]
    public void RangesCanBeMadeFromSeveralThreadsOnOneDocument()
    {
        Body body = new Body(Paragraph("alpha beta gamma delta epsilon zeta eta theta iota kappa"));
        System.Threading.Tasks.Parallel.For(0, 200, i =>
        {
            RangeWrapper range = new RangeWrapper(body);
            int start = i % 40;
            range.SetRange(start, start + 5);
            Assert.Equal(5, range.End - range.Start);
            Assert.Equal(start, range.Start);
            Assert.Equal(5, range.Text.Length);
        });
    }
    #region Cost

    private static Body BigBody(int paragraphs)
    {
        Body body = new Body();
        for (int i = 0; i < paragraphs; i++)
        {
            body.AppendChild(Paragraph("Paragraph number " + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + " has some words in it and ends here."));
        }
        return body;
    }

    [Fact]
    public void FormattingTheFirstWordOfEveryParagraphOfALargeDocumentIsFast()
    {
        Body body = BigBody(6000);
        RangeWrapper content = new RangeWrapper(body);
        Stopwatch watch = Stopwatch.StartNew();

        foreach (ParagraphWrapper paragraph in content.Paragraphs)
        {
            paragraph.Range.Words[0].Font.Bold = true;
        }

        watch.Stop();
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(8), "Took " + watch.Elapsed);
        Assert.Equal(true, new ParagraphWrapper(body.Elements<Paragraph>().Last()).Range.Words[0].Font.Bold);
        Assert.Equal(false, new ParagraphWrapper(body.Elements<Paragraph>().Last()).Range.Words[1].Font.Bold);
    }

    [Fact]
    public void LoopingFindThroughALargeDocumentIsFast()
    {
        Body body = BigBody(6000);
        RangeWrapper range = new RangeWrapper(body);
        int count = 0;
        Stopwatch watch = Stopwatch.StartNew();

        while (range.Find.Execute("words"))
        {
            range.Font.Italic = true;
            count++;
        }

        watch.Stop();
        Assert.Equal(6000, count);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(8), "Took " + watch.Elapsed);
    }

    [Fact]
    public void WalkingAParagraphCollectionByIndexIsFast()
    {
        Body body = BigBody(6000);
        ParagraphWrapperCollection paragraphs = new RangeWrapper(body).Paragraphs;
        Stopwatch watch = Stopwatch.StartNew();

        int length = 0;
        for (int i = 0; i < paragraphs.Count; i++)
        {
            length += paragraphs[i].Text.Length;
        }

        watch.Stop();
        Assert.True(length > 6000 * 40);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(8), "Took " + watch.Elapsed);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraphs[6000]);
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraphs[-1]);
    }

    [Fact]
    public void ReplacingAcrossALargeDocumentIsFast()
    {
        Body body = BigBody(6000);
        Stopwatch watch = Stopwatch.StartNew();

        new RangeWrapper(body).Find.Execute("words", replaceWith: "terms", replace: ReplaceMode.All);

        watch.Stop();
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(8), "Took " + watch.Elapsed);
        Assert.Contains("some terms in it", StoryText(body), StringComparison.Ordinal);
    }

    #endregion
}
