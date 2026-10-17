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
    public void ASpanPastADeletedParagraphUsesItsLastPosition()
    {
        Paragraph doomed = Paragraph("middle");
        Body body = new Body(Paragraph("first"), doomed, Paragraph("last"));
        RangeWrapper inDoomed = Range(body, 6, 12);

        //Removed behind the library's back, so it cannot be followed: the range keeps its last position instead of failing.
        doomed.Remove();

        Assert.True(inDoomed.Start >= 0);
        Assert.True(inDoomed.End >= inDoomed.Start);
        Assert.NotNull(inDoomed.Text);
    }

    #endregion

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
