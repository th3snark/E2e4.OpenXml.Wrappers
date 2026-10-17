using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class PositionIndexTest
{
    private static Paragraph Paragraph(string text)
    {
        return new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    [Fact]
    public void AnIndexGivesTheSamePositionsAsTheRanges()
    {
        TableCell cell = new TableCell(Paragraph("cell"));
        Body body = new Body(Paragraph("first"), new Table(new TableRow(cell)), Paragraph("last words"));
        RangeWrapper content = new RangeWrapper(body);
        RangeWrapper word = content.Words[3];
        RangeWrapper bookmarked = new RangeWrapper(body);
        bookmarked.SetRange(2, 4);
        RangeWrapper cellRange = new TableCellWrapper(cell).Range;
        RangeWrapper paragraph = new ParagraphWrapper(body.Elements<Paragraph>().Last()).Range;

        PositionIndex index = content.CreatePositionIndex();

        Assert.Equal(content.StoryLength, index.Length);
        foreach (RangeWrapper range in new[] { word, bookmarked, cellRange, paragraph, content })
        {
            Assert.Equal(range.Start, index.StartOf(range));
            Assert.Equal(range.End, index.EndOf(range));
        }
    }

    [Fact]
    public void AnIndexIsASnapshotThatCanBeRefreshed()
    {
        Body body = new Body(Paragraph("one two"));
        RangeWrapper two = new RangeWrapper(body);
        two.SetRange(4, 7);
        PositionIndex index = two.CreatePositionIndex();
        Assert.Equal(4, index.StartOf(two));

        new RangeWrapper(body).Words[0].InsertBefore(">>>>");

        //The range followed its text, the index still describes the story as it was built.
        Assert.Equal(8, two.Start);
        Assert.NotEqual(two.Start, index.StartOf(two));
        index.Refresh();
        Assert.Equal(8, index.StartOf(two));
    }

    [Fact]
    public void ARangeFromAnotherStoryIsRefused()
    {
        Body first = new Body(Paragraph("one"));
        Body second = new Body(Paragraph("two"));
        PositionIndex index = new RangeWrapper(first).CreatePositionIndex();

        Assert.Throws<ArgumentException>(() => index.StartOf(new RangeWrapper(second)));
        Assert.Throws<ArgumentNullException>(() => index.StartOf(null));
    }

    [Fact]
    public void ThePositionsOfThousandsOfRangesAreQuickWithAnIndex()
    {
        Body body = new Body();
        for (int i = 0; i < 6000; i++)
        {
            body.AppendChild(Paragraph("Paragraph number " + i.ToString(CultureInfo.InvariantCulture) + " has some words in it and ends here."));
        }
        RangeWrapper content = new RangeWrapper(body);
        RangeWrapper[] words = content.Paragraphs.Select(p => p.Range.Words[0]).ToArray();
        Stopwatch watch = Stopwatch.StartNew();

        PositionIndex index = content.CreatePositionIndex();
        long total = 0;
        foreach (RangeWrapper word in words)
        {
            total += index.StartOf(word);
        }

        watch.Stop();
        Assert.True(total > 0);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(3), "Took " + watch.Elapsed);
        Assert.Equal(words[4000].Start, index.StartOf(words[4000]));
    }
}
