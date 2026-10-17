using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Moving a range by units one block at a time gives the same places as moving through the text of the whole story.
/// </summary>
public class UnitNavigationTest
{
    private static Paragraph Paragraph(string text)
    {
        return new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Body CreateBody()
    {
        TableCell first = new TableCell(Paragraph("ab cd."), Paragraph("ef"));
        TableCell second = new TableCell(Paragraph("x. y"));
        TableCell third = new TableCell(Paragraph("one"));
        TableCell fourth = new TableCell(Paragraph("two three"));
        return new Body(
            Paragraph("The quick brown fox. It jumps over! Does it?"),
            new Paragraph(),
            new Table(new TableRow(first, second), new TableRow(third, fourth)),
            Paragraph("After the table. The end."));
    }

    //Moving through the text of the whole story: the next or previous unit boundary.
    private static int Reference(char[] text, TextUnit unit, int position, int count, out int moved)
    {
        List<int> boundaries = TextUnits.GetBoundaries(text, unit);
        moved = 0;
        int current = position;
        while (moved < System.Math.Abs(count))
        {
            int next;
            if (count > 0)
            {
                next = boundaries.FirstOrDefault(b => b > current);
                if (next <= current)
                {
                    break;
                }
            }
            else
            {
                int index = boundaries.FindLastIndex(b => b < current);
                if (index < 0)
                {
                    break;
                }
                next = boundaries[index];
            }
            current = next;
            moved++;
        }
        moved = count > 0 ? moved : -moved;
        return current;
    }

    [Theory]
    [InlineData(TextUnit.Character)]
    [InlineData(TextUnit.Word)]
    [InlineData(TextUnit.Sentence)]
    [InlineData(TextUnit.Paragraph)]
    public void MovingGivesTheSamePlacesAsMovingThroughTheWholeStory(TextUnit unit)
    {
        Body body = CreateBody();
        char[] text = StoryMap.Build(body).ToCharacters();

        foreach (int count in new[] { -5, -2, -1, 1, 2, 5 })
        {
            for (int position = 0; position <= text.Length; position++)
            {
                int expected = Reference(text, unit, position, count, out int expectedMoved);

                RangeWrapper range = new RangeWrapper(body);
                range.SetRange(position, position);
                int moved = range.Move(unit, count);

                Assert.True(expectedMoved == moved && expected == range.Start, "Move " + unit + " " + count + " from " + position + ": expected " + expected + " (" + expectedMoved + " units) but got " + range.Start + " (" + moved + ")");
            }
        }
    }

    [Fact]
    public void TheStartAndEndMovePastEachOtherAndCollapseAcrossATable()
    {
        Body body = CreateBody();
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(44, 50);

        range.MoveStart(TextUnit.Paragraph, 3);
        Assert.Equal(range.Start, range.End);

        range.SetRange(60, 70);
        range.MoveEnd(TextUnit.Paragraph, -4);
        Assert.Equal(range.Start, range.End);
    }
}
