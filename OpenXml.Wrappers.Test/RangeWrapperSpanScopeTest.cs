using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Moving, inserting paragraphs, deleting across paragraphs, and positions, in the part of the document that the range is in.
/// </summary>
public class RangeWrapperSpanScopeTest
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

    private static Body BigBody(int paragraphs)
    {
        Body body = new Body();
        for (int i = 0; i < paragraphs; i++)
        {
            body.AppendChild(Paragraph("Paragraph number " + i.ToString(CultureInfo.InvariantCulture) + " has some words in it and ends here."));
        }
        return body;
    }

    #region Moving

    [Fact]
    public void MovingTheEndPastTheEndOfTheParagraphGoesToTheNextOne()
    {
        Body body = new Body(Paragraph("one two"), Paragraph("three four"));
        RangeWrapper range = Range(body, 4, 7);

        Assert.Equal(1, range.MoveEnd(TextUnit.Word));

        Assert.Equal(8, range.End);
        Assert.Equal("two\r", range.Text);
        Assert.Equal(2, range.MoveEnd(TextUnit.Word, 2));
        Assert.Equal("two\rthree four", range.Text);
    }

    [Fact]
    public void MovingTheStartPastTheEndCollapsesTheRange()
    {
        Body body = new Body(Paragraph("one two three"));
        RangeWrapper range = Range(body, 4, 7);

        Assert.Equal(2, range.MoveStart(TextUnit.Word, 2));

        Assert.Equal(range.Start, range.End);
        Assert.Equal(13, range.Start);
    }

    [Fact]
    public void MovingTheEndBackPastTheStartCollapsesTheRange()
    {
        Body body = new Body(Paragraph("one two three"));
        RangeWrapper range = Range(body, 8, 13);

        Assert.Equal(-2, range.MoveEnd(TextUnit.Word, -2));

        Assert.Equal(range.Start, range.End);
        Assert.Equal(4, range.End);
    }

    [Fact]
    public void MovingByCharactersAcrossParagraphsStillWorks()
    {
        Body body = new Body(Paragraph("ab"), Paragraph("cd"));
        RangeWrapper range = Range(body, 2, 2);

        Assert.Equal(3, range.MoveEnd(TextUnit.Character, 3));
        Assert.Equal("\rcd", range.Text);
        Assert.Equal(-1, range.MoveStart(TextUnit.Character, -1));
        Assert.Equal("b\rcd", range.Text);
    }

    [Fact]
    public void ARangeOfElementsCanBeMoved()
    {
        Paragraph paragraph = Paragraph("one two three");
        Body body = new Body(paragraph);
        RangeWrapper range = new ParagraphWrapper(paragraph).Range;

        Assert.Equal(-4, range.Move(TextUnit.Character, -4));

        Assert.Equal(range.Start, range.End);
        Assert.Equal(10, range.Start);
    }

    #endregion

    #region Paragraphs

    [Fact]
    public void ATextWithCarriageReturnsMakesParagraphsOneByOne()
    {
        Body body = new Body(Paragraph("start end"));
        RangeWrapper range = Range(body, 6, 6);

        range.InsertAfter("one\rtwo\rthree ");

        Assert.Equal("start one\rtwo\rthree end\r", StoryText(body));
        Assert.Equal(3, body.Elements<Paragraph>().Count());
        Assert.Equal("one\rtwo\rthree ", range.Text);
    }

    [Fact]
    public void ReplacingTextWithCarriageReturnsKeepsTheRangeOnTheNewText()
    {
        Body body = new Body(Paragraph("A cat sat"));
        RangeWrapper cat = Range(body, 2, 5);

        cat.Text = "big\rred cat";

        Assert.Equal("A big\rred cat sat\r", StoryText(body));
        Assert.Equal("big\rred cat", cat.Text);
        Assert.Equal(2, body.Elements<Paragraph>().Count());
    }

    [Fact]
    public void InsertingAParagraphAfterASpanSplitsThere()
    {
        Body body = new Body(Paragraph("Hello world"));
        RangeWrapper hello = Range(body, 0, 5);

        ParagraphWrapper second = hello.InsertParagraphAfter();

        Assert.Equal(" world\r", second.Text);
        Assert.Equal("Hello\r", hello.Text);
        Assert.Equal("Hello\r world\r", StoryText(body));
    }

    [Fact]
    public void InsertingAParagraphBeforeASpanSplitsThere()
    {
        Body body = new Body(Paragraph("Hello world"));
        RangeWrapper world = Range(body, 6, 11);

        world.InsertParagraphBefore();

        Assert.Equal("Hello \rworld\r", StoryText(body));
        Assert.Equal("\rworld", world.Text);
    }

    #endregion

    #region Deleting across paragraphs

    [Fact]
    public void DeletingAcrossParagraphsInABodyOnlyNeedsThoseParagraphs()
    {
        // The work is the same however long the document is, so the time must not grow with it.
        Scaling.AssertGrowth(paragraphs =>
        {
            Body body = BigBody(paragraphs);
            RangeWrapper range = new ParagraphWrapper(body.Elements<Paragraph>().ElementAt(paragraphs / 2)).Range;
            range.Collapse();
            range.MoveEnd(TextUnit.Character, 120);

            TimeSpan took = Scaling.Time(range.Delete);

            Assert.Equal(paragraphs - 2, body.Elements<Paragraph>().Count());
            return took;
        }, maxRatio: 3);
    }

    [Fact]
    public void DeletingAcrossParagraphsInATableCellKeepsTheCell()
    {
        TableCell cell = new TableCell(Paragraph("one"), Paragraph("two"), Paragraph("three"));
        Body body = new Body(new Table(new TableRow(cell)), Paragraph("after"));
        RangeWrapper range = new TableCellWrapper(cell).Range;
        range.SetRange(2, 8);

        range.Delete();

        Assert.Single(cell.Elements<Paragraph>());
        Assert.Equal("onthree\r\a", new ParagraphWrapper(cell.Elements<Paragraph>().Single()).Text);
    }

    #endregion

    #region Windows, finds across paragraphs and moves in a large document

    [Fact]
    public void DeletingFromBeforeATableToAfterItNeedsOnlyThoseBlocks()
    {
        // The work is the same however long the document is, so the time must not grow with it.
        Scaling.AssertGrowth(paragraphs =>
        {
            Body body = BigBody(paragraphs);
            Paragraph anchor = body.Elements<Paragraph>().ElementAt(paragraphs / 2);
            anchor.InsertAfterSelf(Paragraph("after"));
            anchor.InsertAfterSelf(new Table(new TableRow(new TableCell(Paragraph("cell")))));
            anchor.InsertAfterSelf(Paragraph("before"));

            //From after "bef" to after "af": the rest of "before", its paragraph mark, the whole table, and the start of "after".
            RangeWrapper before = new RangeWrapper(body);
            before.Find.Execute("before");
            RangeWrapper after = new RangeWrapper(body);
            after.Find.Execute("after");
            RangeWrapper span = before.Duplicate();
            span.Collapse();
            span.MoveEnd(TextUnit.Character, 3);
            span.Collapse(CollapseDirection.End);
            RangeWrapper till = after.Duplicate();
            till.Collapse();
            till.MoveEnd(TextUnit.Character, 2);
            till.Collapse(CollapseDirection.End);
            span.SetRange(span.Start, till.Start);

            TimeSpan took = Scaling.Time(span.Delete);

            Assert.Equal("befter\r", new ParagraphWrapper(body.Elements<Paragraph>().ElementAt(paragraphs / 2 + 1)).Text);
            Assert.Empty(body.Descendants<Table>());
            Assert.Equal(paragraphs + 1, body.Elements<Paragraph>().Count());
            return took;
        }, maxRatio: 3);
    }
    [Fact]
    public void AFindWithAParagraphMarkFindsAndReplacesAcrossParagraphs()
    {
        Body body = new Body(Paragraph("a"), Paragraph("b a"), Paragraph("b"), Paragraph("x"));

        Assert.True(new RangeWrapper(body).Find.Execute("a^pb a^pb"));

        RangeWrapper all = new RangeWrapper(body);
        Assert.True(all.Find.Execute("a^pb", replaceWith: "X", replace: ReplaceMode.All));
        Assert.Equal("X X\rx\r", new RangeWrapper(body).Text);
        Assert.Equal(2, body.Elements<Paragraph>().Count());
    }

    [Fact]
    public void AFindWithAParagraphMarkScalesLinearlyInALargeDocument()
    {
        Scaling.AssertGrowth(paragraphs =>
        {
            Body body = BigBody(paragraphs);
            int at = paragraphs * 5 / 6;
            body.Elements<Paragraph>().ElementAt(at).InsertAfterSelf(Paragraph("zebra"));
            body.Elements<Paragraph>().ElementAt(at + 1).InsertAfterSelf(Paragraph("crossing"));
            RangeWrapper range = new RangeWrapper(body);
            bool found = false;

            TimeSpan took = Scaling.Time(() => found = range.Find.Execute("zebra^pcrossing"));

            Assert.True(found);
            Assert.Equal("zebra\rcrossing", range.Text);
            return took;
        }, maxRatio: 12);
    }

    [Fact]
    public void MovingByWordsThroughALargeDocumentStaysWithinQuadraticCost()
    {
        Scaling.AssertGrowth(paragraphs =>
        {
            int start = paragraphs * 100000 / 6000;
            RangeWrapper range = new RangeWrapper(BigBody(paragraphs));
            range.SetRange(start, start);

            TimeSpan took = Scaling.Time(() =>
            {
                for (int i = 0; i < paragraphs / 3; i++)
                {
                    range.Move(TextUnit.Word, 7);
                }
            });

            Assert.True(range.Start > start);
            return took;
        }, maxRatio: 32);
    }

    #endregion

    #region Positions

    [Fact]
    public void ReadingPositionsDoesNotBuildAMap()
    {
        Scaling.AssertGrowth(paragraphs =>
        {
            int start = paragraphs * 100000 / 6000;
            RangeWrapper far = new RangeWrapper(BigBody(paragraphs));
            far.SetRange(start, start + 20);
            int total = 0;

            TimeSpan took = Scaling.Time(() =>
            {
                for (int i = 0; i < 50; i++)
                {
                    total += far.Start + far.End;
                }
            });

            Assert.Equal(50 * (start + start + 20), total);
            return took;
        }, maxRatio: 12);
    }

    [Fact]
    public void PositionsOfRangesOfElementsInATable()
    {
        TableCell first = new TableCell(Paragraph("ab"));
        TableCell second = new TableCell(Paragraph("cde"));
        TableRow row = new TableRow(first, second);
        Body body = new Body(Paragraph("x"), new Table(row), Paragraph("y"));

        Assert.Equal(2, new TableCellWrapper(first).Range.Start);
        Assert.Equal(5, new TableCellWrapper(first).Range.End);
        Assert.Equal(5, new TableCellWrapper(second).Range.Start);
        Assert.Equal(9, new TableCellWrapper(second).Range.End);
        Assert.Equal(2, new TableRowWrapper(row).Range.Start);
        Assert.Equal(10, new TableRowWrapper(row).Range.End);
        Assert.Equal(12, new RangeWrapper(body).StoryLength);
    }

    [Fact]
    public void WordsOfTableCellsOfALargeTableScaleLinearly()
    {
        Scaling.AssertGrowth(rows =>
        {
            Table table = new Table();
            for (int i = 0; i < rows; i++)
            {
                table.AppendChild(new TableRow(new TableCell(Paragraph("cell " + i.ToString(CultureInfo.InvariantCulture) + " has words")), new TableCell(Paragraph("second cell"))));
            }
            Body body = new Body(table, Paragraph("end"));

            TimeSpan took = Scaling.Time(() =>
            {
                foreach (TableCellWrapper cell in new RangeWrapper(table).Cells)
                {
                    cell.Range.Words[0].Font.Bold = true;
                }
            });

            Assert.Equal(true, new TableCellWrapper(table.Descendants<TableCell>().Last()).Range.Words[0].Font.Bold);
            return took;
        }, maxRatio: 12, small: 750, large: 3000);
    }

    #endregion
}
