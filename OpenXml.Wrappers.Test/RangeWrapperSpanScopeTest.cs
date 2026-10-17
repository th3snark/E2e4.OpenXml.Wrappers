using System;
using System.Diagnostics;
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

        Assert.Equal(" world", second.Text);
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
        Body body = BigBody(6000);
        RangeWrapper range = new ParagraphWrapper(body.Elements<Paragraph>().ElementAt(3000)).Range;
        range.Collapse();
        range.MoveEnd(TextUnit.Character, 120);
        Stopwatch watch = Stopwatch.StartNew();

        range.Delete();

        watch.Stop();
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(3), "Took " + watch.Elapsed);
        Assert.Equal(5998, body.Elements<Paragraph>().Count());
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
        Assert.Equal("onthree", new ParagraphWrapper(cell.Elements<Paragraph>().Single()).Text);
    }

    #endregion

    #region Positions

    [Fact]
    public void ReadingPositionsDoesNotBuildAMap()
    {
        Body body = BigBody(6000);
        RangeWrapper far = new RangeWrapper(body);
        far.SetRange(100000, 100020);
        Stopwatch watch = Stopwatch.StartNew();

        int total = 0;
        for (int i = 0; i < 50; i++)
        {
            total += far.Start + far.End;
        }

        watch.Stop();
        Assert.Equal(50 * (100000 + 100020), total);
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(5), "Took " + watch.Elapsed);
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
    public void WordsOfTableCellsOfALargeTableAreQuick()
    {
        Table table = new Table();
        for (int i = 0; i < 3000; i++)
        {
            table.AppendChild(new TableRow(new TableCell(Paragraph("cell " + i.ToString(CultureInfo.InvariantCulture) + " has words")), new TableCell(Paragraph("second cell"))));
        }
        Body body = new Body(table, Paragraph("end"));
        Stopwatch watch = Stopwatch.StartNew();

        foreach (TableCellWrapper cell in new RangeWrapper(table).Cells)
        {
            cell.Range.Words[0].Font.Bold = true;
        }

        watch.Stop();
        Assert.True(watch.Elapsed < TimeSpan.FromSeconds(8), "Took " + watch.Elapsed);
        Assert.Equal(true, new TableCellWrapper(table.Descendants<TableCell>().Last()).Range.Words[0].Font.Bold);
    }

    #endregion
}
