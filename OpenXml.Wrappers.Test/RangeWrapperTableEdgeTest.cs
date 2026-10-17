using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The edges of tables and paragraphs, as Word behaves (these cases were checked against Word itself).
/// </summary>
public class RangeWrapperTableEdgeTest
{
    private static Paragraph Paragraph(string text, JustificationValues? alignment = null)
    {
        Paragraph paragraph = new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        if (alignment != null)
        {
            paragraph.PrependChild(new ParagraphProperties(new Justification() { Val = alignment.Value }));
        }
        return paragraph;
    }

    private static RangeWrapper Range(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    private static Body BodyWithTable()
    {
        return new Body(Paragraph("x"), new Table(new TableRow(new TableCell(Paragraph("a")), new TableCell(Paragraph("b")))), Paragraph("after"));
    }

    [Fact]
    public void TextInsertedAtTheEndOfARowGoesAtTheStartOfTheParagraphAfterTheTable()
    {
        Body body = BodyWithTable();

        //x, the end of its paragraph, a and the end of its cell, b and the end of its cell, then the end of the row at 6.
        Range(body, 6, 6).InsertAfter("ZZ");

        Assert.Equal("x\ra\r\ab\r\a\r\aZZafter\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void AParagraphStartedAtTheEndOfARowStartsBeforeTheParagraphAfterTheTable()
    {
        Body body = BodyWithTable();

        ParagraphWrapper added = Range(body, 6, 6).InsertParagraphAfter();

        Assert.Equal(3, body.Elements<Paragraph>().Count() + 0);
        Assert.Equal("after", new ParagraphWrapper(body.Elements<Paragraph>().Last()).Text);
        Assert.NotNull(added);
    }

    [Fact]
    public void ReplacingTextAtTheEndOfARowStartsTheParagraphAfterTheTable()
    {
        Body body = BodyWithTable();

        Range(body, 6, 6).Text = "ZZ";

        Assert.EndsWith("ZZafter\r", new RangeWrapper(body).Text, System.StringComparison.Ordinal);
    }

    [Fact]
    public void JoiningTwoParagraphsKeepsTheFormattingOfTheFirst()
    {
        Body body = new Body(Paragraph("AAA", JustificationValues.Center), Paragraph("BBB", JustificationValues.Right));

        Range(body, 3, 4).Delete();

        Paragraph joined = body.Elements<Paragraph>().Single();
        Assert.Equal("AAABBB\r", new RangeWrapper(body).Text);
        Assert.Equal(JustificationValues.Center, joined.ParagraphProperties.Justification.Val.Value);
    }

    [Fact]
    public void RemovingAnEmptyFirstParagraphLeavesTheFormattingOfTheSecond()
    {
        Body body = new Body(new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center })), Paragraph("BBB", JustificationValues.Right));

        Range(body, 0, 1).Delete();

        Paragraph left = body.Elements<Paragraph>().Single();
        Assert.Equal("BBB\r", new RangeWrapper(body).Text);
        Assert.Equal(JustificationValues.Right, left.ParagraphProperties.Justification.Val.Value);
    }

    [Fact]
    public void JoiningAParagraphToAnEmptyOneKeepsTheFormattingOfTheFirst()
    {
        Body body = new Body(Paragraph("AAA", JustificationValues.Center), new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Right })));

        Range(body, 3, 4).Delete();

        Paragraph joined = body.Elements<Paragraph>().Single();
        Assert.Equal("AAA\r", new RangeWrapper(body).Text);
        Assert.Equal(JustificationValues.Center, joined.ParagraphProperties.Justification.Val.Value);
    }

    [Fact]
    public void DeletingAcrossTheEdgeOfCellsClearsWhatIsInTheRangeAndKeepsTheCells()
    {
        TableCell first = new TableCell(Paragraph("aaa"));
        TableCell second = new TableCell(Paragraph("bbb"));
        TableCell third = new TableCell(Paragraph("ccc"));
        Body body = new Body(new Table(new TableRow(first, second, third)), Paragraph("end"));

        //From the second a of the first cell to the end of the text of the second. Word itself does nothing for a range that crosses cells; here
        //what is in the range goes, and the end of the first cell, which cannot go, stays.
        Range(body, 1, 7).Delete();

        Assert.Equal(3, body.Descendants<TableCell>().Count());
        Assert.Equal("a\r\a\r\accc\r\a\r\aend\r", new RangeWrapper(body).Text);
    }
}
