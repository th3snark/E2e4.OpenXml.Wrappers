using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Positions of hyperlinks, content controls and tracked deletions, and edits that cross the edge of a table. The numbers are the ones
/// Word gives for the same content (found by opening it in Word).
/// </summary>
public class RangeWrapperWordPositionsTest
{
    private static Run Plain(string text)
    {
        return new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Paragraph Para(params OpenXmlElement[] content)
    {
        return new Paragraph(content);
    }

    private static int Length(Body body)
    {
        return new RangeWrapper(body).End;
    }

    [Fact]
    public void AHyperlinkToAPlaceTakesThePositionsOfItsField()
    {
        //a, begin, ' HYPERLINK \l "bk" \h' (the code Word makes up), separator, xy, end, b and the end of the paragraph.
        Assert.Equal(29, Length(new Body(Para(Plain("a"), new Hyperlink(Plain("xy")) { Anchor = "bk" }, Plain("b")))));
        Assert.Equal(27, Length(new Body(Para(Plain("a"), new Hyperlink() { Anchor = "bk" }, Plain("b")))));
    }

    [Fact]
    public void AHyperlinkWithADocumentLocationHasItInItsCode()
    {
        Assert.Equal(38, Length(new Body(Para(Plain("a"), new Hyperlink(Plain("xy")) { Anchor = "bk", DocLocation = "loc" }, Plain("b")))));
    }

    [Fact]
    public void AHyperlinkToAnAddressCountsTheAddress()
    {
        string path = Path.GetTempFileName();
        try
        {
            using WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false);
            document.InitializeMainDocumentPart();
            MainDocumentPart part = document.WordprocessingDocument.MainDocumentPart;
            string id = part.AddHyperlinkRelationship(new Uri("https://example.com/a"), true).Id;
            part.Document.Body.RemoveAllChildren();
            part.Document.Body.AppendChild(Para(Plain("a"), new Hyperlink(Plain("xy")) { Id = id }, Plain("b")));

            Assert.Equal(45, Length(part.Document.Body));

            Hyperlink link = part.Document.Body.Descendants<Hyperlink>().Single();
            link.Tooltip = "tip";
            link.TargetFrame = "_blank";
            Assert.Equal(66, Length(part.Document.Body));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ATextRangeInsideAHyperlinkStillReadsItsText()
    {
        Body body = new Body(Para(Plain("a"), new Hyperlink(Plain("xy")) { Anchor = "bk" }, Plain("b")));

        Assert.Equal("axyb\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void ARangeFollowsItsTextAfterAHyperlink()
    {
        Body body = new Body(Para(Plain("a"), new Hyperlink(Plain("xy")) { Anchor = "bk" }, Plain("b")));
        RangeWrapper b = new RangeWrapper(body);
        Assert.True(b.Find.Execute("b"));

        //a is at 0, the field takes 1 + 19 + 1 + 2 + 1 positions, so b is at 27.
        Assert.Equal(27, b.Start);
        new RangeWrapper(body).Find.Execute("a", replaceWith: "aaa", replace: ReplaceMode.One);
        Assert.Equal(29, b.Start);
    }

    [Fact]
    public void AContentControlTakesAPositionAtEachEnd()
    {
        Assert.Equal(7, Length(new Body(Para(Plain("a"), new SdtRun(new SdtContentRun(Plain("xy"))), Plain("b")))));
    }

    [Fact]
    public void TrackedDeletedTextTakesPositionsButShowsNoText()
    {
        Body body = new Body(Para(Plain("a"), new DeletedRun(new Run(new DeletedText("gone"))), new InsertedRun(Plain("new")), Plain("b")));

        Assert.Equal(10, Length(body));
        Assert.Equal("anewb\r", new RangeWrapper(body).Text);
    }

    private static Body TableBody()
    {
        return new Body(Para(Plain("x")), new Table(new TableRow(new TableCell(Para(Plain("aaa"))), new TableCell(Para(Plain("bbb"))))), Para(Plain("end")));
    }

    private static RangeWrapper Range(Body body, int start, int end)
    {
        RangeWrapper range = new RangeWrapper(body);
        range.SetRange(start, end);
        return range;
    }

    [Theory]
    [InlineData(0, 4)]
    [InlineData(8, 13)]
    [InlineData(0, 10)]
    [InlineData(3, 13)]
    public void DeletingPartOfATableAndSomethingOutsideItFailsAsInWord(int start, int end)
    {
        Body body = TableBody();

        Assert.Throws<InvalidOperationException>(() => Range(body, start, end).Delete());
        Assert.Throws<InvalidOperationException>(() => Range(body, start, end).Text = "ZZ");
        Assert.Equal("x\raaa\r\abbb\r\a\r\aend\r", new RangeWrapper(body).Text);
    }

    [Fact]
    public void DeletingTheWholeTableWithTextBeforeItRemovesTheTable()
    {
        Body body = TableBody();

        Range(body, 0, 11).Delete();

        Assert.Empty(body.Descendants<Table>());
        Assert.Equal("end\r", new RangeWrapper(body).Text);
    }
}
