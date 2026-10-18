using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The members that predate the Word-style API now share its implementation; these tests cover where they behave differently from before.
/// </summary>
public class LegacyMembersTest
{
    [Fact]
    public void ParagraphDeleteRemovesTheBookmarkAnchoredInIt()
    {
        Paragraph first = new Paragraph(new BookmarkStart() { Id = "1", Name = "mark" }, new Run(new Text("a")), new BookmarkEnd() { Id = "1" });
        Body body = new Body(first, new Paragraph(new Run(new Text("b"))));

        new ParagraphWrapper(first).Delete();

        Assert.Single(body.Elements<Paragraph>());
        Assert.Empty(body.Descendants<BookmarkStart>());
        Assert.Empty(body.Descendants<BookmarkEnd>());
    }

    [Fact]
    public void ParagraphDeleteKeepsTheOnlyParagraphOfACell()
    {
        TableCell cell = new TableCell(new Paragraph(new Run(new Text("x"))));
        Table table = new Table(new TableRow(cell));

        new ParagraphWrapper(cell.Elements<Paragraph>().Single()).Delete();

        Assert.Single(cell.Elements<Paragraph>());
        Assert.Empty(cell.Descendants<Run>());
    }

    [Fact]
    public void DeletingTheLastRowDeletesTheTable()
    {
        TableRow row = new TableRow(new TableCell(new Paragraph()));
        Table table = new Table(row);
        Body body = new Body(table, new Paragraph());

        new TableRowWrapper(row).Delete();

        Assert.Empty(body.Elements<Table>());
    }

    [Fact]
    public void SpaceBeforeReadsZeroWhenOnlySpaceAfterIsSet()
    {
        Paragraph paragraph = new Paragraph(new ParagraphProperties(new SpacingBetweenLines() { After = "120" }));
        ParagraphWrapper wrapper = new ParagraphWrapper(paragraph);

        Assert.Equal(0, wrapper.SpaceBefore);
        Assert.Equal(6, wrapper.SpaceAfter);

        wrapper.SpaceBefore = 12;
        Assert.Equal(12, wrapper.SpaceBefore);
        Assert.Equal(6, wrapper.SpaceAfter);
    }

    [Fact]
    public void KeepWithNextCanBeReadAndCleared()
    {
        ParagraphWrapper wrapper = new ParagraphWrapper(new Paragraph());
        Assert.False(wrapper.KeepWithNext);

        wrapper.KeepWithNext = true;
        Assert.True(wrapper.KeepWithNext);
        Assert.True(wrapper.Format.KeepWithNext);

        wrapper.KeepWithNext = false;
        Assert.False(wrapper.KeepWithNext);
    }

    [Fact]
    public void MergeRightAddsTheSpansAndMovesTheText()
    {
        TableCell left = new TableCell(new Paragraph(new Run(new Text("L"))));
        TableCell middle = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }), new Paragraph(new Run(new Text("M"))));
        TableCell right = new TableCell(new Paragraph());
        TableRow row = new TableRow(left, middle, right);
        Table table = new Table(row);

        new TableCellWrapper(left).MergeRight();

        Assert.Equal(2, row.Elements<TableCell>().Count());
        Assert.Equal(3, left.TableCellProperties.GridSpan.Val.Value);
        Assert.Contains("M", new TableCellWrapper(left).Range.Text);
    }

    [Fact]
    public void MergeDownFindsTheCellByGridColumnWhenAnotherCellSpans()
    {
        TableCell topSecond = new TableCell(new Paragraph());
        TableCell bottomSecond = new TableCell(new Paragraph());
        TableCell bottomFirst = new TableCell(new TableCellProperties(new GridSpan() { Val = 2 }), new Paragraph());
        Table table = new Table(
            new TableRow(new TableCell(new Paragraph()), topSecond, new TableCell(new Paragraph())),
            new TableRow(bottomFirst, bottomSecond));

        //The second top cell starts at grid column 1, which the two-column cell below covers.
        new TableCellWrapper(topSecond).MergeDown();

        Assert.NotNull(bottomFirst.TableCellProperties.VerticalMerge);
        Assert.Null(bottomSecond.TableCellProperties?.VerticalMerge);
    }

    [Theory]
    [InlineData(SectionStart.NewPage)]
    [InlineData(SectionStart.Continuous)]
    [InlineData(SectionStart.EvenPage)]
    [InlineData(SectionStart.OddPage)]
    [InlineData(SectionStart.NewColumn)]
    public void InsertSectionBreakSetsHowTheNextSectionStarts(SectionStart expected)
    {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
            {
                document.Paragraphs[0].Range.Text = "one";
                document.Paragraphs[0].InsertSectionBreak(expected);

                Assert.Equal(2, document.Sections.Count);
                Assert.Equal(expected, document.Sections[1].PageSetup.SectionStart);
            }
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void InsertSectionBreakRejectsAnUnknownType()
    {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
            {
                Assert.Throws<ArgumentOutOfRangeException>(() => document.Paragraphs[0].InsertSectionBreak((SectionStart)99));
            }
        }
        finally
        {
            File.Delete(path);
        }
    }
}
