using System;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

public class TableStructureTest
{
    private static Table CreateTable(int rows, int columns)
    {
        Table table = new Table(new TableProperties(), new TableGrid(Enumerable.Range(0, columns).Select(i => new GridColumn() { Width = "1000" })));
        for (int r = 0; r < rows; r++)
        {
            table.AppendChild(new TableRow(Enumerable.Range(0, columns).Select(c => new TableCell(
                new TableCellProperties(new TableCellWidth() { Width = "1000", Type = TableWidthUnitValues.Dxa }),
                new Paragraph(new Run(new Text(r + "," + c)))))));
        }
        return table;
    }

    private static TableCellWrapper Cell(Table table, int row, int column)
    {
        return new TableCellWrapper(table.Elements<TableRow>().ElementAt(row).Elements<TableCell>().ElementAt(column));
    }

    private static string[] Texts(Table table, int row)
    {
        return table.Elements<TableRow>().ElementAt(row).Elements<TableCell>().Select(c => c.InnerText).ToArray();
    }

    #region Merge

    [Fact]
    public void MergingAcrossARowSetsTheSpanAndWidthAndKeepsTheText()
    {
        Table table = CreateTable(1, 3);
        Cell(table, 0, 0).Merge(Cell(table, 0, 2));

        TableCell merged = table.Descendants<TableCell>().Single();
        Assert.Equal(3, merged.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal("3000", merged.TableCellProperties.TableCellWidth.Width.Value);
        Assert.Equal("0,00,10,2", merged.InnerText);
        Assert.Equal(3, merged.Elements<Paragraph>().Count());
    }

    [Fact]
    public void MergingAColumnDownSetsTheVerticalMerge()
    {
        Table table = CreateTable(3, 2);
        Cell(table, 0, 1).Merge(Cell(table, 2, 1));

        Assert.Equal(MergedCellValues.Restart, Cell(table, 0, 1).TableCell.TableCellProperties.VerticalMerge.Val.Value);
        Assert.Null(Cell(table, 1, 1).TableCell.TableCellProperties.VerticalMerge.Val);
        Assert.Null(Cell(table, 2, 1).TableCell.TableCellProperties.VerticalMerge.Val);
        Assert.Equal("0,11,12,1", Cell(table, 0, 1).TableCell.InnerText);
        Assert.Equal(String.Empty, Cell(table, 1, 1).TableCell.InnerText);
        Assert.Single(Cell(table, 1, 1).TableCell.Elements<Paragraph>());
    }

    [Fact]
    public void MergingARectangleWorksFromEitherCorner()
    {
        Table table = CreateTable(3, 3);
        Cell(table, 2, 2).Merge(Cell(table, 1, 1));

        Assert.Equal(new[] { "0,00,10,2" }, new[] { Texts(table, 0).Aggregate(String.Concat) });
        Assert.Equal(2, Texts(table, 1).Length);
        Assert.Equal(2, Cell(table, 1, 1).TableCell.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal(MergedCellValues.Restart, Cell(table, 1, 1).TableCell.TableCellProperties.VerticalMerge.Val.Value);
        Assert.Equal(2, Cell(table, 2, 1).TableCell.TableCellProperties.GridSpan.Val.Value);
        Assert.NotNull(Cell(table, 2, 1).TableCell.TableCellProperties.VerticalMerge);
    }

    [Fact]
    public void MergeUsesTheGridWhenRowsHaveDifferentSpans()
    {
        Table table = CreateTable(2, 3);
        Cell(table, 0, 0).Merge(Cell(table, 0, 1));

        //Row 0 is now [span 2][c2]; merging its first cell with the second row's first cell must cut through nothing.
        Cell(table, 0, 0).Merge(Cell(table, 1, 1));

        Assert.Single(table.Elements<TableRow>().First().Elements<TableCell>(), c => c.TableCellProperties.GridSpan?.Val.Value == 2);
        Assert.Equal(2, table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().Count());
    }

    [Fact]
    public void MergeRejectsACutThroughAMergedCell()
    {
        Table table = CreateTable(2, 3);
        Cell(table, 0, 0).Merge(Cell(table, 0, 1));

        Assert.Throws<NotSupportedException>(() => Cell(table, 0, 1).Merge(Cell(table, 1, 1)));
        Assert.Equal(3, table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().Count(c => c.TableCellProperties.GridSpan == null));
    }

    [Fact]
    public void MergeRejectsCellsFromDifferentTablesAndIgnoresTheSameCell()
    {
        Table one = CreateTable(1, 2);
        Table two = CreateTable(1, 2);

        Assert.Throws<ArgumentException>(() => Cell(one, 0, 0).Merge(Cell(two, 0, 0)));
        Cell(one, 0, 0).Merge(Cell(one, 0, 0));
        Assert.Equal(2, one.Descendants<TableCell>().Count());
    }

    [Fact]
    public void MergeAbsorbsACellThatIsAlreadyVerticallyMerged()
    {
        Table table = CreateTable(3, 1);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Cell(table, 0, 0).Merge(Cell(table, 2, 0));

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(MergedCellValues.Restart, rows[0].Elements<TableCell>().Single().TableCellProperties.VerticalMerge.Val.Value);
        Assert.Null(rows[1].Elements<TableCell>().Single().TableCellProperties.VerticalMerge.Val);
        Assert.Null(rows[2].Elements<TableCell>().Single().TableCellProperties.VerticalMerge.Val);
    }

    [Fact]
    public void MergeCutsAVerticalMergeThatSticksOutOfTheRectangle()
    {
        Table table = CreateTable(3, 1);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Cell(table, 1, 0).Merge(Cell(table, 2, 0));

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Null(rows[0].Elements<TableCell>().Single().TableCellProperties.VerticalMerge);
        Assert.Equal(MergedCellValues.Restart, rows[1].Elements<TableCell>().Single().TableCellProperties.VerticalMerge.Val.Value);
        Assert.NotNull(rows[2].Elements<TableCell>().Single().TableCellProperties.VerticalMerge);
    }

    [Fact]
    public void AVerticalMergeThatSticksOutBelowKeepsTheRestOfItself()
    {
        Table table = CreateTable(4, 2);
        Cell(table, 0, 0).Merge(Cell(table, 3, 0));

        Cell(table, 0, 0).Merge(Cell(table, 1, 1));

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(MergedCellValues.Restart, rows[2].Elements<TableCell>().First().TableCellProperties.VerticalMerge.Val.Value);
        Assert.NotNull(rows[3].Elements<TableCell>().First().TableCellProperties.VerticalMerge);
        Assert.Single(rows[0].Elements<TableCell>());
    }

    [Fact]
    public void MergeAbsorbsAVerticalMergeInsideALargerRectangle()
    {
        Table table = CreateTable(2, 3);
        Cell(table, 0, 1).Merge(Cell(table, 1, 1));

        Cell(table, 0, 0).Merge(Cell(table, 1, 2));

        Assert.Equal(2, table.Elements<TableRow>().Count());
        Assert.All(table.Elements<TableRow>(), row => Assert.Single(row.Elements<TableCell>()));
        Assert.Equal(3, table.Elements<TableRow>().First().Elements<TableCell>().Single().TableCellProperties.GridSpan.Val.Value);
    }

    #endregion

    #region Split

    [Fact]
    public void SplittingIntoColumnsDividesTheGridColumnAndWidensOtherRows()
    {
        Table table = CreateTable(2, 2);
        Cell(table, 0, 0).Split(1, 2);

        Assert.Equal(new[] { "500", "500", "1000" }, table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(g => g.Width.Value).ToArray());
        Assert.Equal(new[] { "0,0", String.Empty, "0,1" }, Texts(table, 0));
        Assert.Equal(new[] { "500", "500", "1000" }, table.Elements<TableRow>().First().Elements<TableCell>().Select(c => c.TableCellProperties.TableCellWidth.Width.Value).ToArray());

        TableCell other = table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().First();
        Assert.Equal(2, other.TableCellProperties.GridSpan.Val.Value);
        Assert.Equal("1000", other.TableCellProperties.TableCellWidth.Width.Value);
        Assert.Null(table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().Last().TableCellProperties.GridSpan);
    }

    [Fact]
    public void SplittingAnOddWidthGivesTheRemainderToTheLastCell()
    {
        Table table = CreateTable(1, 1);
        Cell(table, 0, 0).Split(1, 3);

        Assert.Equal(new[] { "333", "333", "334" }, table.Descendants<TableCell>().Select(c => c.TableCellProperties.TableCellWidth.Width.Value).ToArray());
        Assert.Equal(3, table.GetFirstChild<TableGrid>().Elements<GridColumn>().Count());
    }

    [Fact]
    public void SplittingAMergedCellBackIntoColumnsRestoresTheSpans()
    {
        Table table = CreateTable(2, 3);
        Cell(table, 0, 0).Merge(Cell(table, 0, 2));
        Cell(table, 0, 0).Split(1, 3);

        Assert.Equal(3, table.Elements<TableRow>().First().Elements<TableCell>().Count());
        Assert.All(table.Elements<TableRow>().First().Elements<TableCell>(), c => Assert.Null(c.TableCellProperties.GridSpan));
        Assert.Equal(3, table.GetFirstChild<TableGrid>().Elements<GridColumn>().Count());
    }

    [Fact]
    public void ASpanIsOnlySplitIntoAFactor()
    {
        Table table = CreateTable(1, 3);
        Cell(table, 0, 0).Merge(Cell(table, 0, 2));

        Assert.Throws<NotSupportedException>(() => Cell(table, 0, 0).Split(1, 2));
        Assert.Single(table.Descendants<TableCell>());
    }

    [Fact]
    public void SplittingIntoRowsAddsRowsAndMergesTheOtherCells()
    {
        Table table = CreateTable(2, 2);
        Cell(table, 0, 0).Split(2, 1);

        Assert.Equal(3, table.Elements<TableRow>().Count());
        TableCell kept = Cell(table, 0, 1).TableCell;
        Assert.Equal(MergedCellValues.Restart, kept.TableCellProperties.VerticalMerge.Val.Value);
        Assert.NotNull(Cell(table, 1, 1).TableCell.TableCellProperties.VerticalMerge);
        Assert.Null(Cell(table, 1, 0).TableCell.TableCellProperties.VerticalMerge);
        Assert.Equal(String.Empty, Cell(table, 1, 0).TableCell.InnerText);
        Assert.Equal("0,0", Cell(table, 0, 0).TableCell.InnerText);
        Assert.Equal("1,0", Cell(table, 2, 0).TableCell.InnerText);
    }

    [Fact]
    public void SplittingIntoRowsAndColumnsDoesBoth()
    {
        Table table = CreateTable(1, 2);
        Cell(table, 0, 0).Split(2, 2);

        Assert.Equal(2, table.Elements<TableRow>().Count());
        Assert.Equal(3, table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().Count());
        Assert.Equal(new[] { "500", "500", "1000" }, table.Elements<TableRow>().ElementAt(1).Elements<TableCell>().Select(c => c.TableCellProperties.TableCellWidth.Width.Value).ToArray());
        Assert.NotNull(Cell(table, 1, 2).TableCell.TableCellProperties.VerticalMerge);
        Assert.Null(Cell(table, 1, 0).TableCell.TableCellProperties.VerticalMerge);
    }

    [Fact]
    public void SplitRejectsBadCountsAndMergedCellsIntoRowsAndColumns()
    {
        Table table = CreateTable(3, 2);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Assert.Throws<NotSupportedException>(() => Cell(table, 0, 0).Split(2, 2));
        Assert.Throws<ArgumentOutOfRangeException>(() => Cell(table, 2, 0).Split(0, 1));
        Cell(table, 2, 0).Split(1, 1);
        Assert.Equal(3, table.Elements<TableRow>().Count());
    }

    [Fact]
    public void AVerticallyMergedCellIsSplitIntoColumnsInEveryRowOfTheMerge()
    {
        Table table = CreateTable(3, 2);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Cell(table, 0, 0).Split(1, 2);

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(3, table.GetFirstChild<TableGrid>().Elements<GridColumn>().Count());
        Assert.Equal(3, rows[0].Elements<TableCell>().Count());
        Assert.Equal(3, rows[1].Elements<TableCell>().Count());
        Assert.Equal(2, rows[2].Elements<TableCell>().Count());
        Assert.Equal(2, rows[2].Elements<TableCell>().First().TableCellProperties.GridSpan.Val.Value);
        Assert.All(rows[0].Elements<TableCell>().Take(2), c => Assert.Equal(MergedCellValues.Restart, c.TableCellProperties.VerticalMerge.Val.Value));
        Assert.All(rows[1].Elements<TableCell>().Take(2), c => Assert.NotNull(c.TableCellProperties.VerticalMerge));
    }

    [Fact]
    public void ACellIsSplitIntoRowsNextToACellThatIsAlreadyMergedVertically()
    {
        Table table = CreateTable(3, 2);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Cell(table, 0, 1).Split(2, 1);

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(4, rows.Length);
        Assert.Equal(MergedCellValues.Restart, rows[0].Elements<TableCell>().First().TableCellProperties.VerticalMerge.Val.Value);
        Assert.NotNull(rows[1].Elements<TableCell>().First().TableCellProperties.VerticalMerge);
        Assert.NotNull(rows[2].Elements<TableCell>().First().TableCellProperties.VerticalMerge);
    }

    [Fact]
    public void AMergedCellSplitIntoAsManyRowsAsItHasIsUnmerged()
    {
        Table table = CreateTable(3, 2);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Cell(table, 0, 0).Split(2, 1);

        Assert.Equal(3, table.Elements<TableRow>().Count());
        Assert.All(table.Elements<TableRow>(), r => Assert.Null(r.Elements<TableCell>().First().TableCellProperties?.VerticalMerge));
    }

    [Fact]
    public void AMergedCellSplitIntoMoreRowsGetsNewRowsAndTheOthersRunThroughThem()
    {
        Table table = CreateTable(2, 2);
        Cell(table, 0, 0).Merge(Cell(table, 1, 0));

        Cell(table, 0, 0).Split(3, 1);

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(3, rows.Length);
        Assert.All(rows, r => Assert.Null(r.Elements<TableCell>().First().TableCellProperties?.VerticalMerge));
        Assert.Equal(MergedCellValues.Restart, rows[1].Elements<TableCell>().Last().TableCellProperties.VerticalMerge.Val.Value);
        Assert.NotNull(rows[2].Elements<TableCell>().Last().TableCellProperties.VerticalMerge);
    }

    [Fact]
    public void AMergedCellSplitIntoFewerRowsSharesTheRowsOut()
    {
        Table table = CreateTable(3, 1);
        Cell(table, 0, 0).Merge(Cell(table, 2, 0));

        Cell(table, 0, 0).Split(2, 1);

        TableRow[] rows = table.Elements<TableRow>().ToArray();
        Assert.Equal(MergedCellValues.Restart, rows[0].Elements<TableCell>().Single().TableCellProperties.VerticalMerge.Val.Value);
        Assert.NotNull(rows[1].Elements<TableCell>().Single().TableCellProperties.VerticalMerge);
        Assert.Null(rows[2].Elements<TableCell>().Single().TableCellProperties.VerticalMerge);
    }

    #endregion

    #region Rows.Add

    [Fact]
    public void ARowCanBeAddedBeforeAnotherRow()
    {
        Table table = CreateTable(2, 2);
        TableRowWrapperCollection rows = new TableWrapper(table).Rows;

        TableRowWrapper added = rows.Add(rows[1]);

        Assert.Equal(1, added.Index);
        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { String.Empty, String.Empty }, Texts(table, 1));
        Assert.Equal(new[] { "1,0", "1,1" }, Texts(table, 2));
    }

    [Fact]
    public void ARowAddedBeforeTheFirstRowIsFirst()
    {
        Table table = CreateTable(1, 1);
        TableRowWrapperCollection rows = new TableWrapper(table).Rows;
        rows.Add(rows[0]);

        Assert.Equal(String.Empty, Texts(table, 0)[0]);
        Assert.Equal("0,0", Texts(table, 1)[0]);
    }

    [Fact]
    public void AddingBeforeNullAppends()
    {
        Table table = CreateTable(1, 1);
        TableRowWrapperCollection rows = new TableWrapper(table).Rows;
        rows.Add((TableRowWrapper)null);

        Assert.Equal(2, rows.Count);
        Assert.Equal("0,0", Texts(table, 0)[0]);
    }

    [Fact]
    public void AddingBeforeARowOfAnotherTableIsRejected()
    {
        Table table = CreateTable(1, 1);
        Table other = CreateTable(1, 1);

        Assert.Throws<ArgumentException>(() => new TableWrapper(table).Rows.Add(new TableWrapper(other).Rows[0]));
        Assert.Single(table.Elements<TableRow>());
    }

    #endregion

    #region Tables.Add

    private static (Body, Paragraph) CreateBody(uint pageWidth = 12240, uint left = 1440, uint right = 1440)
    {
        Paragraph paragraph = new Paragraph(new Run(new Text("p")));
        Body body = new Body(paragraph, new SectionProperties(new PageSize() { Width = pageWidth }, new PageMargin() { Left = left, Right = right }));
        return (body, paragraph);
    }

    [Fact]
    public void ATableIsAddedBeforeAParagraphWithTheRightShape()
    {
        (Body body, Paragraph paragraph) = CreateBody();
        TableWrapper table = new TableWrapperCollection(new RangeWrapper(body)).Add(new RangeWrapper(paragraph), 2, 3);

        Assert.Same(paragraph, table.Table.NextSibling());
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(3, table.Columns.Count);
        Assert.Equal(6, table.Table.Descendants<TableCell>().Count());
        Assert.All(table.Table.Descendants<TableCell>(), c => Assert.Single(c.Elements<Paragraph>()));
        Assert.Equal(new[] { "3120", "3120", "3120" }, table.Table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(g => g.Width.Value).ToArray());
        Assert.Equal(468, table.PreferredWidth);
        Assert.Equal(BorderLineStyle.SingleLine, table.Borders[BorderPosition.Vertical].LineStyle);
        Assert.Equal(0.5, table.Borders[BorderPosition.Top].LineWidth);
    }

    [Fact]
    public void TheTableWidthFollowsThePageSetup()
    {
        (Body body, Paragraph paragraph) = CreateBody(pageWidth: 11906, left: 1000, right: 906);
        TableWrapper table = new TableWrapperCollection(new RangeWrapper(body)).Add(new RangeWrapper(paragraph), 1, 1);

        Assert.Equal(10000 / 20.0, table.PreferredWidth);
    }

    [Fact]
    public void ADocumentRangeAddsTheTableAtTheStart()
    {
        (Body body, Paragraph paragraph) = CreateBody();
        TableWrapper table = new TableWrapperCollection(new RangeWrapper(body)).Add(new RangeWrapper(body), 1, 1);

        Assert.Same(table.Table, body.FirstChild);
        Assert.Same(paragraph, table.Table.NextSibling());
    }

    [Fact]
    public void AnEmptyBodyGetsAParagraphAfterTheTable()
    {
        Body body = new Body(new SectionProperties());
        TableWrapper table = new TableWrapperCollection(new RangeWrapper(body)).Add(new RangeWrapper(body), 1, 1);

        Assert.IsType<Paragraph>(table.Table.NextSibling());
        Assert.IsType<SectionProperties>(body.LastChild);
    }

    [Fact]
    public void ATableInACellFillsTheCellAndGoesBeforeItsFirstParagraph()
    {
        Table outer = CreateTable(1, 1);
        TableCell cell = outer.Descendants<TableCell>().Single();
        cell.TableCellProperties.TableCellWidth.Width = "4000";

        TableWrapper inner = new TableWrapperCollection(new RangeWrapper(outer)).Add(new RangeWrapper(cell), 1, 2);

        Assert.Same(inner.Table, cell.Elements().First(e => !(e is TableCellProperties)));
        Assert.IsType<Paragraph>(inner.Table.NextSibling());
        Assert.Equal(200, inner.PreferredWidth);
        Assert.Equal(new TableWrapper(outer).NestingLevel + 1, inner.NestingLevel);
    }

    [Fact]
    public void ABookmarkRangeAddsTheTableBeforeItsParagraph()
    {
        BookmarkStart start = new BookmarkStart() { Id = "1", Name = "b" };
        Paragraph paragraph = new Paragraph(start, new Run(new Text("x")), new BookmarkEnd() { Id = "1" });
        Body body = new Body(paragraph);

        TableWrapper table = new TableWrapperCollection(new RangeWrapper(body)).Add(new BookmarkWrapper(start).Range, 1, 1);

        Assert.Same(paragraph, table.Table.NextSibling());
    }

    [Fact]
    public void BadArgumentsAndRangesAreRejected()
    {
        (Body body, Paragraph paragraph) = CreateBody();
        TableWrapperCollection tables = new TableWrapperCollection(new RangeWrapper(body));

        Assert.Throws<ArgumentOutOfRangeException>(() => tables.Add(new RangeWrapper(paragraph), 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => tables.Add(new RangeWrapper(paragraph), 1, 0));
        Assert.Throws<NotSupportedException>(() => tables.Add(new RangeWrapper(paragraph.Elements<Run>().First()), 1, 1));
        Assert.Empty(body.Descendants<Table>());
    }

    #endregion
}
