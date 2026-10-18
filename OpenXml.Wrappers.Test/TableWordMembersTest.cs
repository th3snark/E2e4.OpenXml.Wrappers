using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// The members of tables, rows, columns, cells, borders and shading that follow Word's object model.
/// </summary>
public class TableWordMembersTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableWordMembersTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
        _document.Content.Text = "Hello world";
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private void AssertValid()
    {
        _document.Save();
        string[] errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2010).Validate(_document.WordprocessingDocument).Select(e => e.Description + " at " + e.Path?.XPath).ToArray();
        Assert.True(errors.Length == 0, String.Join("; ", errors));
    }

    #region Tables.Add

    [Fact]
    public void AnAddedTableReplacesTheTextOfAParagraphRange()
    {
        TableWrapper table = _document.Tables.Add(_document.Paragraphs[0].Range, 2, 2);

        Assert.Same(table.Table, _document.WordprocessingDocument.MainDocumentPart.Document.Body.FirstChild);
        Assert.Equal("", _document.Paragraphs.Last.Text);
        Assert.DoesNotContain("Hello", _document.Content.Text);
        AssertValid();
    }

    [Fact]
    public void AnAddedTableAtACollapsedSpanSplitsTheParagraph()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(5, 5), 1, 2);

        Assert.Equal("Hello", ((Paragraph)table.Table.PreviousSibling()).InnerText);
        Assert.Equal(" world", ((Paragraph)table.Table.NextSibling()).InnerText);
        AssertValid();
    }

    [Fact]
    public void AnAddedTableAtTheStartOfAParagraphDoesNotLeaveAnEmptyParagraphBeforeIt()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 1);

        Assert.Same(table.Table, _document.WordprocessingDocument.MainDocumentPart.Document.Body.FirstChild);
        Assert.Equal("Hello world", ((Paragraph)table.Table.NextSibling()).InnerText);
    }

    [Fact]
    public void AnAddedTableReplacesTheTextOfASpan()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(5, 11), 1, 1);

        Assert.Equal("Hello", ((Paragraph)table.Table.PreviousSibling()).InnerText);
        AssertValid();
    }

    [Fact]
    public void AnAddedTableEmptiesATableCellRange()
    {
        TableWrapper outer = _document.Tables.Add(_document.Range(0, 0), 1, 1);
        outer.Cell(0, 0).Range.Text = "cell text";

        TableWrapper inner = _document.Tables.Add(outer.Cell(0, 0).Range, 1, 2);

        Assert.Equal(1, outer.Cell(0, 0).Tables.Count);
        Assert.Equal(inner, outer.Cell(0, 0).Tables[0]);
        Assert.DoesNotContain("cell text", outer.Cell(0, 0).Range.Text);
        Assert.Equal(1, inner.NestingLevel - 1);
        AssertValid();
    }

    [Fact]
    public void ATableIsFixedUnlessItIsToResize()
    {
        TableWrapper fixedTable = _document.Tables.Add(_document.Range(0, 0), 1, 2);
        Assert.False(fixedTable.AllowAutoFit);

        TableWrapper resizing = _document.Tables.Add(_document.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior);
        Assert.True(resizing.AllowAutoFit);

        TableWrapper window = _document.Tables.Add(_document.Range(0, 0), 1, 2, DefaultTableBehavior.Word8TableBehavior, AutoFitBehavior.Window);
        Assert.Equal(PreferredWidthType.Percent, window.PreferredWidthType);
        Assert.Equal(100, window.PreferredWidth);
        Assert.Equal(PreferredWidthType.Percent, window.Cell(0, 0).PreferredWidthType);
        Assert.Equal(50, window.Cell(0, 0).PreferredWidth);

        TableWrapper content = _document.Tables.Add(_document.Range(0, 0), 1, 2, DefaultTableBehavior.Word8TableBehavior, AutoFitBehavior.Content);
        Assert.Equal(PreferredWidthType.Auto, content.PreferredWidthType);
        Assert.True(content.AllowAutoFit);
        AssertValid();
    }

    #endregion

    #region Table

    [Fact]
    public void PaddingSpacingAndIndentAreInPoints()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);
        Assert.Equal(5.4, table.LeftPadding);
        Assert.Equal(0, table.TopPadding);
        Assert.Equal(0, table.Spacing);

        table.TopPadding = 2;
        table.BottomPadding = 3;
        table.LeftPadding = 4;
        table.RightPadding = 6;
        table.Spacing = 1.5;
        table.LeftIndent = 10;

        Assert.Equal(2, table.TopPadding);
        Assert.Equal(3, table.BottomPadding);
        Assert.Equal(4, table.LeftPadding);
        Assert.Equal(6, table.RightPadding);
        Assert.Equal(1.5, table.Spacing);
        Assert.Equal(10, table.Rows.LeftIndent);
        Assert.Equal(4, table.Cell(0, 0).LeftPadding);
        AssertValid();

        table.Cell(0, 0).LeftPadding = 9;
        Assert.Equal(9, table.Cell(0, 0).LeftPadding);
        Assert.Equal(4, table.Cell(0, 1).LeftPadding);
        Assert.Throws<ArgumentOutOfRangeException>(() => table.TopPadding = -1);
    }

    [Fact]
    public void TheLookSwitchesKeepTheHexadecimalValueInStep()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);
        Assert.True(table.ApplyStyleRowBands);
        Assert.False(table.ApplyStyleHeadingRows);

        table.ApplyStyleHeadingRows = true;
        table.ApplyStyleFirstColumn = true;
        table.ApplyStyleRowBands = false;

        Assert.True(table.ApplyStyleHeadingRows);
        Assert.True(table.ApplyStyleFirstColumn);
        Assert.False(table.ApplyStyleLastRow);
        Assert.False(table.ApplyStyleRowBands);
        Assert.True(table.ApplyStyleColumnBands);
        Assert.Equal("02A0", table.Table.GetFirstChild<TableProperties>().TableLook.Val.Value);
        AssertValid();
    }

    [Fact]
    public void TitleAndDescriptionAreStoredAndRemovedWhenEmpty()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 1);
        Assert.Equal("", table.Title);

        table.Title = "Sales";
        table.Descr = "Sales by region";
        Assert.Equal("Sales", table.Title);
        Assert.Equal("Sales by region", table.Descr);
        AssertValid();

        table.Title = "";
        Assert.Null(table.Table.GetFirstChild<TableProperties>().TableCaption);
    }

    [Fact]
    public void ATableIsUniformWhenEveryRowHasTheSameUnmergedCells()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 3);
        Assert.True(table.Uniform);

        table.Cell(0, 0).MergeRight();
        Assert.False(table.Uniform);
    }

    [Fact]
    public void ATableCanBeSplitBeforeARow()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 2);
        table.Cell(2, 0).Range.Text = "last";

        TableWrapper second = table.Split(table.Rows[2]);

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(1, second.Rows.Count);
        Assert.Equal("last", second.Cell(0, 0).Range.Text.TrimEnd('\a', '\r'));
        Assert.IsType<Paragraph>(table.Table.NextSibling());
        Assert.Same(second.Table, table.Table.NextSibling().NextSibling());
        Assert.Throws<ArgumentException>(() => table.Split(table.Rows[0]));
        AssertValid();
    }

    [Fact]
    public void TableShadingIsAPatternWithTwoColours()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 1);
        Assert.Equal("auto", table.Shading.BackgroundPatternColor);
        Assert.Equal(ShadingTexture.None, table.Shading.Texture);

        table.Shading.BackgroundPatternColor = "FFFF00";
        table.Shading.ForegroundPatternColor = "0000FF";
        table.Shading.Texture = ShadingTexture.Percent25;

        Assert.Equal("FFFF00", table.Shading.BackgroundPatternColor);
        Assert.Equal("0000FF", table.Shading.ForegroundPatternColor);
        Assert.Equal(ShadingTexture.Percent25, table.Shading.Texture);
        AssertValid();

        table.Shading.Clear();
        Assert.Equal("auto", table.Shading.BackgroundPatternColor);
        Assert.Throws<ArgumentException>(() => table.Shading.BackgroundPatternColor = "yellow");
    }

    #endregion

    #region Rows

    [Fact]
    public void RowsHaveFirstLastAndSettersThatApplyToAllOfThem()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 2);

        Assert.Equal(table.Rows[0], table.Rows.First);
        Assert.Equal(table.Rows[2], table.Rows.Last);
        Assert.Equal(false, table.Rows.HeadingFormat);

        table.Rows.Height = 20;
        table.Rows.HeightRule = RowHeightRule.Exactly;
        table.Rows.HeadingFormat = true;
        table.Rows.AllowBreakAcrossPages = false;

        Assert.Equal(20, table.Rows.Height);
        Assert.Equal(RowHeightRule.Exactly, table.Rows.HeightRule);
        Assert.Equal(true, table.Rows.HeadingFormat);
        Assert.Equal(false, table.Rows.AllowBreakAcrossPages);
        Assert.Equal(20, table.Rows[1].Height);

        table.Rows[1].Height = 30;
        Assert.Null(table.Rows.Height);
        AssertValid();
    }

    [Fact]
    public void ARowHasAnAlignmentAndHeightWithARule()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);

        table.Rows[0].Alignment = RowAlignment.Center;
        table.Rows[0].SetHeight(18, RowHeightRule.AtLeast);

        Assert.Equal(RowAlignment.Center, table.Rows[0].Alignment);
        Assert.Equal(RowAlignment.Left, table.Rows[1].Alignment);
        Assert.Equal(18, table.Rows[0].Height);
        Assert.Equal(RowHeightRule.AtLeast, table.Rows[0].HeightRule);
        table.Rows[0].Alignment = RowAlignment.Left;
        Assert.Null(table.Rows[0].TableRow.TableRowProperties.GetFirstChild<TableJustification>());
        AssertValid();
    }

    [Fact]
    public void RowAndRowsShadingShadeTheirCells()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);

        table.Rows[0].Shading.BackgroundPatternColor = "CCCCCC";
        Assert.Equal("CCCCCC", table.Cell(0, 1).Shading.BackgroundPatternColor);
        Assert.Equal("auto", table.Cell(1, 0).Shading.BackgroundPatternColor);
        Assert.Null(table.Rows.Shading.BackgroundPatternColor);

        table.Rows.Shading.BackgroundPatternColor = "EEEEEE";
        Assert.Equal("EEEEEE", table.Rows.Shading.BackgroundPatternColor);
        AssertValid();
    }

    #endregion

    #region Cells

    [Fact]
    public void ACellKnowsItsRowAndColumnAndNestedTables()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 3);
        TableCellWrapper cell = table.Cell(1, 2);

        Assert.Equal(table.Rows[1], cell.Row);
        Assert.Equal(table.Columns[2], cell.Column);
        Assert.Equal(0, cell.Tables.Count);
    }

    [Fact]
    public void ACellSharesTheHeightOfItsRow()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);

        table.Cell(1, 1).Height = 24;
        table.Cell(1, 1).HeightRule = RowHeightRule.Exactly;

        Assert.Equal(24, table.Rows[1].Height);
        Assert.Equal(RowHeightRule.Exactly, table.Rows[1].HeightRule);
        Assert.Equal(24, table.Cell(1, 0).Height);
    }

    [Fact]
    public void WordWrapAndFitTextAreSwitches()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 1);
        TableCellWrapper cell = table.Cell(0, 0);
        Assert.True(cell.WordWrap);
        Assert.False(cell.FitText);

        cell.WordWrap = false;
        cell.FitText = true;
        Assert.False(cell.WordWrap);
        Assert.True(cell.FitText);
        AssertValid();

        cell.WordWrap = true;
        cell.FitText = false;
        Assert.Null(cell.TableCell.TableCellProperties.NoWrap);
        Assert.Null(cell.TableCell.TableCellProperties.TableCellFitText);
    }

    [Fact]
    public void TheCellPreferredWidthHasAType()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 2);
        TableCellWrapper cell = table.Cell(0, 0);
        Assert.Equal(PreferredWidthType.Points, cell.PreferredWidthType);

        cell.PreferredWidth = 100;
        Assert.Equal(100, cell.Width);

        cell.PreferredWidthType = PreferredWidthType.Percent;
        cell.PreferredWidth = 40;
        Assert.Equal(PreferredWidthType.Percent, cell.PreferredWidthType);
        Assert.Equal(40, cell.PreferredWidth);
        Assert.Equal("2000", cell.TableCell.TableCellProperties.TableCellWidth.Width.Value);
        AssertValid();
    }

    [Fact]
    public void ADeletedCellShiftsTheRestOfTheRowLeftAndLeavesItShort()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 3);
        table.Cell(0, 1).Range.Text = "middle";
        table.Cell(0, 2).Range.Text = "end";

        table.Cell(0, 1).Delete(ShiftCells.Left);

        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal("end", table.Cell(0, 1).Range.Text.TrimEnd('\a', '\r'));
        Assert.Equal(1, table.Rows[0].TableRow.TableRowProperties.GetFirstChild<GridAfter>().Val.Value);
        Assert.False(table.Uniform);
        AssertValid();
    }

    [Fact]
    public void DeletingACellCanDeleteTheRowOrColumn()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 3);

        table.Cell(0, 0).Delete(ShiftCells.EntireRow);
        Assert.Equal(2, table.Rows.Count);
        table.Cell(0, 0).Delete(ShiftCells.EntireColumn);
        Assert.Equal(2, table.Columns.Count);
        Assert.Throws<NotSupportedException>(() => table.Cell(0, 0).Delete(ShiftCells.Up));
        AssertValid();
    }

    #endregion

    #region Columns

    [Fact]
    public void AColumnHasItsCellsAndPosition()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 2);
        TableColumnWrapper column = table.Columns[1];

        Assert.Equal(3, column.Cells.Count);
        Assert.Equal(table.Cell(2, 1), column.Cells[2]);
        Assert.True(table.Columns[0].IsFirst);
        Assert.False(column.IsFirst);
        Assert.True(column.IsLast);
        Assert.Equal(table.Columns[0], table.Columns.First);
        Assert.Equal(column, table.Columns.Last);
        Assert.Throws<ArgumentOutOfRangeException>(() => column.Cells[3]);
    }

    [Fact]
    public void AColumnShadesAndSizesItsCells()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 2);

        table.Columns[1].Shading.BackgroundPatternColor = "DDEEFF";
        Assert.Equal("DDEEFF", table.Cell(2, 1).Shading.BackgroundPatternColor);
        Assert.Equal("auto", table.Cell(2, 0).Shading.BackgroundPatternColor);

        table.Columns[0].PreferredWidthType = PreferredWidthType.Percent;
        table.Columns[0].PreferredWidth = 30;
        Assert.Equal(30, table.Cell(1, 0).PreferredWidth);
        Assert.Equal(PreferredWidthType.Percent, table.Columns[0].PreferredWidthType);
        AssertValid();
    }

    [Fact]
    public void AColumnIsAddedBeforeAnotherAndWidensCellsThatSpanIt()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 3);
        table.Cell(0, 0).MergeRight();
        table.Cell(1, 1).Range.Text = "kept";

        TableColumnWrapper added = table.Columns.Add(table.Columns[1]);

        Assert.Equal(4, table.Columns.Count);
        Assert.Equal(table.Columns[1], added);
        Assert.Equal(2, table.Rows[0].Cells.Count);
        Assert.Equal(3, TableCellWrapper.Span(table.Cell(0, 0).TableCell));
        Assert.Equal(4, table.Rows[1].Cells.Count);
        Assert.Equal("", table.Cell(1, 1).Range.Text.TrimEnd('\a', '\r'));
        Assert.Equal("kept", table.Cell(1, 2).Range.Text.TrimEnd('\a', '\r'));
        Assert.Throws<ArgumentException>(() => table.Columns.Add(new TableColumnWrapper(new GridColumn())));
        AssertValid();
    }

    [Fact]
    public void ColumnWidthsCanBeDistributedEvenly()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 3);
        table.Columns[0].Width = 50;
        table.Columns[1].Width = 100;
        table.Columns[2].Width = 150;

        table.Columns.DistributeWidth();

        Assert.All(table.Columns, c => Assert.Equal(100, c.Width));
        AssertValid();
    }

    #endregion

    #region Borders

    [Fact]
    public void ColumnBordersAreOnItsCells()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 3, 2);
        table.Borders.Clear();

        table.Columns[0].Borders[BorderPosition.Top].LineStyle = BorderLineStyle.DoubleLine;
        table.Columns[0].Borders[BorderPosition.Horizontal].LineStyle = BorderLineStyle.SingleLine;

        Assert.Equal(BorderLineStyle.DoubleLine, table.Cell(0, 0).Borders[BorderPosition.Top].LineStyle);
        Assert.Equal(BorderLineStyle.SingleLine, table.Cell(1, 0).Borders[BorderPosition.Top].LineStyle);
        Assert.Equal(BorderLineStyle.SingleLine, table.Cell(0, 0).Borders[BorderPosition.Bottom].LineStyle);
        Assert.Throws<NotSupportedException>(() => table.Columns[0].Borders[BorderPosition.Vertical]);
        table.Columns[0].Borders.Clear();
        Assert.Equal(BorderLineStyle.None, table.Cell(1, 0).Borders[BorderPosition.Top].LineStyle);
        AssertValid();
    }

    [Fact]
    public void TheOutsideAndInsideBordersCanBeReadAndSetAsGroups()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);
        BordersWrapper borders = table.Borders;
        Assert.Equal(6, borders.Count);
        Assert.Equal(8, table.Cell(0, 0).Borders.Count);
        Assert.True(borders.Enable);
        Assert.Equal(BorderLineStyle.SingleLine, borders.OutsideLineStyle);
        Assert.Equal(0.5, borders.InsideLineWidth);

        borders.OutsideLineStyle = BorderLineStyle.DoubleLine;
        borders.OutsideLineWidth = 2.25;
        borders.OutsideColor = "FF0000";
        borders.InsideColor = "00FF00";

        Assert.Equal(BorderLineStyle.DoubleLine, borders[BorderPosition.Left].LineStyle);
        Assert.Equal(2.25, borders.OutsideLineWidth);
        Assert.Equal("FF0000", borders.OutsideColor);
        Assert.Equal("00FF00", borders[BorderPosition.Horizontal].Color);
        Assert.Equal(BorderLineStyle.SingleLine, borders.InsideLineStyle);
        AssertValid();

        borders[BorderPosition.Top].LineWidth = 4;
        Assert.Equal(0, borders.OutsideLineWidth);
        Assert.Throws<ArgumentException>(() => borders.OutsideLineStyle = BorderLineStyle.Other);
    }

    [Fact]
    public void EnableSwitchesTheBordersOnAndOff()
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 2, 2);

        table.Borders.Enable = false;
        Assert.False(table.Borders.Enable);
        Assert.False(table.Borders[BorderPosition.Top].Visible);

        table.Borders.Enable = true;
        Assert.True(table.Borders.Enable);
        Assert.Equal(BorderLineStyle.SingleLine, table.Borders[BorderPosition.Vertical].LineStyle);

        table.Borders[BorderPosition.Top].Visible = false;
        Assert.Equal(BorderLineStyle.None, table.Borders[BorderPosition.Top].LineStyle);
        table.Borders[BorderPosition.Top].Visible = true;
        Assert.True(table.Borders[BorderPosition.Top].Visible);
        AssertValid();
    }

    [Theory]
    [InlineData(BorderLineStyle.DashSmallGap)]
    [InlineData(BorderLineStyle.DashLargeGap)]
    [InlineData(BorderLineStyle.DashDot)]
    [InlineData(BorderLineStyle.DashDotDot)]
    [InlineData(BorderLineStyle.Triple)]
    [InlineData(BorderLineStyle.ThinThickSmallGap)]
    [InlineData(BorderLineStyle.ThickThinMediumGap)]
    [InlineData(BorderLineStyle.ThinThickThinLargeGap)]
    [InlineData(BorderLineStyle.SingleWavy)]
    [InlineData(BorderLineStyle.DoubleWavy)]
    [InlineData(BorderLineStyle.DashDotStroked)]
    [InlineData(BorderLineStyle.Emboss3D)]
    [InlineData(BorderLineStyle.Engrave3D)]
    [InlineData(BorderLineStyle.Outset)]
    [InlineData(BorderLineStyle.Inset)]
    public void EveryLineStyleIsWrittenAndReadBack(BorderLineStyle style)
    {
        TableWrapper table = _document.Tables.Add(_document.Range(0, 0), 1, 1);

        table.Borders[BorderPosition.Top].LineStyle = style;

        Assert.Equal(style, table.Borders[BorderPosition.Top].LineStyle);
        AssertValid();
    }

    [Fact]
    public void TheLineStyleValuesAreThoseOfWord()
    {
        Assert.Equal(1, (int)BorderLineStyle.SingleLine);
        Assert.Equal(3, (int)BorderLineStyle.DashSmallGap);
        Assert.Equal(5, (int)BorderLineStyle.DashDot);
        Assert.Equal(7, (int)BorderLineStyle.DoubleLine);
        Assert.Equal(18, (int)BorderLineStyle.SingleWavy);
        Assert.Equal(23, (int)BorderLineStyle.Outset);
        Assert.Equal(24, (int)BorderLineStyle.Inset);
    }

    #endregion
}
