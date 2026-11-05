using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// <c>Style.Table</c> and its conditional formats. The markup is what Word 16 wrote (the table-style-* scenarios of MarkupConformanceTest).
/// </summary>
public class TableStyleTest : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
    private readonly WordprocessingDocumentWrapper _document;

    public TableStyleTest()
    {
        _document = WordprocessingDocumentWrapper.Create(_path, true, false);
    }

    public void Dispose()
    {
        _document.Dispose();
        File.Delete(_path);
        GC.SuppressFinalize(this);
    }

    private TableStyleWrapper NewTableStyle()
    {
        return _document.Styles.Add("Grid", StyleType.Table).Table;
    }

    [Fact]
    public void OnlyATableStyleHasTableProperties()
    {
        Assert.Null(_document.Styles.Add("Para", StyleType.Paragraph).Table);
        Assert.NotNull(NewTableStyle());
    }

    [Fact]
    public void TheMembersReadWhatTheyAreSetTo()
    {
        TableStyleWrapper table = NewTableStyle();
        Assert.Equal(RowAlignment.Left, table.Alignment);
        Assert.True(table.AllowBreakAcrossPage);
        Assert.Equal(1, table.RowStripe);
        Assert.Equal(0, table.Spacing);

        table.Alignment = RowAlignment.Right;
        table.AllowBreakAcrossPage = false;
        table.RowStripe = 2;
        table.ColumnStripe = 3;
        table.Spacing = 1.5;
        table.LeftIndent = 10;
        table.TopPadding = 5;
        table.RightPadding = 8;

        Assert.Equal(RowAlignment.Right, table.Alignment);
        Assert.False(table.AllowBreakAcrossPage);
        Assert.Equal(2, table.RowStripe);
        Assert.Equal(3, table.ColumnStripe);
        Assert.Equal(1.5, table.Spacing);
        Assert.Equal(10, table.LeftIndent);
        Assert.Equal(5, table.TopPadding);
        Assert.Equal(8, table.RightPadding);
        Assert.Equal(0, table.BottomPadding);

        table.Alignment = RowAlignment.Left;
        table.AllowBreakAcrossPage = true;
        table.Spacing = 0;
        Assert.Equal(RowAlignment.Left, table.Alignment);
        Assert.True(table.AllowBreakAcrossPage);
        Assert.Equal(0, table.Spacing);
    }

    [Fact]
    public void BordersAndShadingAreRead()
    {
        TableStyleWrapper table = NewTableStyle();

        table.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;
        table.Borders[BorderPosition.Horizontal].LineStyle = BorderLineStyle.DoubleLine;
        table.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);

        Assert.Equal(BorderLineStyle.SingleLine, table.Borders[BorderPosition.Top].LineStyle);
        Assert.Equal(BorderLineStyle.DoubleLine, table.Borders[BorderPosition.Horizontal].LineStyle);
        Assert.Equal(ColorHelper.RGB(255, 0, 0), (int)table.Shading.BackgroundPatternColor.Value.Value);
        Assert.Throws<NotSupportedException>(() => table.Borders[BorderPosition.DiagonalDown]);
    }

    [Fact]
    public void AConditionalFormatIsMadeWhenSomethingIsSetOnIt()
    {
        TableStyleWrapper table = NewTableStyle();
        ConditionalStyleWrapper first = table.Condition(ConditionCode.FirstRow);

        Assert.False(first.Font.Bold);
        Assert.Empty(_document.Styles["Grid"].Style.Elements<TableStyleProperties>());

        first.Font.Bold = true;
        first.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        Assert.True(table.Condition(ConditionCode.FirstRow).Font.Bold);
        Assert.Equal(ParagraphAlignment.Center, table.Condition(ConditionCode.FirstRow).ParagraphFormat.Alignment);
        Assert.False(table.Condition(ConditionCode.LastRow).Font.Bold);
        Assert.Single(_document.Styles["Grid"].Style.Elements<TableStyleProperties>());

        first.Font.Bold = false;
        first.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        Assert.Empty(_document.Styles["Grid"].Style.Elements<TableStyleProperties>());
    }

    [Fact]
    public void ConditionalFormatsAreKeptInTheOrderWordWritesThem()
    {
        TableStyleWrapper table = NewTableStyle();

        table.Condition(ConditionCode.SouthWestCell).Font.Bold = true;
        table.Condition(ConditionCode.FirstRow).Font.Bold = true;
        table.Condition(ConditionCode.OddRowBanding).Font.Bold = true;

        Assert.Equal(
            new[] { TableStyleOverrideValues.FirstRow, TableStyleOverrideValues.Band1Horizontal, TableStyleOverrideValues.SouthWestCell },
            _document.Styles["Grid"].Style.Elements<TableStyleProperties>().Select(c => c.Type.Value).ToArray());
    }

    [Fact]
    public void ShadingABorderOrAPaddingOfAConditionalFormatWritesWhatWordWrites()
    {
        TableStyleWrapper table = NewTableStyle();

        table.Condition(ConditionCode.FirstRow).Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);
        table.Condition(ConditionCode.LastRow).TopPadding = 3;
        table.Condition(ConditionCode.FirstColumn).Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.SingleLine;

        TableStyleProperties first = TableOf(ConditionCode.FirstRow);
        Assert.NotNull(first.TableStyleConditionalFormattingTableProperties);
        Assert.Equal(6, first.TableStyleConditionalFormattingTableCellProperties.TableCellBorders.ChildElements.Count);
        Assert.Equal(BorderLineStyle.None, table.Condition(ConditionCode.FirstRow).Borders[BorderPosition.Top].LineStyle);

        Assert.Equal(3, table.Condition(ConditionCode.LastRow).TopPadding);
        Assert.Equal(0, table.Condition(ConditionCode.LastRow).LeftPadding);
        Assert.Null(TableOf(ConditionCode.LastRow).TableStyleConditionalFormattingTableCellProperties.TableCellBorders);

        Assert.Equal(BorderLineStyle.SingleLine, table.Condition(ConditionCode.FirstColumn).Borders[BorderPosition.Bottom].LineStyle);
        Assert.Throws<NotSupportedException>(() => table.Condition(ConditionCode.FirstColumn).Borders[BorderPosition.DiagonalUp]);
    }

    [Fact]
    public void ACodeThatIsNotOneIsRejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => NewTableStyle().Condition((ConditionCode)99));
    }

    private TableStyleProperties TableOf(ConditionCode code)
    {
        return _document.Styles["Grid"].Style.Elements<TableStyleProperties>().First(c => c.Type.Value == Map(code));
    }

    private static TableStyleOverrideValues Map(ConditionCode code)
    {
        return code == ConditionCode.FirstRow ? TableStyleOverrideValues.FirstRow : code == ConditionCode.LastRow ? TableStyleOverrideValues.LastRow : TableStyleOverrideValues.FirstColumn;
    }
}
