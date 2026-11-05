using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXml.Wrappers.Helpers;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// How the other cells or columns change when one is given a width, as <c>WdRulerStyle</c> in Word.
/// </summary>
public enum RulerStyle
{
    /// <summary>Nothing else changes, so the row or the table gets wider or narrower.</summary>
    None = 0,

    /// <summary>The cells or columns after it take up the difference in proportion to their widths.</summary>
    Proportional = 1,

    /// <summary>The next one takes up the whole difference.</summary>
    FirstColumn = 2,

    /// <summary>The ones after it share the difference so that they are the same width.</summary>
    SameWidth = 3
}

//Where cells of different rows end at different places the table has more grid columns than any row has cells, and a cell spans several.
internal static class TableWidths
{
    //The width a cell takes in twips: what it asks for when that is in twentieths of a point, otherwise the grid columns it covers.
    internal static int EffectiveWidth(TableCell cell, IList<int> grid)
    {
        TableCellWidth width = cell.TableCellProperties?.TableCellWidth;
        bool isDxa = width != null && (width.Type == null || !width.Type.HasValue || width.Type.Value == TableWidthUnitValues.Dxa);
        if (isDxa && int.TryParse(width.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) && twips > 0)
        {
            return twips;
        }
        return grid.Skip(TableCellWrapper.GetGridStart(cell)).Take(TableCellWrapper.Span(cell)).Sum();
    }

    internal static List<int> ReadGrid(Table table)
    {
        return table.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(c => int.TryParse(c.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int w) ? w : 0).ToList();
    }

    /// <summary>
    /// Makes the grid fit the cells, as Word does when a cell is given a width of its own: every place where a cell of any row ends is the end of
    /// a grid column, and each cell spans the grid columns between its own ends.
    /// </summary>
    /// <summary>
    /// Turns the older horizontal merges (<c>hMerge</c>, a cell that starts a merge and the cells after it that continue it) into cells that
    /// span columns (<c>gridSpan</c>), which is what Word writes now. The text of the continuing cells goes to the end of the first.
    /// </summary>
    internal static void NormaliseHorizontalMerges(Table table)
    {
        if (!table.Descendants<HorizontalMerge>().Any())
        {
            return;
        }
        List<int> grid = ReadGrid(table);
        foreach (TableRow row in table.Elements<TableRow>())
        {
            List<TableCell> cells = row.Elements<TableCell>().ToList();
            for (int index = 0; index < cells.Count; index++)
            {
                TableCell cell = cells[index];
                HorizontalMerge merge = cell.TableCellProperties?.HorizontalMerge;
                if (merge == null)
                {
                    continue;
                }
                merge.Remove();
                if (merge.Val?.Value != MergedCellValues.Restart)
                {
                    //A continuation with nothing to continue is an ordinary cell.
                    continue;
                }

                int span = TableCellWrapper.Span(cell);
                int width = EffectiveWidth(cell, grid);
                while (index + 1 < cells.Count && cells[index + 1].TableCellProperties?.HorizontalMerge is HorizontalMerge next && next.Val?.Value != MergedCellValues.Restart)
                {
                    TableCell continued = cells[++index];
                    span += TableCellWrapper.Span(continued);
                    width += EffectiveWidth(continued, grid);
                    foreach (Paragraph paragraph in continued.Elements<Paragraph>().Where(p => p.Descendants<Run>().Any()).ToList())
                    {
                        paragraph.Remove();
                        cell.AppendChild(paragraph);
                    }
                    RangeWrapper.DeleteElements(new OpenXmlElement[] { continued });
                }
                if (span > 1)
                {
                    cell.TableCellProperties.GridSpan = new GridSpan() { Val = span };
                }
                new TableCellWrapper(cell).WidthTwips = width;
            }
        }
    }

    private const int SliverTwips = 2;

    internal static void FitGridToCells(Table table)
    {
        TableGrid tableGrid = table?.GetFirstChild<TableGrid>();
        if (tableGrid == null)
        {
            return;
        }
        List<TableRow> rows = table.Elements<TableRow>().ToList();

        List<int> grid = ReadGrid(table);
        SortedSet<int> ends = new SortedSet<int>();
        Dictionary<TableCell, int[]> extents = new Dictionary<TableCell, int[]>();
        Dictionary<TableRow, int> starts = new Dictionary<TableRow, int>();
        foreach (TableRow row in rows)
        {
            //A row may start after some grid columns that it leaves unused.
            int position = int.TryParse(row.TableRowProperties?.GetFirstChild<WidthBeforeTableRow>()?.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int before) && row.TableRowProperties.GetFirstChild<GridBefore>() != null ? before : 0;
            starts[row] = position;
            if (position > 0)
            {
                ends.Add(position);
            }
            foreach (TableCell cell in row.Elements<TableCell>())
            {
                int start = position;
                position += EffectiveWidth(cell, grid);
                extents[cell] = new[] { start, position };
                ends.Add(position);
            }
        }
        if (ends.Count == 0 || ends.Min <= 0)
        {
            return;
        }

        //Word does not make a grid column of a sliver: ends that are within a couple of twips of each other are one end, the furthest.
        List<int> boundaries = new List<int>();
        foreach (int end in ends)
        {
            if (boundaries.Count > 0 && end - boundaries[boundaries.Count - 1] <= SliverTwips)
            {
                boundaries[boundaries.Count - 1] = end;
            }
            else
            {
                boundaries.Add(end);
            }
        }
        Func<int, int> snap = position => position <= 0 ? 0 : boundaries.FirstOrDefault(b => b >= position) is int found && found > 0 ? found : position;
        foreach (KeyValuePair<TableCell, int[]> pair in extents)
        {
            int span = boundaries.Count(b => b > snap(pair.Value[0]) && b <= snap(pair.Value[1]));
            if (span > 1)
            {
                TableCellProperties properties = pair.Key.TableCellProperties;
                if (properties == null)
                {
                    properties = new TableCellProperties();
                    pair.Key.PrependChild(properties);
                }
                properties.GridSpan = new GridSpan() { Val = span };
            }
            else
            {
                pair.Key.TableCellProperties?.GridSpan?.Remove();
            }
        }

        tableGrid.RemoveAllChildren<GridColumn>();
        int previous = 0;
        foreach (int end in boundaries)
        {
            tableGrid.AppendChild(new GridColumn() { Width = (end - previous).ToString(CultureInfo.InvariantCulture) });
            previous = end;
        }

        //A row that starts after the start of the table, or ends before the end of it, leaves the grid columns around it unused.
        foreach (TableRow row in rows)
        {
            int rowStart = snap(starts[row]);
            int rowEnd = snap(row.Elements<TableCell>().Select(c => extents[c][1]).DefaultIfEmpty(0).Max());
            int unusedBefore = boundaries.Count(b => b > 0 && b <= rowStart);
            int unusedAfter = boundaries.Count(b => b > rowEnd);
            TableRowProperties properties = row.TableRowProperties;
            properties?.GetFirstChild<GridBefore>()?.Remove();
            properties?.GetFirstChild<WidthBeforeTableRow>()?.Remove();
            properties?.GetFirstChild<GridAfter>()?.Remove();
            properties?.GetFirstChild<WidthAfterTableRow>()?.Remove();
            bool after = unusedAfter > 0 && rowEnd > 0;
            if (unusedBefore > 0 || after)
            {
                if (properties == null)
                {
                    properties = new TableRowProperties();
                    row.InsertAfter(properties, row.GetFirstChild<TablePropertyExceptions>());
                }

                //In the order Word writes them: gridBefore, wBefore, gridAfter, wAfter.
                if (after)
                {
                    properties.PrependChild(new WidthAfterTableRow() { Width = (boundaries[boundaries.Count - 1] - rowEnd).ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
                    properties.PrependChild(new GridAfter() { Val = unusedAfter });
                }
                if (unusedBefore > 0)
                {
                    properties.PrependChild(new WidthBeforeTableRow() { Width = rowStart.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
                    properties.PrependChild(new GridBefore() { Val = unusedBefore });
                }
            }
        }
    }

    /// <summary>
    /// Shares <paramref name="total"/> among a number of places in proportion to <paramref name="weights"/>, as whole twips, with the last one
    /// taking what is left over, as Word does.
    /// </summary>
    internal static int[] Share(int total, IList<int> weights)
    {
        long weightTotal = weights.Sum(w => (long)w);
        int[] shares = new int[weights.Count];
        long cumulative = 0;
        int given = 0;
        for (int index = 0; index < weights.Count; index++)
        {
            cumulative += weightTotal == 0 ? 1 : weights[index];
            long denominator = weightTotal == 0 ? weights.Count : weightTotal;
            int upTo = (int)(total * cumulative / denominator);
            shares[index] = upTo - given;
            given = upTo;
        }
        return shares;
    }

    /// <summary>
    /// The widths the <paramref name="following"/> places take when one before them changes by <paramref name="delta"/> twips, so that the
    /// whole stays as wide as it was.
    /// </summary>
    internal static int[] Adjust(IList<int> following, int delta, RulerStyle style)
    {
        int[] widths = following.ToArray();
        if (widths.Length == 0)
        {
            return widths;
        }
        int total = Math.Max(0, widths.Sum() - delta);
        if (style == RulerStyle.FirstColumn)
        {
            widths[0] = Math.Max(0, widths[0] - delta);
        }
        else if (style == RulerStyle.SameWidth)
        {
            widths = Share(total, widths.Select(_ => 1).ToList());
        }
        else if (style == RulerStyle.Proportional)
        {
            widths = Share(total, widths);
        }
        return widths;
    }
}

public partial class TableCellWrapper
{
    /// <summary>
    /// Gives the cell a width, and the cells after it in the row the difference according to <paramref name="rulerStyle"/>, as
    /// <c>Cell.SetWidth</c> in Word. With <see cref="RulerStyle.None"/> it is the same as setting <see cref="Width"/>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public void SetWidth(double columnWidth, RulerStyle rulerStyle)
    {
        int twips = ConversionHelper.PointsToTwips(columnWidth, nameof(columnWidth));
        Table table = _tableCell.Ancestors<Table>().FirstOrDefault();
        if (rulerStyle != RulerStyle.None && table != null)
        {
            List<int> grid = TableWidths.ReadGrid(table);
            List<TableCell> following = _tableCell.ElementsAfter().OfType<TableCell>().ToList();
            int delta = twips - TableWidths.EffectiveWidth(_tableCell, grid);
            int[] adjusted = TableWidths.Adjust(following.Select(c => TableWidths.EffectiveWidth(c, grid)).ToList(), delta, rulerStyle);
            for (int index = 0; index < following.Count; index++)
            {
                new TableCellWrapper(following[index]).WidthTwips = adjusted[index];
            }
        }
        WidthTwips = twips;
        TableWidths.FitGridToCells(table);
    }
}

public partial class TableCellWrapperCollection
{
    /// <summary>
    /// The width in points of the cells; null when they differ. Setting it gives every cell that width and brings the grid of the table into
    /// line with the cells, as Word does.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public double? Width
    {
        get { return ReadAll(c => c.WidthTwips) is int twips ? ConversionHelper.TwipsToPoints(twips) : (double?)null; }
        set
        {
            if (value.HasValue)
            {
                SetWidth(value.Value, RulerStyle.None);
            }
        }
    }

    /// <summary>
    /// Gives every cell the width, and, in each row, the cells after the last of them that are not in the collection the difference
    /// according to <paramref name="rulerStyle"/>, as <c>Cells.SetWidth</c> in Word.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public void SetWidth(double columnWidth, RulerStyle rulerStyle)
    {
        int twips = ConversionHelper.PointsToTwips(columnWidth, nameof(columnWidth));
        IList<TableCellWrapper> cells = Items;
        foreach (IGrouping<TableRow, TableCellWrapper> row in cells.GroupBy(c => c.Parent.TableRow))
        {
            Table table = row.Key.Ancestors<Table>().FirstOrDefault();
            List<int> grid = table == null ? new List<int>() : TableWidths.ReadGrid(table);
            TableCell last = row.Last().TableCell;
            List<TableCell> following = last.ElementsAfter().OfType<TableCell>().Where(c => !row.Any(r => ReferenceEquals(r.TableCell, c))).ToList();
            int delta = row.Sum(c => twips - TableWidths.EffectiveWidth(c.TableCell, grid));
            int[] adjusted = rulerStyle == RulerStyle.None
                ? new int[following.Count]
                : TableWidths.Adjust(following.Select(c => TableWidths.EffectiveWidth(c, grid)).ToList(), delta, rulerStyle);
            for (int index = 0; index < following.Count && rulerStyle != RulerStyle.None; index++)
            {
                new TableCellWrapper(following[index]).WidthTwips = adjusted[index];
            }
        }
        foreach (TableCellWrapper cell in cells)
        {
            cell.WidthTwips = twips;
        }
        foreach (Table table in cells.Select(c => c.TableCell.Ancestors<Table>().FirstOrDefault()).Where(t => t != null).Distinct())
        {
            TableWidths.FitGridToCells(table);
        }
    }

    /// <summary>
    /// Makes the cells of each row the same width, sharing out between them the width they have together, as
    /// <c>Cells.DistributeWidth</c> in Word.
    /// </summary>
    public void DistributeWidth()
    {
        IList<TableCellWrapper> cells = Items;
        foreach (IGrouping<TableRow, TableCellWrapper> row in cells.GroupBy(c => c.Parent.TableRow))
        {
            Table table = row.Key.Ancestors<Table>().FirstOrDefault();
            List<int> grid = table == null ? new List<int>() : TableWidths.ReadGrid(table);
            List<TableCellWrapper> members = row.ToList();
            int[] widths = TableWidths.Share(members.Sum(c => TableWidths.EffectiveWidth(c.TableCell, grid)), members.Select(_ => 1).ToList());
            for (int index = 0; index < members.Count; index++)
            {
                members[index].WidthTwips = widths[index];
            }
        }
        foreach (Table table in cells.Select(c => c.TableCell.Ancestors<Table>().FirstOrDefault()).Where(t => t != null).Distinct())
        {
            TableWidths.FitGridToCells(table);
        }
    }
}

public partial class TableColumnWrapper
{
    /// <summary>
    /// Gives the column a width, and the columns after it the difference according to <paramref name="rulerStyle"/>, as
    /// <c>Column.SetWidth</c> in Word. With <see cref="RulerStyle.None"/> it is the same as setting <see cref="Width"/>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public void SetWidth(double columnWidth, RulerStyle rulerStyle)
    {
        int twips = ConversionHelper.PointsToTwips(columnWidth, nameof(columnWidth));
        Table table = _gridColumn.Ancestors<Table>().First();
        if (rulerStyle != RulerStyle.None)
        {
            List<GridColumn> following = _gridColumn.ElementsAfter().OfType<GridColumn>().ToList();
            int[] adjusted = TableWidths.Adjust(following.Select(c => new TableColumnWrapper(c).WidthTwips).ToList(), twips - WidthOrZero, rulerStyle);
            for (int index = 0; index < following.Count; index++)
            {
                following[index].Width = adjusted[index].ToString(CultureInfo.InvariantCulture);
            }
        }
        _gridColumn.Width = twips.ToString(CultureInfo.InvariantCulture);
        TableWrapper.ColumnsChanged(table);
    }
}

public partial class TableColumnWrapperCollection
{
    /// <summary>
    /// Gives every column the width, as <c>Columns.SetWidth</c> in Word. The columns after the last of them take up the difference
    /// according to <paramref name="rulerStyle"/>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The width is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public void SetWidth(double columnWidth, RulerStyle rulerStyle)
    {
        IList<TableColumnWrapper> columns = Items;
        foreach (TableColumnWrapper column in columns)
        {
            column.SetWidth(columnWidth, rulerStyle);
        }
    }
}

public partial class TableRowWrapper
{
    /// <summary>
    /// The distance in points between the left margin and the left edge of the row: the space the row leaves unused before its first cell,
    /// 0 when it has none (<c>Row.LeftIndent</c> in Word). Setting it moves the row without changing its cells, as
    /// <c>SetLeftIndent(value, RulerStyle.None)</c>.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public double LeftIndent
    {
        get
        {
            TableRowProperties properties = _tableRow.TableRowProperties;
            return int.TryParse(properties?.GetFirstChild<WidthBeforeTableRow>()?.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) && properties.GetFirstChild<GridBefore>() != null
                ? ConversionHelper.TwipsToPoints(twips)
                : 0;
        }
        set { SetLeftIndent(value, RulerStyle.None); }
    }

    /// <summary>
    /// Moves the row to the right, as <c>Row.SetLeftIndent</c> in Word, by leaving space before its first cell. Like the rest of the table it
    /// is made ragged: the grid gets a column where the row starts. With <see cref="RulerStyle.None"/> the cells keep their widths and the row
    /// is pushed out; otherwise the cells give up the width of the indent (<see cref="RulerStyle.FirstColumn"/>: the first cell,
    /// <see cref="RulerStyle.Proportional"/>: in proportion, <see cref="RulerStyle.SameWidth"/>: all the same) so that the row ends where it did.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is negative, not a number or more than 1584 points, the most Word allows.</exception>
    public void SetLeftIndent(double leftIndent, RulerStyle rulerStyle)
    {
        int twips = ConversionHelper.PointsToTwips(leftIndent, nameof(leftIndent));
        Table table = _tableRow.Ancestors<Table>().FirstOrDefault();
        if (table == null)
        {
            return;
        }

        List<int> grid = TableWidths.ReadGrid(table);
        List<TableCell> cells = _tableRow.Elements<TableCell>().ToList();
        int current = (int)Math.Round(LeftIndent * 20, MidpointRounding.AwayFromZero);
        if (rulerStyle != RulerStyle.None && cells.Count > 0)
        {
            List<int> widths = cells.Select(c => TableWidths.EffectiveWidth(c, grid)).ToList();
            int delta = twips - current;
            int[] adjusted;
            if (rulerStyle == RulerStyle.SameWidth)
            {
                //Word gives each cell a whole number of twips and does not share out what is left.
                adjusted = widths.Select(_ => Math.Max(0, (widths.Sum() - delta) / widths.Count)).ToArray();
            }
            else
            {
                adjusted = TableWidths.Adjust(widths, delta, rulerStyle);
                if (rulerStyle == RulerStyle.Proportional)
                {
                    adjusted = TableWidths.Share(Math.Max(0, widths.Sum() - delta), widths);
                }
            }
            for (int index = 0; index < cells.Count; index++)
            {
                new TableCellWrapper(cells[index]).WidthTwips = adjusted[index];
            }
        }

        TableRowProperties properties = GetOrAddProperties();
        properties.GetFirstChild<GridBefore>()?.Remove();
        properties.GetFirstChild<WidthBeforeTableRow>()?.Remove();
        if (twips > 0)
        {
            properties.PrependChild(new WidthBeforeTableRow() { Width = twips.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
            properties.PrependChild(new GridBefore() { Val = 1 });
        }
        TableWidths.FitGridToCells(table);

        //The cells are as wide as they were given; only the width of the table follows the grid.
        TableWrapper.ColumnsChanged(table, false);
    }
}
