using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

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
    internal static void FitGridToCells(Table table)
    {
        TableGrid tableGrid = table?.GetFirstChild<TableGrid>();
        if (tableGrid == null)
        {
            return;
        }
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        if (rows.Any(r => r.TableRowProperties?.GetFirstChild<GridBefore>() != null))
        {
            return;
        }

        List<int> grid = ReadGrid(table);
        SortedSet<int> ends = new SortedSet<int>();
        Dictionary<TableCell, int[]> extents = new Dictionary<TableCell, int[]>();
        foreach (TableRow row in rows)
        {
            int position = 0;
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

        List<int> boundaries = ends.ToList();
        foreach (KeyValuePair<TableCell, int[]> pair in extents)
        {
            int span = boundaries.Count(b => b > pair.Value[0] && b <= pair.Value[1]);
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
        int before = 0;
        foreach (int end in boundaries)
        {
            tableGrid.AppendChild(new GridColumn() { Width = (end - before).ToString(CultureInfo.InvariantCulture) });
            before = end;
        }

        //A row that ends before the end of the table leaves the grid columns after it unused.
        foreach (TableRow row in rows)
        {
            int rowEnd = row.Elements<TableCell>().Select(c => extents[c][1]).DefaultIfEmpty(0).Max();
            int unused = boundaries.Count(b => b > rowEnd);
            TableRowProperties properties = row.TableRowProperties;
            properties?.GetFirstChild<GridAfter>()?.Remove();
            properties?.GetFirstChild<WidthAfterTableRow>()?.Remove();
            if (unused > 0 && rowEnd > 0)
            {
                if (properties == null)
                {
                    properties = new TableRowProperties();
                    row.InsertAfter(properties, row.GetFirstChild<TablePropertyExceptions>());
                }
                properties.PrependChild(new WidthAfterTableRow() { Width = (boundaries[boundaries.Count - 1] - rowEnd).ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa });
                properties.PrependChild(new GridAfter() { Val = unused });
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
