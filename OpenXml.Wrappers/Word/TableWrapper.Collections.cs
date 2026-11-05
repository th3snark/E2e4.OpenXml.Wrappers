using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The formatting that Word's <c>Cells</c> applies to all the cells of a range at once. Reading gives null where the cells differ, and
/// setting applies to every one.
/// </summary>
public partial class TableCellWrapperCollection
{
    private IList<TableCellWrapper> Items
    {
        get { return this.ToList(); }
    }

    private T? ReadAll<T>(Func<TableCellWrapper, T> read) where T : struct
    {
        IList<TableCellWrapper> cells = Items;
        if (cells.Count == 0)
        {
            return null;
        }
        T first = read(cells[0]);
        return cells.All(c => EqualityComparer<T>.Default.Equals(read(c), first)) ? first : (T?)null;
    }

    //The rows the cells are in, each once.
    private IList<TableRowWrapper> Rows
    {
        get { return Items.Select(c => c.Parent).Distinct().ToList(); }
    }

    /// <summary>
    /// The shading of all the cells, as <c>Cells.Shading</c> in Word. When the cells are all those of a table the shading of the table is set
    /// too, as Word does.
    /// </summary>
    public ShadingWrapper Shading
    {
        get { return new ShadingWrapper(create => ShadingOf(Items, create)); }
    }

    //The shading of each cell, and that of each table that the cells are all the cells of.
    internal static IList<Shading> ShadingOf(IList<TableCellWrapper> cells, bool create)
    {
        List<Shading> shadings = cells.Select(c => c.GetShading(create)).ToList();
        foreach (IGrouping<Table, TableCellWrapper> group in cells.GroupBy(c => c.TableCell.Ancestors<Table>().First()))
        {
            if (group.Count() == group.Key.Elements<TableRow>().Sum(r => r.Elements<TableCell>().Count()))
            {
                shadings.Add(new TableWrapper(group.Key).GetTableShading(create));
            }
        }
        return shadings;
    }

    /// <summary>The vertical alignment of the text in the cells; null when they differ.</summary>
    public CellVerticalAlignment? VerticalAlignment
    {
        get { return ReadAll(c => c.VerticalAlignment); }
        set
        {
            if (value.HasValue)
            {
                foreach (TableCellWrapper cell in Items)
                {
                    cell.VerticalAlignment = value.Value;
                }
            }
        }
    }

    /// <summary>What <see cref="PreferredWidth"/> means: points, a percentage or automatic; null when the cells differ.</summary>
    public PreferredWidthType? PreferredWidthType
    {
        get { return ReadAll(c => c.PreferredWidthType); }
        set
        {
            if (value.HasValue)
            {
                foreach (TableCellWrapper cell in Items)
                {
                    cell.PreferredWidthType = value.Value;
                }
            }
        }
    }

    /// <summary>
    /// The preferred width of the cells, in points or percent according to <see cref="PreferredWidthType"/>; null when they differ. Setting it
    /// in points also sets the width of every column that the cells are all the cells of, as Word does.
    /// </summary>
    public double? PreferredWidth
    {
        get { return ReadAll(c => c.PreferredWidth); }
        set
        {
            if (!value.HasValue)
            {
                return;
            }
            IList<TableCellWrapper> cells = Items;
            foreach (TableCellWrapper cell in cells)
            {
                cell.PreferredWidth = value.Value;
            }
            if (cells.Count > 0 && cells.All(c => c.PreferredWidthType == Word.PreferredWidthType.Points))
            {
                FitColumnsToCells(cells);
            }
        }
    }

    //A column that all of the cells of it are in takes the width the cells ask for.
    private static void FitColumnsToCells(IList<TableCellWrapper> cells)
    {
        foreach (IGrouping<Table, TableCellWrapper> group in cells.GroupBy(c => c.TableCell.Ancestors<Table>().First()))
        {
            List<GridColumn> grid = group.Key.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
            if (grid == null)
            {
                continue;
            }
            bool changed = false;
            for (int index = 0; index < grid.Count; index++)
            {
                List<TableCell> inColumn = group.Key.Elements<TableRow>().Select(r => TableCellWrapper.CellAtGridColumn(r, index)).Where(c => c != null).ToList();
                if (inColumn.Count > 0 && inColumn.All(c => TableCellWrapper.Span(c) == 1 && group.Any(g => ReferenceEquals(g.TableCell, c))))
                {
                    grid[index].Width = new TableCellWrapper(inColumn[0]).WidthTwips.ToString(CultureInfo.InvariantCulture);
                    changed = true;
                }
            }
            if (changed)
            {
                TableWrapper.ColumnsChanged(group.Key);
            }
        }
    }

    /// <summary>The height in points of the rows the cells are in, as <c>Cells.Height</c> in Word; null when they differ.</summary>
    public double? Height
    {
        get
        {
            IList<TableRowWrapper> rows = Rows;
            if (rows.Count == 0)
            {
                return null;
            }
            double first = rows[0].Height;
            return rows.All(r => r.Height == first) ? first : (double?)null;
        }
        set
        {
            if (value.HasValue)
            {
                foreach (TableRowWrapper row in Rows)
                {
                    row.Height = value.Value;
                }
            }
        }
    }

    /// <summary>How the height of the rows the cells are in is applied; null when they differ.</summary>
    public RowHeightRule? HeightRule
    {
        get
        {
            IList<TableRowWrapper> rows = Rows;
            if (rows.Count == 0)
            {
                return null;
            }
            RowHeightRule first = rows[0].HeightRule;
            return rows.All(r => r.HeightRule == first) ? first : (RowHeightRule?)null;
        }
        set
        {
            if (value.HasValue)
            {
                foreach (TableRowWrapper row in Rows)
                {
                    row.HeightRule = value.Value;
                }
            }
        }
    }

    /// <summary>Sets the height and its rule of the rows the cells are in, as <c>Cells.SetHeight</c> in Word.</summary>
    public void SetHeight(double rowHeight, RowHeightRule heightRule)
    {
        foreach (TableRowWrapper row in Rows)
        {
            row.SetHeight(rowHeight, heightRule);
        }
    }
}

public partial class TableCellWrapper
{
    /// <summary>The level of the table the cell is in: 1 for a table in the body, 2 for one inside a cell, and so on (<c>Cell.NestingLevel</c> in Word).</summary>
    public int NestingLevel
    {
        get { return Parent.Parent.NestingLevel; }
    }

    /// <summary>Sets the height and its rule of the row of the cell, as <c>Cell.SetHeight</c> in Word.</summary>
    public void SetHeight(double rowHeight, RowHeightRule heightRule)
    {
        Parent.SetHeight(rowHeight, heightRule);
    }
}

public partial class TableRowWrapper
{
    /// <summary>The level of the table the row is in (<c>Row.NestingLevel</c> in Word).</summary>
    public int NestingLevel
    {
        get { return Parent.NestingLevel; }
    }
}

public partial class TableRowWrapperCollection
{
    /// <summary>Deletes all the rows, as <c>Rows.Delete</c> in Word; deleting every row of a table deletes the table.</summary>
    public void Delete()
    {
        foreach (TableRowWrapper row in Items.Reverse())
        {
            row.Delete();
        }
    }

    /// <summary>Sets the height and its rule of all the rows, as <c>Rows.SetHeight</c> in Word.</summary>
    public void SetHeight(double rowHeight, RowHeightRule heightRule)
    {
        WriteAll(r => r.SetHeight(rowHeight, heightRule));
    }
}

public partial class TableColumnWrapper
{
    /// <summary>The level of the table the column is in (<c>Column.NestingLevel</c> in Word).</summary>
    public int NestingLevel
    {
        get { return Parent.NestingLevel; }
    }

    /// <summary>The column after this one, or null for the last (<c>Column.Next</c> in Word).</summary>
    public TableColumnWrapper Next()
    {
        return Parent.Columns.ElementAtOrDefault(Index + 1);
    }

    /// <summary>The column before this one, or null for the first (<c>Column.Previous</c> in Word).</summary>
    public TableColumnWrapper Previous()
    {
        return Index == 0 ? null : Parent.Columns.ElementAtOrDefault(Index - 1);
    }
}

public partial class TableColumnWrapperCollection
{
    private T? ReadAll<T>(Func<TableColumnWrapper, T> read) where T : struct
    {
        IList<TableColumnWrapper> columns = Items;
        if (columns.Count == 0)
        {
            return null;
        }
        T first = read(columns[0]);
        return columns.All(c => EqualityComparer<T>.Default.Equals(read(c), first)) ? first : (T?)null;
    }

    /// <summary>The level of the table the columns are in (<c>Columns.NestingLevel</c> in Word).</summary>
    public int NestingLevel
    {
        get { return Parent.NestingLevel; }
    }

    /// <summary>
    /// The width of the columns in points; null when they differ. Setting it sets every column, as <see cref="TableColumnWrapper.Width"/> does.
    /// </summary>
    public double? Width
    {
        get { return ReadAll(c => c.Width); }
        set
        {
            if (value.HasValue)
            {
                foreach (TableColumnWrapper column in Items)
                {
                    column.Width = value.Value;
                }
            }
        }
    }

    /// <summary>What <see cref="PreferredWidth"/> means: points, a percentage or automatic; null when the columns differ.</summary>
    public PreferredWidthType? PreferredWidthType
    {
        get { return ReadAll(c => c.PreferredWidthType); }
        set
        {
            if (value.HasValue)
            {
                foreach (TableColumnWrapper column in Items)
                {
                    column.PreferredWidthType = value.Value;
                }
            }
        }
    }

    /// <summary>
    /// The preferred width of the cells of the columns; null when they differ. In points it is the width of the columns, as in Word, so the
    /// grid and the cells follow.
    /// </summary>
    public double? PreferredWidth
    {
        get { return ReadAll(c => c.PreferredWidth); }
        set
        {
            if (!value.HasValue)
            {
                return;
            }
            foreach (TableColumnWrapper column in Items)
            {
                if (column.PreferredWidthType == Word.PreferredWidthType.Percent)
                {
                    column.PreferredWidth = value.Value;
                }
                else
                {
                    column.Width = value.Value;
                }
            }
        }
    }

    /// <summary>
    /// The shading of the cells of all the columns. When the columns are all those of a table the shading of the table is set too, as Word does.
    /// </summary>
    public ShadingWrapper Shading
    {
        get { return new ShadingWrapper(create => TableCellWrapperCollection.ShadingOf(Items.SelectMany(c => c.Cells).ToList(), create)); }
    }

    /// <summary>Deletes all the columns, as <c>Columns.Delete</c> in Word; deleting every column deletes the table.</summary>
    public void Delete()
    {
        foreach (TableColumnWrapper column in Items.Reverse())
        {
            column.Delete();
        }
    }
}
