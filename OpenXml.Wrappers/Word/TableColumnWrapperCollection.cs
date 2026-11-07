using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class TableColumnWrapperCollection : IEnumerable<TableColumnWrapper>
{
    private RangeWrapper _range;

    public TableColumnWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>
    /// The table the columns belong to, or null when the collection is not the columns of a table.
    /// </summary>
    public TableWrapper Parent
    {
        get { return _range.StartElement is Table table ? new TableWrapper(table) : null; }
    }

    public TableColumnWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableColumnWrapper(GetElementList()[index]);
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    private IList<GridColumn> GetElementList()
    {
        //A table without a grid has no columns.
        TableGrid grid = _range.StartElement.Elements<TableGrid>().FirstOrDefault();
        return grid == null ? new List<GridColumn>() : grid.Elements<GridColumn>().ToList();
    }

    public TableColumnWrapper Add()
    {
        Table table = _range.StartElement as Table;
        if (table == null)
        {
            throw new InvalidOperationException("Columns can only be added to a collection of the columns of a table.");
        }

        //Get the table grid of table.
        TableGrid tableGrid = table.Elements<TableGrid>().FirstOrDefault();
        if (tableGrid == null || tableGrid.LastChild == null)
        {
            throw new InvalidOperationException("A column cannot be added to a table that has no columns to use as a template.");
        }

        //Clone a copy of the last grid column and add it as the new last grid column to table grid.
        GridColumn gridColumn = (GridColumn)tableGrid.LastChild.Clone();
        tableGrid.Append(gridColumn);
        int gridWidth = gridColumn.Width == null ? 0 : Convert.ToInt32(gridColumn.Width.Value, CultureInfo.InvariantCulture);

        //Add an empty copy of the last table cell in each table row to the end of each table row. The new cell spans just the new grid
        //column, has the width of that column, and does not carry over the text or the bookmark and comment anchors of the original.
        foreach (TableRow tableRow in table.Elements<TableRow>())
        {
            TableCell lastTableCell = tableRow.Elements<TableCell>().LastOrDefault();
            if (lastTableCell == null)
            {
                continue;
            }

            TableCell tableCell = (TableCell)lastTableCell.Clone();
            tableCell.RemoveAnchors();
            tableCell.ClearText();
            tableCell.SetGridSpan(1);
            if (gridWidth > 0)
            {
                tableCell.SetWidth(gridWidth, TableWidthUnitValues.Dxa);
            }
            lastTableCell.InsertAfterSelf(tableCell);
        }

        TableWrapper.ColumnsChanged(table);
        return new TableColumnWrapper(gridColumn);
    }

    /// <summary>
    /// Inserts columns before a column of the grid of a table (after the last one when the index is the number of columns). The new
    /// columns have the width of the column at <paramref name="templateIndex"/>. A row where a cell spans the place gets no cell: that
    /// cell grows. Returns the new cells, row by row.
    /// </summary>
    internal static List<TableCell> InsertColumns(Table table, int index, int count, int templateIndex)
    {
        TableGrid grid = table.Elements<TableGrid>().FirstOrDefault();
        List<GridColumn> columns = grid == null ? new List<GridColumn>() : grid.Elements<GridColumn>().ToList();
        if (columns.Count == 0)
        {
            throw new InvalidOperationException("A column cannot be added to a table that has no columns to use as a template.");
        }
        index = Math.Min(index, columns.Count);
        GridColumn template = columns[Math.Min(templateIndex, columns.Count - 1)];
        for (int i = 0; i < count; i++)
        {
            GridColumn clone = (GridColumn)template.Clone();
            if (index < columns.Count)
            {
                columns[index].InsertBeforeSelf(clone);
            }
            else
            {
                grid.Append(clone);
            }
        }

        List<TableCell> added = new List<TableCell>();
        foreach (TableRow row in table.Elements<TableRow>())
        {
            List<TableCell> cells = row.Elements<TableCell>().ToList();
            if (cells.Count == 0)
            {
                continue;
            }
            TableCell before = cells.FirstOrDefault(c => TableCellWrapper.GetGridStart(c) == index);
            TableCell spanning = before != null ? null : cells.FirstOrDefault(c => TableCellWrapper.GetGridStart(c) < index && index < TableCellWrapper.GetGridStart(c) + TableCellWrapper.Span(c));
            if (spanning != null)
            {
                spanning.SetGridSpan(TableCellWrapper.Span(spanning) + count);
                continue;
            }
            TableCell source = before ?? cells[cells.Count - 1];
            for (int i = 0; i < count; i++)
            {
                TableCell cell = (TableCell)source.Clone();
                cell.RemoveAnchors();
                cell.ClearText();
                cell.SetGridSpan(1);
                if (before != null)
                {
                    before.InsertBeforeSelf(cell);
                }
                else
                {
                    cells[cells.Count - 1].InsertAfterSelf(cell);
                }
                added.Add(cell);
            }
        }

        TableWrapper.ColumnsChanged(table);
        return added;
    }

    #region IEnumerable Members

    public IEnumerator<TableColumnWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (GridColumn element in GetElementList())
        {
            yield return new TableColumnWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
