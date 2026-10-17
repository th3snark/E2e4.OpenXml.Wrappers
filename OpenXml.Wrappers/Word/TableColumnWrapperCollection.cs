using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableColumnWrapperCollection : IEnumerable<TableColumnWrapper>
{
    private RangeWrapper _range;

    public TableColumnWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>
    /// The table the columns belong to.
    /// </summary>
    public TableWrapper Parent
    {
        get { return new TableWrapper((Table)_range.StartElement); }
    }

    public TableColumnWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableColumnWrapper(GetElementList()[index]);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new ArgumentOutOfRangeException(String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index), ex);
            }
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    private IList<GridColumn> GetElementList()
    {
        return _range.StartElement.Elements<TableGrid>().First().Elements<GridColumn>().ToList();
    }

    public TableColumnWrapper Add()
    {
        Table table = _range.StartElement as Table;
        if (table == null)
        {
            throw new InvalidOperationException("Columns can only be added to a collection of the columns of a table.");
        }

        //Get the table grid of table.
        TableGrid tableGrid = table.Elements<TableGrid>().First();
        if (tableGrid.LastChild == null)
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

        return new TableColumnWrapper(gridColumn);
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
