using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableColumnWrapperCollection : IEnumerable
{
    private RangeWrapper _range;

    public TableColumnWrapperCollection(RangeWrapper range)
    {
        _range = range;
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
        return _range.Start.Elements<TableGrid>().First().Elements<GridColumn>().ToList();
    }

    public TableColumnWrapper Add()
    {
        //Get the table grid of table.
        TableGrid tableGrid = ((Table)_range.Start).Elements<TableGrid>().First();

        //Clone a copy of the last grid column and add it as the new last grid column to table grid.
        GridColumn gridColumn = (GridColumn)tableGrid.LastChild.Clone();
        tableGrid.Append(gridColumn);

        //Add a cloned copy of the last table cell in each table row to the end of each table row.
        foreach (TableRow tableRow in ((Table)_range.Start).Elements<TableRow>())
        {
            TableCell lastTableCell = (TableCell)tableRow.Elements<TableCell>().Last();
            TableCell tableCell = (TableCell)lastTableCell.Clone();
            lastTableCell.InsertAfterSelf(tableCell);
        }

        return new TableColumnWrapper(gridColumn);
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new GridColumnWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class GridColumnWrapperEnumerator : IEnumerator, IDisposable
    {
        private TableColumnWrapperCollection _commentWrapperCollection;
        private int _index;
        private TableColumnWrapper _current;

        public GridColumnWrapperEnumerator(TableColumnWrapperCollection commentWrapperCollection)
        {
            _commentWrapperCollection = commentWrapperCollection;
            _index = -1;
            _current = null;
        }

        #region IEnumerator Members

        public object Current
        {
            get { return _current; }
        }

        public bool MoveNext()
        {
            if (++_index >= _commentWrapperCollection.Count)
                return false;
            else
                _current = _commentWrapperCollection[_index];
            return true;
        }

        public void Reset()
        {
            _current = null;
            _index = -1;
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            _commentWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
