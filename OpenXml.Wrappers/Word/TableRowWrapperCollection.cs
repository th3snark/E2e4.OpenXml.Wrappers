using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableRowWrapperCollection : IEnumerable
{
    private RangeWrapper _range;

    public TableRowWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    public TableWrapper Parent
    {
        get { return new TableWrapper((Table)_range.Start); }
    }

    public TableRowWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableRowWrapper(GetElementList()[index]);
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

    private IList<TableRow> GetElementList()
    {
        IList<TableRow> list = null;
        if (_range.Start is BookmarkStart)
        {
            list = new List<TableRow>();
            foreach (TableRow tableRow in _range.Start.GetElements<TableRow>(_range.End))
            {
                list.Add(tableRow);
            }
        }
        else
        {
            list = _range.Start.DescendantsAtRelativeNestingLevel<TableRow>(1).ToList();
        }
        return list;
    }

    public TableRowWrapper Add()
    {
        //Get the last row in the table.
        TableRow lastTableRow = ((Table)_range.Start).Elements<TableRow>().Last();

        //Clone a copy of the last table row and add it as the new last row in the table.
        TableRow tableRow = (TableRow)lastTableRow.Clone();
        tableRow.ClearText();
        lastTableRow.InsertAfterSelf(tableRow);

        return new TableRowWrapper(tableRow);
    }

    public int NestingLevel
    {
        get { return _range.Start.NestingLevel<TableRow>() + 1; }
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new TableRowWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class TableRowWrapperEnumerator : IEnumerator, IDisposable
    {
        private TableRowWrapperCollection _tableRowWrapperCollection;
        private int _index;
        private TableRowWrapper _current;

        public TableRowWrapperEnumerator(TableRowWrapperCollection tableRowWrapperCollection)
        {
            _tableRowWrapperCollection = tableRowWrapperCollection;
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
            if (++_index >= _tableRowWrapperCollection.Count)
                return false;
            else
                _current = _tableRowWrapperCollection[_index];
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
            _tableRowWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
