using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableCellWrapperCollection : IEnumerable
{
    private RangeWrapper _range;

    public TableCellWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    public TableRowWrapper Parent
    {
        get { return new TableRowWrapper((TableRow)_range.Start); }
    }

    public TableCellWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableCellWrapper(GetElementList()[index]);
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

    public int NestingLevel
    {
        get { return _range.Start.NestingLevel<TableCell>() + 1; }
    }

    private IList<TableCell> GetElementList()
    {
        IList<TableCell> list = null;
        if (_range.Start is BookmarkStart)
        {
            list = new List<TableCell>();
            foreach (TableCell tableCell in _range.Start.GetElements<TableCell>(_range.End))
            {
                list.Add(tableCell);
            }
        }
        else
        {
            list = _range.Start.DescendantsAtRelativeNestingLevel<TableCell>(1);
        }
        return list;
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new TableCellWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class TableCellWrapperEnumerator : IEnumerator, IDisposable
    {
        private TableCellWrapperCollection _tableCellWrapperCollection;
        private int _index;
        private TableCellWrapper _current;

        public TableCellWrapperEnumerator(TableCellWrapperCollection tableCellWrapperCollection)
        {
            _tableCellWrapperCollection = tableCellWrapperCollection;
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
            if (++_index >= _tableCellWrapperCollection.Count)
                return false;
            else
                _current = _tableCellWrapperCollection[_index];
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
            _tableCellWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
