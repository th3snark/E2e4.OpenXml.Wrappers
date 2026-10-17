using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableWrapperCollection : IEnumerable
{
    private RangeWrapper _rangeWrapper;

    public TableWrapperCollection(RangeWrapper rangeWrapper)
    {
        _rangeWrapper = rangeWrapper;
    }

    public TableWrapper Parent
    {
        get
        {
            //A table is itself at the nesting level of the collection's parent so its parent table is one level out.
            Table table = _rangeWrapper.Start.ParentAtRelativeNestingLevel<Table>(_rangeWrapper.Start is Table ? -1 : 0);
            return table != null ? new TableWrapper(table) : null;
        }
    }

    public TableWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableWrapper(GetElementList()[index]);
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

    private IList<Table> GetElementList()
    {
        return _rangeWrapper.Start.DescendantsAtRelativeNestingLevel<Table>(1);
    }

    public int NestingLevel
    {
        get { return _rangeWrapper.Start.NestingLevel<Table>() + 1; }
    }

    #region IEnumerable Members

    public IEnumerator GetEnumerator()
    {
        return new TableWrapperEnumerator(this);
    }

    #endregion

    #region Nested Private Classes

    private sealed class TableWrapperEnumerator : IEnumerator, IDisposable
    {
        private TableWrapperCollection _tableWrapperCollection;
        private int _index;
        private TableWrapper _current;

        public TableWrapperEnumerator(TableWrapperCollection tableWrapperCollection)
        {
            _tableWrapperCollection = tableWrapperCollection;
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
            if (++_index >= _tableWrapperCollection.Count)
                return false;
            else
                _current = _tableWrapperCollection[_index];
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
            _tableWrapperCollection = null;
            _current = null;
            _index = -1;
        }

        #endregion
    }

    #endregion
}
