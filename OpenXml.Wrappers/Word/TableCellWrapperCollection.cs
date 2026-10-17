using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableCellWrapperCollection : IEnumerable<TableCellWrapper>
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

    public IEnumerator<TableCellWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (TableCell element in GetElementList())
        {
            yield return new TableCellWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
