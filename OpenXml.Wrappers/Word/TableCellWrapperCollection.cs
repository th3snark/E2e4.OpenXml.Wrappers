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
        get { return _range.StartElement is TableRow row ? new TableRowWrapper(row) : null; }
    }

    public TableCellWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableCellWrapper(GetElementList()[index]);
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new ArgumentOutOfRangeException(nameof(index), String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
        }
    }

    public int Count
    {
        get { return GetElementList().Count; }
    }

    public int NestingLevel
    {
        get { return _range.StartElement.NestingLevel<TableCell>() + 1; }
    }

    private IList<TableCell> GetElementList()
    {
        IList<TableCell> list = null;
        if (_range.StartElement is BookmarkStart)
        {
            list = new List<TableCell>();
            foreach (TableCell tableCell in _range.StartElement.GetElements<TableCell>(_range.EndElement))
            {
                list.Add(tableCell);
            }
        }
        else
        {
            list = _range.StartElement.DescendantsAtRelativeNestingLevel<TableCell>(1);
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
