using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableWrapperCollection : IEnumerable<TableWrapper>
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
        if (_rangeWrapper.Start is BookmarkStart)
        {
            //The tables touched by the bookmark. Where tables are nested only the outermost ones are returned, as for any other range.
            IList<Table> tables = _rangeWrapper.Start.GetElements<Table>(_rangeWrapper.End).ToList();
            if (tables.Count > 0)
            {
                int outermostLevel = tables.Min(e => e.NestingLevel<Table>());
                tables = tables.Where(e => e.NestingLevel<Table>() == outermostLevel).ToList();
            }
            return tables;
        }

        return _rangeWrapper.Start.DescendantsAtRelativeNestingLevel<Table>(1);
    }

    public int NestingLevel
    {
        get { return _rangeWrapper.Start.NestingLevel<Table>() + 1; }
    }

    #region IEnumerable Members

    public IEnumerator<TableWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (Table element in GetElementList())
        {
            yield return new TableWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
