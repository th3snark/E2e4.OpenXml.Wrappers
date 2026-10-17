using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public class TableRowWrapperCollection : IEnumerable<TableRowWrapper>
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
        Table table = _range.Start as Table;
        if (table == null)
        {
            throw new InvalidOperationException("Rows can only be added to a collection of the rows of a table.");
        }

        //Get the last row in the table.
        TableRow lastTableRow = table.Elements<TableRow>().LastOrDefault();
        if (lastTableRow == null)
        {
            throw new InvalidOperationException("A row cannot be added to a table that has no rows to use as a template.");
        }

        //Clone a copy of the last table row and add it as the new last row in the table. The clone must not carry any of the bookmark or
        //comment anchors of the original since their ids would be duplicated.
        TableRow tableRow = (TableRow)lastTableRow.Clone();
        tableRow.RemoveAnchors();
        tableRow.ClearText();
        lastTableRow.InsertAfterSelf(tableRow);

        return new TableRowWrapper(tableRow);
    }

    public int NestingLevel
    {
        get { return _range.Start.NestingLevel<TableRow>() + 1; }
    }

    #region IEnumerable Members

    public IEnumerator<TableRowWrapper> GetEnumerator()
    {
        //The elements are fetched once for the whole enumeration so that enumerating is linear and so that deleting the
        //current item while enumerating does not disturb the enumeration.
        foreach (TableRow element in GetElementList())
        {
            yield return new TableRowWrapper(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
