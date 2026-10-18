using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class TableRowWrapperCollection : IEnumerable<TableRowWrapper>
{
    private RangeWrapper _range;

    public TableRowWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    public TableWrapper Parent
    {
        get { return _range.StartElement is Table table ? new TableWrapper(table) : null; }
    }

    public TableRowWrapper this[int index]
    {
        get
        {
            try
            {
                return new TableRowWrapper(GetElementList()[index]);
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

    private IList<TableRow> GetElementList()
    {
        IList<TableRow> list = null;
        if (_range.StartElement is BookmarkStart)
        {
            list = new List<TableRow>();
            foreach (TableRow tableRow in _range.StartElement.GetElements<TableRow>(_range.EndElement))
            {
                list.Add(tableRow);
            }
        }
        else
        {
            list = _range.StartElement.DescendantsAtRelativeNestingLevel<TableRow>(1).ToList();
        }
        return list;
    }

    /// <summary>
    /// Adds an empty row before <paramref name="beforeRow"/>, as <c>Rows.Add(BeforeRow)</c> in Word. The new row has the same cells and
    /// row and cell formatting as that row. With null the row is added at the end, as <see cref="Add()"/> does.
    /// </summary>
    /// <exception cref="ArgumentException">The row is not a row of this table.</exception>
    public TableRowWrapper Add(TableRowWrapper beforeRow)
    {
        if (beforeRow == null)
        {
            return Add();
        }
        if (!(_range.StartElement is Table table) || beforeRow.TableRow.Parent != table)
        {
            throw new ArgumentException("The row must be a row of this table.", nameof(beforeRow));
        }

        TableRow tableRow = (TableRow)beforeRow.TableRow.Clone();
        tableRow.RemoveAnchors();
        tableRow.ClearText();
        beforeRow.TableRow.InsertBeforeSelf(tableRow);
        return new TableRowWrapper(tableRow);
    }

    public TableRowWrapper Add()
    {
        Table table = _range.StartElement as Table;
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
        get { return _range.StartElement.NestingLevel<TableRow>() + 1; }
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
