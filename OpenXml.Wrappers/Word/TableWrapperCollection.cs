using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
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
            Table table = _rangeWrapper.StartElement.ParentAtRelativeNestingLevel<Table>(_rangeWrapper.StartElement is Table ? -1 : 0);
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
        if (_rangeWrapper.StartElement is BookmarkStart)
        {
            //The tables touched by the bookmark. Where tables are nested only the outermost ones are returned, as for any other range.
            IList<Table> tables = _rangeWrapper.StartElement.GetElements<Table>(_rangeWrapper.EndElement).ToList();
            if (tables.Count > 0)
            {
                int outermostLevel = tables.Min(e => e.NestingLevel<Table>());
                tables = tables.Where(e => e.NestingLevel<Table>() == outermostLevel).ToList();
            }
            return tables;
        }

        return _rangeWrapper.StartElement.DescendantsAtRelativeNestingLevel<Table>(1);
    }

    /// <summary>
    /// Adds an empty table with the given number of rows and columns immediately before the range, as <c>Tables.Add</c> in Word, and
    /// returns it. The table has single 0.5 point borders and fills the width available to it.
    /// </summary>
    /// <remarks>
    /// A range of a paragraph, table, table cell (the table goes before its first block), bookmark or comment (before its paragraph),
    /// or of the whole document (at the start) is supported. Unlike Word the text of the range is not replaced.
    /// </remarks>
    /// <exception cref="NotSupportedException">The range is of a type that cannot take a table.</exception>
    public TableWrapper Add(RangeWrapper range, int numRows, int numColumns)
    {
        Guard.ThrowIfNull(range);
        Guard.ThrowIfLessThan(numRows, 1);
        Guard.ThrowIfLessThan(numColumns, 1);

        OpenXmlElement anchor = range.StartElement;
        if (anchor is BookmarkStart || anchor is CommentRangeStart)
        {
            anchor = anchor.Ancestors<Paragraph>().FirstOrDefault();
        }
        else if (anchor is TableCell cell)
        {
            anchor = cell.Elements().FirstOrDefault(e => e is Paragraph || e is Table);
        }
        else if (anchor is Document document)
        {
            anchor = document.Body;
        }

        if (anchor is Body body)
        {
            anchor = body.Elements().FirstOrDefault(e => !(e is SectionProperties));
            if (anchor == null)
            {
                //The table needs a paragraph after it, and the section properties must stay last.
                Paragraph paragraph = new Paragraph();
                SectionProperties sectionProperties = body.Elements<SectionProperties>().FirstOrDefault();
                if (sectionProperties != null)
                {
                    sectionProperties.InsertBeforeSelf(paragraph);
                }
                else
                {
                    body.AppendChild(paragraph);
                }
                anchor = paragraph;
            }
        }
        if (!(anchor is Paragraph || anchor is Table))
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Adding a table to a range starting with the type {0} is not currently supported.", range.StartElement.GetType().Name));
        }

        int totalWidth = GetAvailableWidth(anchor);
        int columnWidth = totalWidth / numColumns;
        Table table = new Table(
            new TableProperties(new TableWidth() { Width = totalWidth.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa }),
            new TableGrid(Enumerable.Range(0, numColumns).Select(i => new GridColumn() { Width = columnWidth.ToString(CultureInfo.InvariantCulture) })));
        for (int row = 0; row < numRows; row++)
        {
            table.AppendChild(new TableRow(Enumerable.Range(0, numColumns).Select(i => new TableCell(
                new TableCellProperties(new TableCellWidth() { Width = columnWidth.ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa }),
                new Paragraph()))));
        }

        anchor.InsertBeforeSelf(table);
        TableWrapper wrapper = new TableWrapper(table);
        wrapper.Borders.SetOutside(BorderLineStyle.SingleLine, 0.5, "auto");
        wrapper.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "auto");
        return wrapper;
    }

    /// <summary>
    /// The width, in twips, between the margins of the page or the width of the table cell that the element is in.
    /// </summary>
    private static int GetAvailableWidth(OpenXmlElement element)
    {
        TableCell cell = element.Ancestors<TableCell>().FirstOrDefault();
        TableWidthUnitValues? cellType = cell?.TableCellProperties?.TableCellWidth?.Type?.Value;
        if (cell != null && (cellType == null || cellType == TableWidthUnitValues.Dxa)
            && Int32.TryParse(cell.TableCellProperties?.TableCellWidth?.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int cellWidth) && cellWidth > 0)
        {
            return cellWidth;
        }

        //The section properties at the end of the body are those of the last section; Word's default page is 8.5 x 11 inches with 1 inch margins.
        SectionProperties sectionProperties = element.Ancestors<Body>().LastOrDefault()?.Elements<SectionProperties>().LastOrDefault();
        uint pageWidth = sectionProperties?.GetFirstChild<PageSize>()?.Width?.Value ?? 12240;
        PageMargin margin = sectionProperties?.GetFirstChild<PageMargin>();
        long left = margin?.Left?.Value ?? 1440;
        long right = margin?.Right?.Value ?? 1440;
        long width = pageWidth - left - right;
        return width > 0 ? (int)width : 9360;
    }

    public int NestingLevel
    {
        get { return _rangeWrapper.StartElement.NestingLevel<Table>() + 1; }
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
