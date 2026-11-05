using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
            catch (ArgumentOutOfRangeException)
            {
                throw new ArgumentOutOfRangeException(nameof(index), string.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
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
    /// Adds an empty table with the given number of rows and columns at the range, as <c>Tables.Add</c> in Word, and returns it. The table has
    /// single 0.5 point borders and fills the width available to it.
    /// </summary>
    /// <remarks>
    /// <para>As in Word the table replaces the range: the text of a span, paragraph, table cell or bookmark is removed (the paragraph that ends it
    /// stays, after the table), so collapse the range first to keep it. A span in the middle of a paragraph splits the paragraph. A
    /// collapsed span, a table, a comment range or a document or header that is empty gets the table in front of it. The whole
    /// document (<c>Content</c>) is replaced, as in Word.</para>
    /// <para><paramref name="defaultTableBehavior"/> <see cref="DefaultTableBehavior.Word9TableBehavior"/> lets the columns resize to fit their
    /// content; <see cref="AutoFitBehavior.Content"/> and <see cref="AutoFitBehavior.Window"/> size the table to its content or to the window
    /// (the width between the margins).</para>
    /// </remarks>
    /// <exception cref="NotSupportedException">The range is of a type that cannot take a table.</exception>
    public TableWrapper Add(RangeWrapper range, int numRows, int numColumns, DefaultTableBehavior defaultTableBehavior = DefaultTableBehavior.Word8TableBehavior,
        AutoFitBehavior autoFitBehavior = AutoFitBehavior.Fixed)
    {
        Guard.ThrowIfNull(range);
        Guard.ThrowIfLessThan(numRows, 1);
        Guard.ThrowIfLessThan(numColumns, 1);
        if (!EnumCompat.IsDefined(defaultTableBehavior))
        {
            throw new ArgumentOutOfRangeException(nameof(defaultTableBehavior));
        }
        if (!EnumCompat.IsDefined(autoFitBehavior))
        {
            throw new ArgumentOutOfRangeException(nameof(autoFitBehavior));
        }

        OpenXmlElement anchor;
        if (range.IsSpan)
        {
            anchor = range.PrepareInsertBlock();
        }
        else
        {
            anchor = range.StartElement;
            if (anchor is BookmarkStart || anchor is CommentRangeStart)
            {
                if (anchor is BookmarkStart)
                {
                    range.Text = string.Empty;
                }
                anchor = anchor.Ancestors<Paragraph>().FirstOrDefault();
            }
            else if (anchor is Paragraph)
            {
                range.Text = string.Empty;
            }
            else if (anchor is TableCell cell)
            {
                range.Text = string.Empty;
                anchor = cell.Elements().FirstOrDefault(e => e is Paragraph || e is Table);
            }
            else if (anchor is Document || anchor is Body)
            {
                range.Delete();
                anchor = (anchor is Document document ? document.Body : (Body)anchor).Elements().FirstOrDefault(e => !(e is SectionProperties));
            }
            else if (anchor is Header || anchor is Footer)
            {
                range.Delete();
                anchor = anchor.Elements().FirstOrDefault(e => e is Paragraph || e is Table);
            }
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
            throw new NotSupportedException(string.Format(CultureInfo.InvariantCulture, "Adding a table to a range starting with the type {0} is not currently supported.", range.IsSpan ? "span" : range.StartElement.GetType().Name));
        }

        Table table = CreateTable(anchor, numRows, numColumns, defaultTableBehavior == DefaultTableBehavior.Word9TableBehavior, autoFitBehavior);
        anchor.InsertBeforeSelf(table);
        return new TableWrapper(table);
    }

    //The table as Word writes it. With the behaviour of Word 8, which is Tables.Add's default, it has no style and no borders, and fixed
    //columns whatever is asked for. With that of Word 2000 and later it is in Table Grid and 10 twips narrower than the space it is in, and the
    //columns can resize. Fitting the window gives it a width of 100%, and fitting the content (Word 2000 and later) leaves the width to the
    //content.
    private static Table CreateTable(OpenXmlElement anchor, int numRows, int numColumns, bool word9, AutoFitBehavior autoFit)
    {
        int width = word9 ? GetSpaceForTable(anchor) : GetAvailableWidth(anchor);
        bool window = autoFit == AutoFitBehavior.Window;
        bool content = word9 && autoFit == AutoFitBehavior.Content;
        int[] columns = Enumerable.Range(0, numColumns).Select(i => content ? ContentColumnWidth : Share(width, i, numColumns)).ToArray();

        TableProperties properties = new TableProperties();
        if (word9)
        {
            MainDocumentPart mainDocumentPart = StyleHelper.GetMainDocumentPart(anchor);
            if (mainDocumentPart != null)
            {
                BuiltInStyles.GetOrCreate(mainDocumentPart, "TableGrid");
            }
            properties.TableStyle = new TableStyle() { Val = "TableGrid" };
        }
        properties.TableWidth = window
            ? new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct }
            : new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
        if (!word9)
        {
            properties.TableLayout = new TableLayout() { Type = TableLayoutValues.Fixed };
        }
        properties.TableLook = word9
            ? new TableLook() { Val = "04A0", FirstRow = OpenXmlExtensions.OnValue(), LastRow = OpenXmlExtensions.OffValue(), FirstColumn = OpenXmlExtensions.OnValue(), LastColumn = OpenXmlExtensions.OffValue(), NoHorizontalBand = OpenXmlExtensions.OffValue(), NoVerticalBand = OpenXmlExtensions.OnValue() }
            : new TableLook() { Val = "0000", FirstRow = OpenXmlExtensions.OffValue(), LastRow = OpenXmlExtensions.OffValue(), FirstColumn = OpenXmlExtensions.OffValue(), LastColumn = OpenXmlExtensions.OffValue(), NoHorizontalBand = OpenXmlExtensions.OffValue(), NoVerticalBand = OpenXmlExtensions.OffValue() };

        Table table = new Table(properties, new TableGrid(columns.Select(w => new GridColumn() { Width = w.ToString(CultureInfo.InvariantCulture) })));
        for (int row = 0; row < numRows; row++)
        {
            TableRow tableRow = new TableRow();
            if (!word9)
            {
                tableRow.AppendChild(new TablePropertyExceptions(new TableCellMarginDefault(
                    new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa },
                    new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa })));
            }
            for (int column = 0; column < numColumns; column++)
            {
                TableCellWidth cellWidth = window
                    ? new TableCellWidth() { Width = Share(5000, column, numColumns).ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Pct }
                    : content
                        ? new TableCellWidth() { Width = "0", Type = TableWidthUnitValues.Auto }
                        : new TableCellWidth() { Width = columns[column].ToString(CultureInfo.InvariantCulture), Type = TableWidthUnitValues.Dxa };
                tableRow.AppendChild(new TableCell(new TableCellProperties(cellWidth), new Paragraph()));
            }
            table.AppendChild(tableRow);
        }
        return table;
    }

    //The width Word gives each column of a table that fits its content, before there is any content.
    private const int ContentColumnWidth = 222;

    //The left and right cell margins of the default table style, which a nested table has to fit between.
    private const int DefaultCellMargins = 216;

    //The width of a column when a total is shared out: the columns end at whole numbers, so the later ones take up what is left over.
    private static int Share(int total, int index, int count)
    {
        return (int)(((long)total * (index + 1) / count) - ((long)total * index / count));
    }

    /// <summary>
    /// The width, in twips, that Word gives a table at the element and measures its columns against: 10 less than <see cref="GetAvailableWidth"/>.
    /// </summary>
    internal static int GetSpaceForTable(OpenXmlElement element)
    {
        return Math.Max(GetAvailableWidth(element) - 10, 1);
    }

    /// <summary>
    /// The width, in twips, between the margins of the page or inside the margins of the table cell that the element is in.
    /// </summary>
    private static int GetAvailableWidth(OpenXmlElement element)
    {
        TableCell cell = element.Ancestors<TableCell>().FirstOrDefault();
        TableWidthUnitValues? cellType = cell?.TableCellProperties?.TableCellWidth?.Type?.Value;
        if (cell != null && (cellType == null || cellType == TableWidthUnitValues.Dxa)
            && int.TryParse(cell.TableCellProperties?.TableCellWidth?.Width?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int cellWidth) && cellWidth > 0)
        {
            return Math.Max(cellWidth - DefaultCellMargins, 0);
        }

        //A section ends with the paragraph that holds its properties, and those at the end of the body are of the last section. Word's default
        //page is 8.5 x 11 inches with 1 inch margins.
        Body body = element.Ancestors<Body>().LastOrDefault();
        OpenXmlElement block = element.Parent is Body ? element : element.Ancestors().FirstOrDefault(e => e.Parent is Body);
        SectionProperties sectionProperties = null;
        for (OpenXmlElement next = block; next != null && sectionProperties == null; next = next.NextSibling())
        {
            sectionProperties = (next as Paragraph)?.ParagraphProperties?.GetFirstChild<SectionProperties>();
        }
        sectionProperties = sectionProperties ?? body?.Elements<SectionProperties>().LastOrDefault();
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
