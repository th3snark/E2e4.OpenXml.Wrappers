using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What separates the text of cells, or of the fields of a line, when a table is converted to text or text to a table, as
/// <c>WdTableFieldSeparator</c> in Word (with the same values).
/// </summary>
public enum TableFieldSeparator
{
    /// <summary>Each cell, or each field, is a paragraph.</summary>
    Paragraphs = 0,

    /// <summary>A tab.</summary>
    Tabs = 1,

    /// <summary>A comma and a space.</summary>
    Commas = 2,

    /// <summary>The default separator, a hyphen as Word 16 writes it.</summary>
    Default = 3
}

/// <summary>
/// What a column is sorted as, as <c>WdSortFieldType</c> in Word (with the same values for these).
/// </summary>
public enum SortFieldType
{
    /// <summary>As text, which the current culture orders and without regard to case.</summary>
    Alphanumeric = 0,

    /// <summary>As the first number in the text.</summary>
    Numeric = 1,

    /// <summary>As dates.</summary>
    Date = 2
}

/// <summary>
/// The direction of a sort, as <c>WdSortOrder</c> in Word (with the same values).
/// </summary>
public enum SortOrder
{
    Ascending = 0,
    Descending = 1
}

public partial class TableWrapper
{
    //What a separator is as text; null for paragraphs and a tab.
    private static string SeparatorText(TableFieldSeparator separator)
    {
        if (!EnumCompat.IsDefined(separator))
        {
            throw new ArgumentOutOfRangeException(nameof(separator));
        }
        return separator == TableFieldSeparator.Commas ? ", " : separator == TableFieldSeparator.Default ? "-" : null;
    }

    /// <summary>
    /// Replaces the table with its text, as <c>Table.ConvertToText</c> in Word. With <see cref="TableFieldSeparator.Paragraphs"/> each paragraph
    /// of a cell is a paragraph; otherwise each row is a paragraph with the text of its cells one after the other, divided by the separator. The
    /// result is the new text.
    /// </summary>
    /// <param name="separator">What divides the text of the cells.</param>
    /// <param name="nestedTables">Whether tables inside the cells are converted too. Converting them is not supported.</param>
    /// <exception cref="NotSupportedException">The table has tables inside its cells.</exception>
    public RangeWrapper ConvertToText(TableFieldSeparator separator = TableFieldSeparator.Paragraphs, bool nestedTables = true)
    {
        return ConvertToText(separator == TableFieldSeparator.Paragraphs, separator == TableFieldSeparator.Tabs, SeparatorText(separator), nestedTables);
    }

    /// <summary>
    /// Replaces the table with its text, the text of the cells of a row divided by <paramref name="separator"/>, as <c>Table.ConvertToText</c> in Word.
    /// </summary>
    /// <exception cref="ArgumentException">The separator is empty.</exception>
    /// <exception cref="NotSupportedException">The table has tables inside its cells.</exception>
    public RangeWrapper ConvertToText(string separator, bool nestedTables = true)
    {
        if (string.IsNullOrEmpty(separator))
        {
            throw new ArgumentException("The separator cannot be empty.", nameof(separator));
        }
        return ConvertToText(false, false, separator, nestedTables);
    }

    private RangeWrapper ConvertToText(bool paragraphs, bool tabs, string separator, bool nestedTables)
    {
        if (_table.Descendants<Table>().Any())
        {
            throw new NotSupportedException("Converting a table that has tables inside its cells is not supported.");
        }

        int margin = (int)Math.Round(LeftPadding * 20, MidpointRounding.AwayFromZero);
        List<int> grid = TableWidths.ReadGrid(_table);
        List<Paragraph> produced = new List<Paragraph>();
        foreach (TableRow row in _table.Elements<TableRow>().ToList())
        {
            List<TableCell> cells = row.Elements<TableCell>().ToList();
            if (paragraphs)
            {
                produced.AddRange(cells.SelectMany(c => c.Elements<Paragraph>()).ToList());
                continue;
            }

            Paragraph paragraph = new Paragraph();
            if (tabs && cells.Count > 1)
            {
                //A tab stop at the left edge of each cell after the first, so that the text lines up as it did.
                Tabs stops = new Tabs();
                int position = 0;
                foreach (TableCell cell in cells.Take(cells.Count - 1))
                {
                    position += TableWidths.EffectiveWidth(cell, grid);
                    stops.AppendChild(new TabStop() { Val = TabStopValues.Left, Position = position + margin });
                }
                paragraph.AppendChild(new ParagraphProperties(stops, new Indentation() { Left = margin.ToString(CultureInfo.InvariantCulture) }));
            }
            for (int index = 0; index < cells.Count; index++)
            {
                if (index > 0)
                {
                    paragraph.AppendChild(tabs ? new Run(new TabChar()) : new Run(OpenXmlExtensions.NewText(separator)));
                }
                bool first = true;
                foreach (Paragraph source in cells[index].Elements<Paragraph>().ToList())
                {
                    if (!first)
                    {
                        paragraph.AppendChild(new Run(OpenXmlExtensions.NewText(" ")));
                    }
                    first = false;
                    foreach (OpenXmlElement child in source.ChildElements.Where(e => !(e is ParagraphProperties)).ToList())
                    {
                        child.Remove();
                        paragraph.AppendChild(child);
                    }
                }
            }
            paragraph.CoalesceRuns();
            produced.Add(paragraph);
        }

        foreach (Paragraph paragraph in produced)
        {
            paragraph.Parent?.RemoveChild(paragraph);
            _table.InsertBeforeSelf(paragraph);
        }
        RangeWrapper.DeleteElements(new OpenXmlElement[] { _table });
        return produced.Count == 0 ? null : new RangeWrapper(produced[0], produced[produced.Count - 1]);
    }

    /// <summary>
    /// Sorts the rows by the text of one column, as <c>Table.Sort</c> in Word. Rows that are equal keep their order.
    /// </summary>
    /// <param name="excludeHeader">Whether the first row is left where it is.</param>
    /// <param name="fieldNumber">The column to sort by (0-based).</param>
    /// <param name="sortFieldType">How the text is compared.</param>
    /// <param name="sortOrder">Ascending or descending.</param>
    /// <exception cref="ArgumentOutOfRangeException">The column is not in the table.</exception>
    /// <exception cref="NotSupportedException">The table has cells that are merged.</exception>
    public void Sort(bool excludeHeader = false, int fieldNumber = 0, SortFieldType sortFieldType = SortFieldType.Alphanumeric, SortOrder sortOrder = SortOrder.Ascending)
    {
        if (!EnumCompat.IsDefined(sortFieldType))
        {
            throw new ArgumentOutOfRangeException(nameof(sortFieldType));
        }
        if (!EnumCompat.IsDefined(sortOrder))
        {
            throw new ArgumentOutOfRangeException(nameof(sortOrder));
        }
        int columns = TableWidths.ReadGrid(_table).Count;
        if (fieldNumber < 0 || fieldNumber >= columns)
        {
            throw new ArgumentOutOfRangeException(nameof(fieldNumber));
        }
        if (_table.Descendants<TableCell>().Any(c => c.TableCellProperties?.VerticalMerge != null || TableCellWrapper.Span(c) != 1))
        {
            throw new NotSupportedException("A table with merged cells cannot be sorted.");
        }

        List<TableRow> rows = _table.Elements<TableRow>().ToList();
        TableRow header = excludeHeader ? rows.FirstOrDefault() : null;
        List<TableRow> body = rows.Where(r => r != header).ToList();
        Func<TableRow, string> text = r => string.Concat(TableCellWrapper.CellAtGridColumn(r, fieldNumber)?.Descendants<Text>().Select(t => t.Text) ?? Enumerable.Empty<string>());
        IEnumerable<TableRow> sorted;
        if (sortFieldType == SortFieldType.Numeric)
        {
            sorted = Order(body, r => FirstNumber(text(r)), Comparer<double?>.Default, sortOrder);
        }
        else if (sortFieldType == SortFieldType.Date)
        {
            sorted = Order(body, r => DateTime.TryParse(text(r), CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime date) ? date : (DateTime?)null, Comparer<DateTime?>.Default, sortOrder);
        }
        else
        {
            sorted = Order(body, text, StringComparer.CurrentCultureIgnoreCase, sortOrder);
        }

        List<TableRow> ordered = sorted.ToList();
        foreach (TableRow row in body)
        {
            row.Remove();
        }
        foreach (TableRow row in ordered)
        {
            _table.AppendChild(row);
        }
    }

    private static IEnumerable<TableRow> Order<TKey>(IEnumerable<TableRow> rows, Func<TableRow, TKey> key, IComparer<TKey> comparer, SortOrder order)
    {
        return order == SortOrder.Ascending ? rows.OrderBy(key, comparer) : rows.OrderByDescending(key, comparer);
    }

    //The first number in some text, or null when there is none.
    private static double? FirstNumber(string text)
    {
        Match match = Regex.Match(text, @"-?\d[\d,]*(\.\d+)?");
        return match.Success && double.TryParse(match.Value.Replace(",", string.Empty), NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : (double?)null;
    }
}

public partial class TableRowWrapper
{
    /// <summary>
    /// Replaces the row with its text, as <c>Row.ConvertToText</c> in Word: the table is split around it, and the text of the row is the
    /// paragraph between the tables. See <see cref="TableWrapper.ConvertToText(TableFieldSeparator, bool)"/>.
    /// </summary>
    /// <exception cref="NotSupportedException">The row has tables inside its cells.</exception>
    public RangeWrapper ConvertToText(TableFieldSeparator separator = TableFieldSeparator.Paragraphs, bool nestedTables = true)
    {
        return TableRowWrapperCollection.ConvertToText(new List<TableRow> { _tableRow }, table => table.ConvertToText(separator, nestedTables));
    }
}

public partial class TableRowWrapperCollection
{
    /// <summary>
    /// Replaces the rows with their text, as <c>Rows.ConvertToText</c> in Word: the table is split around the rows, which become the
    /// paragraphs between the tables. See <see cref="TableWrapper.ConvertToText(TableFieldSeparator, bool)"/>.
    /// </summary>
    /// <exception cref="NotSupportedException">The rows have tables inside their cells.</exception>
    public RangeWrapper ConvertToText(TableFieldSeparator separator = TableFieldSeparator.Paragraphs, bool nestedTables = true)
    {
        return ConvertToText(Items.Select(r => r.TableRow).ToList(), table => table.ConvertToText(separator, nestedTables));
    }

    //Splits the table before the first row and after the last, and converts what is between.
    internal static RangeWrapper ConvertToText(IList<TableRow> rows, Func<TableWrapper, RangeWrapper> convert)
    {
        if (rows.Count == 0)
        {
            return null;
        }
        Table table = (Table)rows[0].Parent;
        List<TableRow> all = table.Elements<TableRow>().ToList();
        int first = all.IndexOf(rows[0]);
        int last = all.IndexOf(rows[rows.Count - 1]);
        TableWrapper middle = new TableWrapper(table);
        Paragraph before = null;
        Paragraph after = null;
        if (first > 0)
        {
            middle = middle.Split(new TableRowWrapper(rows[0]));
            before = (Paragraph)middle.Table.PreviousSibling();
        }
        if (last < all.Count - 1)
        {
            TableRow next = all[last + 1];
            middle.Split(new TableRowWrapper(next));
            after = (Paragraph)middle.Table.NextSibling();
        }
        RangeWrapper result = convert(middle);

        //The paragraphs that Split put between the tables are not wanted: the text takes their place.
        before?.Remove();
        after?.Remove();
        return result;
    }
}
