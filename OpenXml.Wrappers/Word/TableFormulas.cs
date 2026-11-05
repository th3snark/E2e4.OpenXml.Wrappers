using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

//The formulas Word puts in table cells that can be worked out here: a function of the numbers above, below, to the left or to the right of
//the cell, as in =SUM(ABOVE). Anything else (arithmetic on cell references, bookmarks) is left for Word to work out when the document is opened.
internal static class TableFormulas
{
    internal const string UnexpectedEnd = "!Unexpected End of Formula";

    private static string CellText(TableCell cell)
    {
        return string.Concat(cell.Descendants<Text>().Select(t => t.Text)).Trim();
    }

    private static bool TryNumber(TableCell cell, out double number)
    {
        return double.TryParse(CellText(cell), NumberStyles.Number | NumberStyles.AllowCurrencySymbol, CultureInfo.CurrentCulture, out number);
    }

    //The cells in a direction from a cell, nearest first.
    private static IEnumerable<TableCell> Neighbours(TableCell cell, string direction)
    {
        TableRow row = (TableRow)cell.Parent;
        Table table = (Table)row.Parent;
        if (direction == "LEFT")
        {
            return cell.ElementsBefore().OfType<TableCell>().Reverse();
        }
        if (direction == "RIGHT")
        {
            return cell.ElementsAfter().OfType<TableCell>();
        }
        List<TableRow> rows = table.Elements<TableRow>().ToList();
        int start = TableCellWrapper.GetGridStart(cell);
        IEnumerable<TableRow> others = direction == "ABOVE" ? rows.Take(rows.IndexOf(row)).Reverse() : rows.Skip(rows.IndexOf(row) + 1);
        return others.Select(r => TableCellWrapper.CellAtGridColumn(r, start)).TakeWhile(c => c != null);
    }

    //The numbers in a direction, up to the first cell that is empty or has no number in it, as Word takes them.
    private static List<double> Numbers(TableCell cell, string direction)
    {
        List<double> numbers = new List<double>();
        foreach (TableCell neighbour in Neighbours(cell, direction))
        {
            if (!TryNumber(neighbour, out double number))
            {
                break;
            }
            numbers.Add(number);
        }
        return numbers;
    }

    /// <summary>
    /// The code of the formula that Word's AutoSum puts in a cell: the sum of the numbers above when there is one directly above, otherwise
    /// of those to the left.
    /// </summary>
    internal static string AutoSumCode(TableCell cell)
    {
        if (Numbers(cell, "ABOVE").Count > 0)
        {
            return "=SUM(ABOVE)";
        }
        return Numbers(cell, "LEFT").Count > 0 ? "=SUM(LEFT)" : "=SUM()";
    }

    /// <summary>
    /// The result of a formula field in a table cell, or null when it is not one that can be worked out here.
    /// </summary>
    internal static string Compute(TableCell cell, string expression, string picture)
    {
        if (cell == null)
        {
            return null;
        }
        Match match = Regex.Match(expression.Trim(), @"^=?\s*([A-Za-z]+)\(\s*([A-Za-z]*)\s*\)$");
        if (!match.Success)
        {
            return null;
        }
        string function = match.Groups[1].Value.ToUpperInvariant();
        string direction = match.Groups[2].Value.ToUpperInvariant();
        if (direction.Length == 0)
        {
            return function == "SUM" ? UnexpectedEnd : null;
        }
        if (direction != "ABOVE" && direction != "BELOW" && direction != "LEFT" && direction != "RIGHT")
        {
            return null;
        }

        List<double> numbers = Numbers(cell, direction);
        double value;
        switch (function)
        {
            case "SUM":
                value = numbers.Sum();
                break;
            case "COUNT":
                value = numbers.Count;
                break;
            case "AVERAGE":
                if (numbers.Count == 0)
                {
                    return "!Zero Divide";
                }
                value = numbers.Average();
                break;
            case "MIN":
                value = numbers.Count == 0 ? 0 : numbers.Min();
                break;
            case "MAX":
                value = numbers.Count == 0 ? 0 : numbers.Max();
                break;
            case "PRODUCT":
                value = numbers.Count == 0 ? 0 : numbers.Aggregate(1.0, (a, b) => a * b);
                break;
            default:
                return null;
        }

        CultureInfo culture = CultureInfo.CurrentCulture;
        try
        {
            return picture == null ? value.ToString("0.##########", culture) : value.ToString(picture, culture);
        }
        catch (FormatException)
        {
            return null;
        }
    }
}

public partial class TableCellWrapper
{
    /// <summary>
    /// Puts a formula field at the start of the cell, as <c>Cell.Formula</c> in Word. Without a formula it is the sum of the numbers above the
    /// cell, or when there are none directly above, of those to the left. The result is worked out here for the functions
    /// <c>SUM</c>, <c>AVERAGE</c>, <c>COUNT</c>, <c>MIN</c>, <c>MAX</c> and <c>PRODUCT</c> of <c>ABOVE</c>, <c>BELOW</c>, <c>LEFT</c> or
    /// <c>RIGHT</c>; other formulas are left for Word to work out when the document is opened.
    /// </summary>
    /// <param name="formula">The formula, with or without the equals sign, or null for the sum of the numbers above or to the left.</param>
    /// <param name="numFormat">A numeric picture such as <c>0.00</c> for the result, or null.</param>
    public void Formula(string formula = null, string numFormat = null)
    {
        string code = string.IsNullOrWhiteSpace(formula) ? TableFormulas.AutoSumCode(_tableCell) : formula.Trim()[0] == '=' ? formula.Trim() : "=" + formula.Trim();
        if (!string.IsNullOrEmpty(numFormat))
        {
            code += " \\# \"" + numFormat + "\"";
        }
        RangeWrapper start = Range;
        start.Collapse(CollapseDirection.Start);
        start.Fields.Add(start, code);
    }

    /// <summary>Puts the sum of the numbers above the cell, or to its left, in the cell; the same as <see cref="Formula"/> without arguments.</summary>
    public void AutoSum()
    {
        Formula();
    }
}
