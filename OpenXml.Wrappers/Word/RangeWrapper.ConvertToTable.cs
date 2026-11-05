using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// Replaces the paragraphs of the range with a table, as <c>Range.ConvertToTable</c> in Word. Each paragraph is a row and the fields of
    /// it, divided by the separator, are its cells. With <see cref="TableFieldSeparator.Paragraphs"/> each paragraph is a cell and
    /// <paramref name="numColumns"/> of them make a row. The cells hold the text only: the formatting of the text is not kept.
    /// </summary>
    /// <param name="separator">What divides the fields of a paragraph.</param>
    /// <param name="numRows">The number of rows, when more than the text needs.</param>
    /// <param name="numColumns">The number of columns; the most fields in a paragraph when omitted, and 1 for paragraphs.</param>
    /// <param name="defaultTableBehavior">As for <see cref="TableWrapperCollection.Add"/>.</param>
    /// <param name="autoFitBehavior">As for <see cref="TableWrapperCollection.Add"/>.</param>
    /// <exception cref="InvalidOperationException">The range has no text.</exception>
    public TableWrapper ConvertToTable(TableFieldSeparator separator = TableFieldSeparator.Paragraphs, int? numRows = null, int? numColumns = null, DefaultTableBehavior defaultTableBehavior = DefaultTableBehavior.Word8TableBehavior, AutoFitBehavior autoFitBehavior = AutoFitBehavior.Fixed)
    {
        if (!EnumCompat.IsDefined(separator))
        {
            throw new ArgumentOutOfRangeException(nameof(separator));
        }
        string text = separator == TableFieldSeparator.Commas ? ", " : separator == TableFieldSeparator.Default ? "-" : null;
        return ConvertToTable(separator == TableFieldSeparator.Paragraphs, separator == TableFieldSeparator.Tabs ? "\t" : text, numRows, numColumns, defaultTableBehavior, autoFitBehavior);
    }

    /// <summary>
    /// Replaces the paragraphs of the range with a table whose cells are the fields of each paragraph divided by <paramref name="separator"/>,
    /// as <c>Range.ConvertToTable</c> in Word.
    /// </summary>
    /// <exception cref="ArgumentException">The separator is empty.</exception>
    /// <exception cref="InvalidOperationException">The range has no text.</exception>
    public TableWrapper ConvertToTable(string separator, int? numRows = null, int? numColumns = null, DefaultTableBehavior defaultTableBehavior = DefaultTableBehavior.Word8TableBehavior, AutoFitBehavior autoFitBehavior = AutoFitBehavior.Fixed)
    {
        if (string.IsNullOrEmpty(separator))
        {
            throw new ArgumentException("The separator cannot be empty.", nameof(separator));
        }
        return ConvertToTable(false, separator, numRows, numColumns, defaultTableBehavior, autoFitBehavior);
    }

    private TableWrapper ConvertToTable(bool paragraphs, string separator, int? numRows, int? numColumns, DefaultTableBehavior defaultTableBehavior, AutoFitBehavior autoFitBehavior)
    {
        if (numRows.HasValue)
        {
            Guard.ThrowIfLessThan(numRows.Value, 1);
        }
        if (numColumns.HasValue)
        {
            Guard.ThrowIfLessThan(numColumns.Value, 1);
        }

        List<string> lines = Text.Split('\r').ToList();
        if (lines.Count > 0 && lines[lines.Count - 1].Length == 0)
        {
            lines.RemoveAt(lines.Count - 1);
        }
        if (lines.All(l => l.Length == 0))
        {
            throw new InvalidOperationException("The range has no text to convert.");
        }

        List<List<string>> fields;
        if (paragraphs)
        {
            int columns = numColumns ?? 1;
            fields = new List<List<string>>();
            for (int index = 0; index < lines.Count; index += columns)
            {
                fields.Add(lines.Skip(index).Take(columns).ToList());
            }
        }
        else
        {
            fields = lines.Select(l => l.Split(new[] { separator }, StringSplitOptions.None).ToList()).ToList();
        }
        int width = Math.Max(numColumns ?? 0, fields.Max(f => f.Count));
        int height = Math.Max(numRows ?? 0, fields.Count);

        TableWrapper table = Tables.Add(this, height, width, defaultTableBehavior, autoFitBehavior);
        for (int row = 0; row < fields.Count; row++)
        {
            for (int column = 0; column < fields[row].Count; column++)
            {
                if (fields[row][column].Length > 0)
                {
                    table.Cell(row, column).Range.Text = fields[row][column];
                }
            }
        }
        return table;
    }
}
