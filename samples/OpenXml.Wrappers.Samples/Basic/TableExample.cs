using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Tables and their columns: adding a table, filling it, sizing and aligning it, and giving it borders.
/// </summary>
internal static class TableExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            document.Paragraphs[0].Range.Text = "Sales by region (thousands)";
            document.Paragraphs[0].Range.Font.Bold = true;

            //Tables.Add replaces the range it is given (as in Word) and puts the table before its paragraph, so give it an empty paragraph, which follows the table.
            ParagraphWrapper after = document.Content.InsertParagraphAfter();
            TableWrapper table = document.Tables.Add(after.Range, 4, 3);

            int[][] sales = { new[] { 120, 134 }, new[] { 98, 105 }, new[] { 143, 151 } };
            string[] regions = { "North", "South", "West" };
            table.Cell(0, 0).Range.Text = "Region";
            table.Cell(0, 1).Range.Text = "2024";
            table.Cell(0, 2).Range.Text = "2025";
            for (int row = 0; row < regions.Length; row++)
            {
                //Rows and columns are 0-based, unlike in Word.
                table.Cell(row + 1, 0).Range.Text = regions[row];
                table.Cell(row + 1, 1).Range.Text = sales[row][0].ToString(CultureInfo.InvariantCulture);
                table.Cell(row + 1, 2).Range.Text = sales[row][1].ToString(CultureInfo.InvariantCulture);
            }

            //A column is added at the right, and a row at the bottom.
            TableColumnWrapper change = table.Columns.Add();
            table.Cell(0, 3).Range.Text = "Change";
            for (int row = 0; row < regions.Length; row++)
            {
                int percent = (sales[row][1] - sales[row][0]) * 100 / sales[row][0];
                table.Cell(row + 1, 3).Range.Text = string.Format(CultureInfo.InvariantCulture, "{0}%", percent);
            }

            table.Rows.Add();
            int total2024 = sales.Sum(s => s[0]);
            int total2025 = sales.Sum(s => s[1]);
            table.Cell(4, 0).Range.Text = "Total";
            table.Cell(4, 1).Range.Text = total2024.ToString(CultureInfo.InvariantCulture);
            table.Cell(4, 2).Range.Text = total2025.ToString(CultureInfo.InvariantCulture);
            table.Cell(4, 3).Range.Text = string.Format(CultureInfo.InvariantCulture, "{0}%", (total2025 - total2024) * 100 / total2024);
            table.Rows[0].Range.Font.Bold = true;
            table.Rows[4].Range.Font.Bold = true;

            //Column widths are in twentieths of a point; the table alignment uses Word's names.
            table.Columns[0].Width = 2400;
            change.Width = 1400;
            table.Alignment = RowAlignment.Center;

            //Borders work as in Word: set the outside and the inside lines separately.
            table.Borders.SetOutside(BorderLineStyle.SingleLine, 1.5, "1F3864");
            table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "808080");
        }
    }
}
