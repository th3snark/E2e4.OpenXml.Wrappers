using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Cells: merging and splitting, shading, alignment and borders.
/// </summary>
internal static class CellExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            document.Paragraphs[0].Range.Text = "Merged, split and shaded cells";
            ParagraphWrapper after = document.Range.InsertParagraphAfter();
            TableWrapper table = document.Tables.Add(after.Range, 4, 4);
            for (int row = 0; row < 4; row++)
            {
                for (int column = 0; column < 4; column++)
                {
                    table.Cell(row, column).Range.Text = string.Format(CultureInfo.InvariantCulture, "R{0}C{1}", row + 1, column + 1);
                }
            }

            //Merge joins the rectangle between two cells; the text of the other cells would move into the first one.
            //The cells that are merged into the first one are emptied first, as their text would be moved into it.
            for (int column = 1; column < 4; column++)
            {
                table.Cell(0, column).Range.Text = string.Empty;
            }
            table.Cell(0, 0).Merge(table.Cell(0, 3));
            table.Cell(0, 0).Range.Text = "Title across four columns";
            table.Cell(0, 0).Range.Font.Bold = true;
            table.Cell(0, 0).Range.Font.Color = "FFFFFF";
            table.Cell(0, 0).BackgroundColor = "1F3864";

            table.Cell(2, 0).Range.Text = string.Empty;
            table.Cell(3, 0).Range.Text = string.Empty;
            table.Cell(1, 0).Merge(table.Cell(3, 0));
            table.Cell(1, 0).Range.Text = "Three rows tall";
            table.Cell(1, 0).VerticalAlignment = CellVerticalAlignment.Center;

            //Split divides a cell into rows and columns.
            table.Cell(1, 1).Split(1, 2);

            //Shading, alignment and borders of single cells.
            table.Cell(2, 1).BackgroundColor = "FFF2CC";
            table.Cell(2, 2).BackgroundColor = "E2EFDA";
            table.Cell(3, 3).VerticalAlignment = CellVerticalAlignment.Bottom;
            table.Cell(3, 3).Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.DoubleLine;
            table.Cell(3, 3).Borders[BorderPosition.Bottom].Color = "C00000";

            //Next and Previous move along the row and wrap to the next row.
            table.Cell(2, 2).Next().Range.Font.Italic = true;

            table.Borders.SetOutside(BorderLineStyle.SingleLine, 1, "808080");
            table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");
        }
    }
}
