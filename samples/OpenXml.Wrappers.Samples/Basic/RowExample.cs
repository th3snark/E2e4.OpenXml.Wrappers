using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Basic;

/// <summary>
/// Rows: height, repeating header rows, keeping a row on one page, inserting a row and row borders.
/// </summary>
internal static class RowExample
{
    public static void Create(string path)
    {
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(path, true, false))
        {
            document.Paragraphs[0].Range.Text = "Task list";
            ParagraphWrapper after = document.Content.InsertParagraphAfter();
            TableWrapper table = document.Tables.Add(after.Range, 3, 2);
            table.Cell(0, 0).Range.Text = "Task";
            table.Cell(0, 1).Range.Text = "Owner";
            table.Cell(1, 0).Range.Text = "Write the specification";
            table.Cell(1, 1).Range.Text = "Ana";
            table.Cell(2, 0).Range.Text = "Review the specification";
            table.Cell(2, 1).Range.Text = "Ben";

            //The first row repeats at the top of every page the table runs onto.
            TableRowWrapper header = table.Rows[0];
            header.HeadingFormat = true;
            header.Range.Font.Bold = true;
            header.Height = 24;
            header.HeightRule = RowHeightRule.AtLeast;
            header.Cells[0].BackgroundColor = "D9E2F3";
            header.Cells[1].BackgroundColor = "D9E2F3";

            //A row that must not be split across two pages, and one with an exact height.
            table.Rows[1].AllowBreakAcrossPages = false;
            table.Rows[2].Height = 30;
            table.Rows[2].HeightRule = RowHeightRule.Exactly;

            //Add(beforeRow) inserts an empty row with the same cells before a row. Without an argument it adds a row at the end.
            TableRowWrapper inserted = table.Rows.Add(table.Rows[2]);
            inserted.Cells[0].Range.Text = "Fix the problems";
            inserted.Cells[1].Range.Text = "Cy";
            TableRowWrapper appended = table.Rows.Add();
            appended.Cells[0].Range.Text = "Publish";

            //Row borders are written to the cells of the row.
            header.Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.DoubleLine;
            header.Borders[BorderPosition.Bottom].Color = "1F3864";
            table.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");

            //Next, Previous, IsFirst and IsLast move around the rows.
            TableRowWrapper row = table.Rows[0];
            while (!row.IsLast)
            {
                row = row.Next();
            }
            row.Cells[1].Range.Text = "Ana";
        }
    }
}
