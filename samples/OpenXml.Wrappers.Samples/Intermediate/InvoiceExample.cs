using System.Globalization;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Intermediate;

/// <summary>
/// An invoice: tables for the addresses and the line items, a repeating header row, right-aligned amounts, merged cells for the totals,
/// shading and borders, with the totals worked out from the data.
/// </summary>
internal static class InvoiceExample
{
    private sealed record Item(string Description, int Quantity, decimal UnitPrice);

    public static void Create(string path)
    {
        Item[] items =
        {
            new("Atlas planner, A5, 2026", 40, 12.50m),
            new("Refill pack, 12 sheets", 120, 3.20m),
            new("Gift wrapping", 15, 1.75m),
            new("Delivery, tracked", 1, 18.00m),
        };
        const decimal TaxRate = 0.20m;

        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Invoice INV-2026-0314"))
        {
            DocumentBuilder builder = new DocumentBuilder(document);
            builder.Paragraph("Invoice", "Title");

            //Two columns without borders: who it is from, who it is to.
            TableWrapper parties = builder.Table(1, 2);
            parties.Cell(0, 0).Range.Text = "Atlas Stationery Ltd\n1 Paper Lane\nLeeds LS1 4AB\nVAT 123 4567 89";
            parties.Cell(0, 1).Range.Text = "Bill to\nEastside Offices Ltd\n77 Canal Street\nManchester M1 3GG";
            parties.Borders.SetOutside(BorderLineStyle.None, 0.5, "auto");
            parties.Borders.SetInside(BorderLineStyle.None, 0.5, "auto");

            builder.Paragraph("Invoice number INV-2026-0314   |   Date 14 March 2026   |   Payment due 13 April 2026").Format.SpaceBefore = 12;

            //Header row, one row for each item, and three rows for the totals.
            TableWrapper table = builder.Table(1 + items.Length + 3, 4);
            string[] titles = { "Description", "Quantity", "Unit price", "Amount" };
            for (int column = 0; column < titles.Length; column++)
            {
                table.Cell(0, column).Range.Text = titles[column];
                table.Cell(0, column).Shading.BackgroundPatternColor = "1F3864";
            }
            table.Rows[0].Range.Font.Bold = true;
            table.Rows[0].Range.Font.Color = "FFFFFF";
            table.Rows[0].HeadingFormat = true;

            decimal subtotal = 0;
            for (int i = 0; i < items.Length; i++)
            {
                decimal amount = items[i].Quantity * items[i].UnitPrice;
                subtotal += amount;
                table.Cell(i + 1, 0).Range.Text = items[i].Description;
                table.Cell(i + 1, 1).Range.Text = items[i].Quantity.ToString(CultureInfo.InvariantCulture);
                table.Cell(i + 1, 2).Range.Text = items[i].UnitPrice.ToString("N2", CultureInfo.InvariantCulture);
                table.Cell(i + 1, 3).Range.Text = amount.ToString("N2", CultureInfo.InvariantCulture);
                table.Rows[i + 1].AllowBreakAcrossPages = false;
            }

            //The totals: the label goes in the first cell, and the first three cells of the row are merged to make room for it.
            decimal tax = Math.Round(subtotal * TaxRate, 2);
            (string Label, decimal Value)[] totals = { ("Subtotal", subtotal), ("VAT at 20%", tax), ("Total due (GBP)", subtotal + tax) };
            for (int i = 0; i < totals.Length; i++)
            {
                int row = 1 + items.Length + i;
                table.Cell(row, 0).Range.Text = totals[i].Label;
                table.Cell(row, 3).Range.Text = totals[i].Value.ToString("N2", CultureInfo.InvariantCulture);
                table.Cell(row, 0).Merge(table.Cell(row, 2));
                table.Rows[row].Range.Font.Bold = i == totals.Length - 1;
            }
            table.Rows[table.Rows.Count - 1].Cells[0].Shading.BackgroundPatternColor = "D9E2F3";
            table.Rows[table.Rows.Count - 1].Cells[1].Shading.BackgroundPatternColor = "D9E2F3";

            //Numbers are right-aligned, and so are the headings above them. After the merge the total rows have two cells: the label and
            //the amount.
            for (int column = 1; column < 4; column++)
            {
                for (int row = 0; row <= items.Length; row++)
                {
                    table.Cell(row, column).Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                }
            }
            for (int row = 1 + items.Length; row < table.Rows.Count; row++)
            {
                table.Rows[row].Cells[1].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            }

            table.Columns[0].Width = 4200;
            table.Columns[1].Width = 1300;
            table.Columns[2].Width = 1500;
            table.Columns[3].Width = 1600;
            table.Borders.SetOutside(BorderLineStyle.SingleLine, 1, "1F3864");
            table.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");

            builder.Paragraph("Payment by bank transfer to Atlas Stationery Ltd, sort code 12-34-56, account 87654321, quoting the invoice number.").Format.SpaceBefore = 12;
        }
    }
}
