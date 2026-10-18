using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.Samples.Advanced;

/// <summary>
/// A product catalogue made from a list of products: one section for each category with its own page header, a block for each product
/// with a table of specifications nested in it, an index table at the end, bookmarks for every product, and a page footer. It shows
/// loops over data, nested tables, sections with different headers, and bookmarks together.
/// </summary>
internal static class CatalogueExample
{
    private sealed record Product(string Sku, string Name, string Category, decimal Price, string Description, (string Name, string Value)[] Specs);

    public static void Create(string path)
    {
        Product[] products =
        {
            new("TK-100", "Trail kettle", "Camping", 24.95m, "A 1 litre kettle that packs flat and boils in four minutes on a small stove.", new[] { ("Capacity", "1.0 l"), ("Weight", "180 g"), ("Material", "Anodised aluminium") }),
            new("TK-200", "Trail stove", "Camping", 39.00m, "A folding single-burner stove for gas canisters, with a built-in igniter.", new[] { ("Output", "2.8 kW"), ("Weight", "95 g"), ("Fuel", "Screw-fit gas canister") }),
            new("TK-300", "Dome tent, two person", "Camping", 189.00m, "A freestanding two-pole tent with a full-length porch and taped seams.", new[] { ("Weight", "2.1 kg"), ("Pitch", "Inner first"), ("Waterproofing", "3,000 mm") }),
            new("OF-100", "Standing desk, 140 cm", "Office", 349.00m, "A height-adjustable desk with two motors and four memory settings.", new[] { ("Width", "140 cm"), ("Height range", "62 to 128 cm"), ("Load", "100 kg") }),
            new("OF-200", "Task chair", "Office", 215.00m, "A chair with adjustable lumbar support and arms, and a mesh back.", new[] { ("Seat height", "42 to 52 cm"), ("Weight", "14 kg"), ("Warranty", "5 years") }),
            new("OF-300", "Monitor arm", "Office", 59.00m, "A gas-spring arm for one screen up to 32 inches.", new[] { ("Load", "2 to 9 kg"), ("Mount", "Clamp or grommet"), ("Rotation", "360 degrees") }),
            new("KT-100", "Chef's knife, 20 cm", "Kitchen", 64.00m, "A forged steel knife with a pakkawood handle, balanced at the bolster.", new[] { ("Blade", "20 cm"), ("Steel", "X50CrMoV15"), ("Hardness", "58 HRC") }),
            new("KT-200", "Cast iron pan, 26 cm", "Kitchen", 45.00m, "A pre-seasoned pan for the hob, the oven and the campfire.", new[] { ("Diameter", "26 cm"), ("Weight", "2.4 kg"), ("Finish", "Pre-seasoned") }),
        };
        string[] categories = products.Select(p => p.Category).Distinct().ToList().ToArray();

        using (WordprocessingDocumentWrapper document = SampleDocument.Create(path, "Spring catalogue 2026"))
        {
            DocumentBuilder builder = new DocumentBuilder(document);

            //Cover page.
            builder.Paragraph("Spring catalogue 2026", "Title").Format.SpaceBefore = 150;
            builder.Paragraph(string.Format(CultureInfo.InvariantCulture, "{0} products in {1} categories", products.Length, categories.Length)).Range.Font.Size = 16;
            ParagraphWrapper cover = builder.Paragraph("Prices include VAT. Valid until 30 June 2026.");
            cover.InsertBreak(BreakValues.Page);

            //One section for each category. The page header of each section is set after all of the content is in place.
            for (int c = 0; c < categories.Length; c++)
            {
                builder.Paragraph(categories[c], "Heading 1");
                foreach (Product product in products.Where(p => p.Category == categories[c]))
                {
                    AddProduct(document, builder, product);
                }

                //Each category, and then the index, is a section of its own.
                builder.Paragraph(string.Empty).InsertSectionBreak("nextPage");
            }

            //Index of all products.
            builder.Paragraph("Index of products", "Heading 1");
            TableWrapper index = builder.Table(1 + products.Length, 3);
            index.Cell(0, 0).Range.Text = "Code";
            index.Cell(0, 1).Range.Text = "Product";
            index.Cell(0, 2).Range.Text = "Price (GBP)";
            for (int i = 0; i < products.Length; i++)
            {
                index.Cell(i + 1, 0).Range.Text = products[i].Sku;
                index.Cell(i + 1, 1).Range.Text = products[i].Name;
                index.Cell(i + 1, 2).Range.Text = products[i].Price.ToString("N2", CultureInfo.InvariantCulture);
                index.Rows[i + 1].Cells[i % 2 == 0 ? 0 : 1].BackgroundColor = "F2F2F2";
            }
            TableLook.Grid(index, true);
            TableLook.AlignRight(index, 2);

            //Headers: the first section creates its own; each later one is first unlinked from the one before and then given its text.
            for (int s = 0; s < document.Sections.Count; s++)
            {
                HeaderFooterWrapper header = document.Sections[s].Headers[HeaderFooterIndex.Primary];
                if (s > 0)
                {
                    header.LinkToPrevious = false;
                }
                string title = s < categories.Length ? categories[s] : "Index";
                header.Range.Text = "Spring catalogue 2026 - " + title;
                header.Range.Style = "Header";
                header.Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            }

            //One footer, linked through all of the sections.
            RangeWrapper footer = document.Sections[0].Footers[HeaderFooterIndex.Primary].Range;
            footer.Text = "Page ";
            footer.Style = "Footer";
            footer.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            document.Fields.Add(footer, FieldType.Page).Result = "1";

            document.BuiltInDocumentProperties.Subject = "Product catalogue";
            document.BuiltInDocumentProperties.Keywords = string.Join("; ", categories);
            document.CustomDocumentProperties.Add("Products", DocumentPropertyType.Number, products.Length);
        }
    }

    //A block for one product: a two-column table with the name and price on the left and the description and a table of specifications
    //on the right. The block is bookmarked by the code of the product.
    private static void AddProduct(WordprocessingDocumentWrapper document, DocumentBuilder builder, Product product)
    {
        TableWrapper block = builder.Table(1, 2);
        block.Columns[0].Width = 2600;
        block.Columns[1].Width = 5900;

        TableCellWrapper left = block.Cell(0, 0);
        left.Range.Text = product.Name + "\n" + product.Price.ToString("C", CultureInfo.GetCultureInfo("en-GB")) + "\n" + product.Sku;
        left.BackgroundColor = "D9E2F3";
        left.Range.Font.Bold = true;
        left.VerticalAlignment = CellVerticalAlignment.Center;

        TableCellWrapper right = block.Cell(0, 1);
        right.Range.Text = product.Description;

        //A table inside a cell: Tables.Add replaces the range it is given (as in Word) and puts the table before its paragraph, so give it a new, empty paragraph after the description.
        ParagraphWrapper spot = right.Range.Paragraphs[0].Range.InsertParagraphAfter();
        TableWrapper specs = document.Tables.Add(spot.Range, product.Specs.Length, 2);
        for (int i = 0; i < product.Specs.Length; i++)
        {
            specs.Cell(i, 0).Range.Text = product.Specs[i].Name;
            specs.Cell(i, 1).Range.Text = product.Specs[i].Value;
            specs.Cell(i, 0).Range.Font.Italic = true;
        }
        specs.Borders.SetOutside(BorderLineStyle.None, 0.5, "auto");
        specs.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "D9D9D9");
        specs.Columns[0].Width = 1700;
        specs.Columns[1].Width = 3900;

        block.Borders.SetOutside(BorderLineStyle.SingleLine, 0.75, "808080");
        block.Borders.SetInside(BorderLineStyle.SingleLine, 0.5, "BFBFBF");
        block.Rows[0].AllowBreakAcrossPages = false;

        //Bookmark names cannot contain a hyphen, so the code TK-100 becomes sku_TK_100.
        document.Bookmarks.Add("sku_" + product.Sku.Replace('-', '_'), block.Range);
    }
}
