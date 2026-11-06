using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Compares the markup the library writes with what Word itself writes for the same operation. Each scenario folder under
/// <c>TestFiles/WordReference</c> holds the normalised parts of a document Word saved after running the scenario of the same name in
/// <c>tools/WordReference/scenarios.ps1</c> on a new document; each case here does the same through the library on the output of
/// <see cref="WordprocessingDocumentWrapper.Create(Stream, bool, bool)"/>, normalises it with <see cref="MarkupNormalizer"/> and compares the
/// parts (or the elements of a part) that the scenario is about.
/// </summary>
/// <remarks>
/// Regenerate the references with <c>tools/WordReference/New-WordReferences.ps1</c>.
/// </remarks>
public class MarkupConformanceTest
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string ReferenceFolder = Path.Combine(AppContext.BaseDirectory, "TestFiles", "WordReference");

    //A 3x3 table whose cells hold c1 to c9.
    private static TableWrapper NumberedTable(WordprocessingDocumentWrapper d)
    {
        TableWrapper table = d.Tables.Add(d.Range(0, 0), 3, 3);
        int number = 1;
        foreach (TableCellWrapper cell in table.Range.Cells)
        {
            cell.Range.Text = "c" + number.ToString(System.Globalization.CultureInfo.InvariantCulture);
            number++;
        }
        return table;
    }

    private static readonly Scenario[] Scenarios = new[]
    {
        new Scenario("blank-document", d => { }, AllParts()),

        new Scenario("insert-after-tab-and-line-break",
            d => d.Content.InsertAfter("Before\tafter the tab\vafter the break"),
            Body()),

        new Scenario("bold-on",
            d =>
            {
                d.Content.InsertAfter("Bold text");
                d.Range(0, 4).Font.Bold = true;
            },
            Body()),

        new Scenario("bold-on-then-off",
            d =>
            {
                d.Content.InsertAfter("Bold text");
                d.Range(0, 4).Font.Bold = true;
                d.Range(0, 4).Font.Bold = false;
            },
            Body()),

        new Scenario("font-name",
            d =>
            {
                d.Content.InsertAfter("Arial text");
                d.Range(0, 5).Font.Name = "Arial";
            },
            Body()),
        new Scenario("font-complex-script",
            d =>
            {
                d.Content.InsertAfter("Complex text here");
                FontWrapper font = d.Range(0, 7).Font;
                font.BoldBi = true;
                font.ItalicBi = true;
                font.SizeBi = 14;
                font.NameBi = "Tahoma";
                font.NameFarEast = "MS Gothic";
                font.NameOther = "Verdana";
            },
            Body()),

        new Scenario("font-emphasis-grid",
            d =>
            {
                d.Content.InsertAfter("Mark text here");
                d.Range(0, 4).Font.EmphasisMark = EmphasisMark.OverSolid;
                d.Range(5, 9).Font.EmphasisMark = EmphasisMark.UnderSolid;
                d.Range(10, 14).Font.DisableCharacterSpaceGrid = true;
            },
            Body()),

        new Scenario("font-name-after-complex",
            d =>
            {
                d.Content.InsertAfter("Name text here");
                d.Range(0, 4).Font.NameBi = "Tahoma";
                d.Range(0, 4).Font.Name = "Calibri";
            },
            Body()),

        new Scenario("font-grow-shrink",
            d =>
            {
                d.Content.InsertAfter("Grow shrink text");
                d.Range(0, 4).Font.Grow();
                d.Range(5, 11).Font.Shrink();
                d.Range(12, 16).Font.Grow();
                d.Range(12, 16).Font.Grow();
            },
            Body()),

        new Scenario("paragraph-spacing-shortcuts",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree\rFour");
                d.Paragraphs[0].OpenUp();
                d.Paragraphs[1].Space15();
                d.Paragraphs[2].Indent();
                d.Paragraphs[3].Format.TabHangingIndent(1);
            },
            Body()),

        new Scenario("paragraph-east-asian-switches",
            d =>
            {
                d.Content.InsertAfter("One\rTwo");
                ParagraphFormatWrapper format = d.Paragraphs[0].Format;
                format.AddSpaceBetweenFarEastAndAlpha = false;
                format.BaseLineAlignment = BaselineAlignment.Center;
                format.HangingPunctuation = false;
                format.FarEastLineBreakControl = false;
                format.HalfWidthPunctuationOnTopOfLine = true;
                format.DisableLineHeightGrid = true;
                format.TextboxTightWrap = TextboxTightWrap.AllLines;
                format.WordWrap = false;
                format.AutoAdjustRightIndent = false;
                format.AddSpaceBetweenFarEastAndDigit = false;
            },
            Body()),

        new Scenario("paragraph-right-to-left",
            d =>
            {
                d.Content.InsertAfter("One\rTwo");
                d.Paragraphs[0].Format.ReadingOrder = ReadingOrder.RightToLeft;
            },
            Body()),
        new Scenario("paragraph-borders-shading",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree\rFour text");
                d.Paragraphs[0].Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
                BordersWrapper outside = d.Paragraphs[1].Borders;
                outside.OutsideLineStyle = BorderLineStyle.SingleLine;
                outside.OutsideLineWidth = 1.5;
                outside.OutsideColor = ColorHelper.RGB(255, 0, 0);
                d.Paragraphs[2].Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
                d.Paragraphs[2].Borders[BorderPosition.Left].LineStyle = BorderLineStyle.SingleLine;
                d.Paragraphs[3].Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
                d.Paragraphs[3].Shading.Texture = ShadingTexture.Percent10;
            },
            Body()),

        new Scenario("font-border-shading",
            d =>
            {
                d.Content.InsertAfter("One two three four");
                d.Range(0, 3).Font.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
                d.Range(4, 7).Font.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);
                d.Range(8, 13).Font.Borders.OutsideLineStyle = BorderLineStyle.DoubleLine;
            },
            Body()),
        new Scenario("paragraph-outline",
            d =>
            {
                d.Content.InsertAfter("A\rB\rC\rD\rE");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[1].Style = "Heading 2";
                d.Paragraphs[4].Style = "Heading 3";
                d.Paragraphs[3].Format.OutlineLevel = ParagraphOutlineLevel.Level3;
                d.Paragraphs[0].OutlineDemote();
                d.Paragraphs[1].OutlinePromote();
                d.Paragraphs[2].OutlineDemote();
                d.Paragraphs[3].OutlineDemote();
                d.Paragraphs[4].OutlineDemoteToBody();
            },
            Body()),

        new Scenario("paragraphs-spacing-add",
            d =>
            {
                d.Content.InsertAfter("A\rB\rC");
                d.Paragraphs.IncreaseSpacing();
                d.Paragraphs.IncreaseSpacing();
                d.Paragraphs.DecreaseSpacing();
                d.Paragraphs.Indent();
                d.Paragraphs.Add(d.Paragraphs[1].Range);
            },
            Body()),
        new Scenario("cells-collection-format",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                TableCellWrapperCollection cells = d.Tables[0].Range.Cells;
                cells.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
                cells.VerticalAlignment = CellVerticalAlignment.Center;
                cells.Height = 30;
                cells.PreferredWidthType = PreferredWidthType.Points;
                cells.PreferredWidth = 80;
                d.Tables[0].Rows[1].Cells.SetHeight(40, RowHeightRule.Exactly);
            },
            Body()),

        new Scenario("columns-collection-format",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                TableColumnWrapperCollection columns = d.Tables[0].Columns;
                columns.Width = 100;
                columns.PreferredWidthType = PreferredWidthType.Points;
                columns.PreferredWidth = 90;
                columns.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
            },
            Body()),

        new Scenario("rows-set-height",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Rows.SetHeight(40, RowHeightRule.Exactly);
                d.Tables[0].Cell(2, 0).SetHeight(25, RowHeightRule.AtLeast);
            },
            Body()),
        new Scenario("cell-width-ragged",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Cell(1, 1).Width = 100;
            },
            Body()),
        new Scenario("cell-set-width-proportional",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Cell(1, 0).SetWidth(100, RulerStyle.Proportional);
            },
            Body()),
        new Scenario("cell-set-width-first-column",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Cell(1, 0).SetWidth(100, RulerStyle.FirstColumn);
            },
            Body()),
        new Scenario("cells-set-width-row",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Rows[1].Cells.SetWidth(100, RulerStyle.None);
            },
            Body()),
        new Scenario("cells-distribute-width",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Cell(1, 1).Width = 100;
                d.Tables[0].Range.Cells.DistributeWidth();
            },
            Body()),
        new Scenario("columns-set-width",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Columns[0].SetWidth(100, RulerStyle.FirstColumn);
                d.Tables[0].Columns[1].SetWidth(120, RulerStyle.Proportional);
            },
            Body()),
        new Scenario("columns-distribute-width",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Columns[0].Width = 100;
                d.Tables[0].Columns.DistributeWidth();
            },
            Body()),
        new Scenario("cells-borders-table",
            d =>
            {
                d.Tables.Add(d.Range(0, 0), 3, 3);
                d.Tables[0].Range.Cells.Borders.Enable = true;
            },
            Body()),
        new Scenario("cells-borders-row",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 3, 3);
                BordersWrapper borders = t.Rows[1].Cells.Borders;
                borders.OutsideLineStyle = BorderLineStyle.SingleLine;
                borders.InsideLineStyle = BorderLineStyle.SingleLine;
            },
            Body()),

        new Scenario("rows-borders",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 3, 3);
                t.Rows.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;
                t.Rows.Borders.InsideLineStyle = BorderLineStyle.SingleLine;
            },
            Body()),
        new Scenario("columns-borders",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 3, 3);
                t.Columns.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;
                t.Columns.Borders.InsideLineStyle = BorderLineStyle.SingleLine;
            },
            Body()),
        new Scenario("cells-merge-row",
            d =>
            {
                TableWrapper t = NumberedTable(d);
                t.Rows[1].Cells.Merge();
            },
            Body()),
        new Scenario("cells-merge-block", d => NumberedTable(d).Cell(0, 0).Merge(d.Tables[0].Cell(1, 1)), Body()),
        new Scenario("table-convert-to-text-tabs", d => NumberedTable(d).ConvertToText(TableFieldSeparator.Tabs), Body()),
        new Scenario("table-convert-to-text-paragraphs", d => NumberedTable(d).ConvertToText(TableFieldSeparator.Paragraphs), Body()),
        new Scenario("table-convert-to-text-commas", d => NumberedTable(d).ConvertToText(";"), Body()),
        new Scenario("table-convert-to-text-commas-enum", d => NumberedTable(d).ConvertToText(TableFieldSeparator.Commas), Body()),
        new Scenario("table-convert-to-text-default", d => NumberedTable(d).ConvertToText(TableFieldSeparator.Default), Body()),
        new Scenario("row-convert-to-text", d => NumberedTable(d).Rows[1].ConvertToText(TableFieldSeparator.Tabs), Body()),
        new Scenario("range-convert-to-table-tabs",
            d =>
            {
                d.Content.InsertAfter("a\tb\tc\rd\te\tf");
                d.Content.ConvertToTable(TableFieldSeparator.Tabs);
            },
            Body()),
        new Scenario("range-convert-to-table-paragraphs",
            d =>
            {
                d.Content.InsertAfter("a\rb\rc\rd");
                d.Content.ConvertToTable(TableFieldSeparator.Paragraphs, null, 2);
            },
            Body()),
        new Scenario("table-sort-text",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 4, 2);
                string[] values = { "b", "10", "c", "2", "a", "1", "d", "20" };
                int index = 0;
                foreach (TableCellWrapper cell in t.Range.Cells)
                {
                    cell.Range.Text = values[index++];
                }
                t.Sort(false, 0);
            },
            Body()),
        new Scenario("table-sort-numeric-header",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 4, 2);
                string[] values = { "h", "n", "c", "10", "b", "2", "a", "1" };
                int index = 0;
                foreach (TableCellWrapper cell in t.Range.Cells)
                {
                    cell.Range.Text = values[index++];
                }
                t.Sort(true, 1, SortFieldType.Numeric, SortOrder.Descending);
            },
            Body()),
        new Scenario("cell-formula-sum-above",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 3, 2);
                t.Cell(0, 0).Range.Text = "1";
                t.Cell(1, 0).Range.Text = "2";
                t.Cell(2, 0).Formula("=SUM(ABOVE)", "0.00");
            },
            Body()),
        new Scenario("cell-autosum",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 3, 2);
                t.Cell(0, 0).Range.Text = "1";
                t.Cell(1, 0).Range.Text = "2";
                t.Cell(2, 0).Formula();
            },
            Body()),
        new Scenario("row-left-indent", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].LeftIndent = 20, Body()),
        new Scenario("row-set-left-indent", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].SetLeftIndent(30, RulerStyle.None), Body()),
        new Scenario("row-set-left-indent-proportional", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].SetLeftIndent(30, RulerStyle.Proportional), Body()),
        new Scenario("row-set-left-indent-first-column", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].SetLeftIndent(30, RulerStyle.FirstColumn), Body()),
        new Scenario("row-set-left-indent-same-width", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].SetLeftIndent(30, RulerStyle.SameWidth), Body()),
        new Scenario("style-add-paragraph", d => d.Styles.Add("RefStyle", StyleType.Paragraph), RefStyles()),
        new Scenario("style-add-character", d => d.Styles.Add("RefStyle", StyleType.Character), RefStyles()),
        new Scenario("style-flags",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Paragraph);
                s.QuickStyle = true;
                s.Priority = 5;
                s.Locked = true;
                s.UnhideWhenUsed = true;
                s.AutomaticallyUpdate = true;
                s.NoSpaceBetweenParagraphsOfSameStyle = true;
                s.Visibility = false;
            },
            RefStyles()),
        new Scenario("style-based-next",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Paragraph);
                s.BaseStyle = "Heading 1";
                s.NextParagraphStyle = "Heading 2";
            },
            RefStyles()),
        new Scenario("style-language",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Paragraph);
                s.LanguageID = 1033;
                s.LanguageIDFarEast = 1041;
                s.NoProofing = true;
            },
            RefStyles()),
        new Scenario("cell-delete-up",
            d =>
            {
                TableWrapper t = NumberedTable(d);
                t.Cell(1, 1).Delete(ShiftCells.Up);
            },
            Body()),
        new Scenario("cells-delete-row-up",
            d =>
            {
                TableWrapper t = NumberedTable(d);
                t.Rows[1].Cells.Delete(ShiftCells.Up);
            },
            Body()),
        new Scenario("cells-add-end", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].Cells.Add(), Body()),
        new Scenario("cells-add-before",
            d =>
            {
                TableWrapper t = d.Tables.Add(d.Range(0, 0), 3, 3);
                t.Rows[1].Cells.Add(t.Cell(1, 1));
            },
            Body()),
        new Scenario("cells-split-columns", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].Cells.Split(1, 2), Body()),
        new Scenario("cells-split-rows", d => d.Tables.Add(d.Range(0, 0), 3, 3).Rows[1].Cells.Split(2, 1, false), Body()),
        new Scenario("table-3x3", d => d.Tables.Add(d.Range(0, 0), 3, 3), Body()),

        new Scenario("table-3x3-word9-behavior",
            d => d.Tables.Add(d.Range(0, 0), 3, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed),
            Body()),

        new Scenario("cell-merge",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 3, 3);
                table.Cell(0, 0).Merge(table.Cell(0, 1));
            },
            Body()),

        new Scenario("comment-with-reply",
            d =>
            {
                d.Content.InsertAfter("Commented text");
                CommentWrapper comment = d.Comments.Add(d.Range(0, 9), "A comment");
                CommentWrapper reply = comment.Replies.Add("A reply");

                AsReference(comment);
                AsReference(reply);
            },
            CommentParts()),

        new Scenario("bookmark",
            d =>
            {
                d.Content.InsertAfter("Bookmarked text");
                d.Bookmarks.Add("Mark", d.Range(0, 10));
            },
            Body()),

        new Scenario("date-field", d => d.Fields.Add(d.Range(0, 0), FieldType.Date), Body()),

        new Scenario("header-footer-text",
            d =>
            {
                SectionWrapper section = d.Sections[0];
                section.Headers[HeaderFooterIndex.Primary].Range.Text = "Header text";
                section.Footers[HeaderFooterIndex.Primary].Range.Text = "Footer text";
            },
            //Word also writes empty headers and footers for even pages and the first page, which say nothing, so the library leaves them out.
            HeaderFooter("header", "default"), HeaderFooter("footer", "default")),

        new Scenario("section-add",
            d =>
            {
                d.Content.InsertAfter("First section");
                d.Sections.Add();
            },
            Body()),

        new Scenario("page-borders", d => d.Sections[0].Borders.Enable = true, Elements("document.xml", "pgBorders")),

        new Scenario("protect-read-only",
            d => d.Protect(ProtectionType.AllowOnlyReading, "password"),
            Elements("settings.xml", "documentProtection")),

        new Scenario("page-border-options",
            d =>
            {
                BordersWrapper borders = d.Sections[0].Borders;
                borders.Enable = true;
                borders.AlwaysInFront = false;
                borders.DistanceFrom = BorderDistanceFrom.FromText;
                borders.EnableFirstPageInSection = false;
                borders.SurroundHeader = false;
                borders.SurroundFooter = false;
            },
            Elements("document.xml", "pgBorders"),
            Elements("settings.xml", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter")),

        new Scenario("page-setup-direction",
            d =>
            {
                d.PageSetup.TogglePortrait();
                d.PageSetup.SectionDirection = SectionDirection.Rtl;
            },
            Elements("document.xml", "pgSz", "bidi")),

        new Scenario("page-numbers-add-inside", d => d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers.Add(PageNumberAlignment.Inside), PageNumber()),

        new Scenario("page-numbers-add-outside", d => d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers.Add(PageNumberAlignment.Outside), PageNumber()),

        new Scenario("page-numbers-chapter",
            d =>
            {
                PageNumbersWrapper numbers = d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers;
                numbers.IncludeChapterNumber = true;
                numbers.HeadingLevelForChapter = 2;
                numbers.ChapterPageSeparator = SeparatorType.Colon;
                numbers.ShowFirstPageNumber = false;
            },
            Elements("document.xml", "pgNumType", "titlePg")),

        new Scenario("page-numbers-chapter-default",
            d =>
            {
                PageNumbersWrapper numbers = d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers;
                numbers.IncludeChapterNumber = true;
                numbers.ChapterPageSeparator = SeparatorType.Period;
                numbers.ChapterPageSeparator = SeparatorType.Hyphen;
            },
            Elements("document.xml", "pgNumType")),

        new Scenario("text-columns-add",
            d =>
            {
                TextColumnsWrapper columns = d.Sections[0].PageSetup.TextColumns;
                columns.Add(72, 10);
                columns.Add(100, 20);
                columns.Add(90);
            },
            Elements("document.xml", "cols")),


        new Scenario("text-columns-add-no-width",
            d =>
            {
                TextColumnsWrapper columns = d.Sections[0].PageSetup.TextColumns;
                columns.Add();
            },
            Elements("document.xml", "cols")),

        new Scenario("text-columns-items",
            d =>
            {
                TextColumnsWrapper columns = d.Sections[0].PageSetup.TextColumns;
                columns.SetCount(3);
                columns[0].Width = 100;
                columns[1].SpaceAfter = 10;
            },
            Elements("document.xml", "cols")),

        new Scenario("text-columns-flow",
            d =>
            {
                d.Sections[0].PageSetup.TextColumns.SetCount(2);
                d.Sections[0].PageSetup.TextColumns.FlowDirection = FlowDirection.Rtl;
            },
            Elements("document.xml", "cols", "bidi")),


        new Scenario("section-protected-for-forms",
            d =>
            {
                d.Content.InsertAfter("a");
                d.Sections.Add();
                d.Sections[0].ProtectedForForms = false;
                d.Sections[1].ProtectedForForms = false;
                d.Sections[1].ProtectedForForms = true;
            },
            Body()),

        new Scenario("fields-unlink",
            d =>
            {
                d.Content.InsertAfter("text");
                d.Fields.Add(d.Range(0, 0), FieldType.Empty, "SEQ a", false);
                d.Fields.Add(d.Range(0, 0), FieldType.Empty, "SEQ b", false);
                d.Fields.Unlink();
            },
            Body()),

        new Scenario("document-convert",
            d =>
            {
                d.SetCompatibilityMode(14);
                d.Convert();
            },
            Elements("settings.xml", "compat")),

        new Scenario("hyperlink-add",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Hyperlinks.Add(d.Range(0, 5), "https://example.com/");
            },
            HyperlinkParts()),

        new Scenario("hyperlink-add-place-only",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Bookmarks.Add("Mark", d.Range(6, 11));
                d.Hyperlinks.Add(d.Range(0, 5), string.Empty, "Mark", "Go to mark", "Display");
            },
            HyperlinkParts()),

        new Scenario("hyperlink-add-all-arguments",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Hyperlinks.Add(d.Range(0, 5), "https://example.org/page", "sec1", "Tip", "Shown", "_blank");
            },
            HyperlinkParts()),

        new Scenario("hyperlink-add-formatted",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Range(0, 5).Font.Bold = true;
                d.Range(3, 5).Font.Italic = true;
                d.Hyperlinks.Add(d.Range(0, 5), "https://example.com/");
            },
            HyperlinkParts()),

        new Scenario("hyperlink-add-collapsed-text",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Hyperlinks.Add(d.Range(5, 5), "https://example.com/", string.Empty, string.Empty, "Inserted");
            },
            HyperlinkParts()),

        new Scenario("hyperlink-add-collapsed",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Hyperlinks.Add(d.Range(5, 5), "https://example.com/");
            },
            HyperlinkParts()),

        new Scenario("hyperlink-set-text",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                HyperlinkWrapper link = d.Hyperlinks.Add(d.Range(0, 5), "https://example.com/");
                link.TextToDisplay = "Longer replacement";
            },
            HyperlinkParts()),

        new Scenario("hyperlink-edit",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                HyperlinkWrapper link = d.Hyperlinks.Add(d.Range(0, 5), "https://example.com/");
                link.Address = "https://changed.example/";
                link.ScreenTip = "A tip";
                link.SubAddress = "Mark2";
                link.Target = "_top";
            },
            HyperlinkParts()),

        new Scenario("hyperlink-email-subject",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                HyperlinkWrapper link = d.Hyperlinks.Add(d.Range(0, 5), "mailto:a@b.com");
                link.EmailSubject = "Hi there";
            },
            HyperlinkParts()),

        new Scenario("hyperlink-delete",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Hyperlinks.Add(d.Range(0, 5), "https://example.com/");
                d.Hyperlinks[0].Delete();
            },
            HyperlinkParts()),

        new Scenario("hyperlink-two-same-address",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Hyperlinks.Add(d.Range(12, 17), "https://example.com/");
                d.Hyperlinks.Add(d.Range(0, 5), "https://example.com/");
            },
            HyperlinkParts()),
        new Scenario("list-defaults",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Paragraphs[0].Range.ListFormat.ApplyBulletDefault();
                d.Paragraphs[1].Range.ListFormat.ApplyNumberDefault();
                d.Paragraphs[2].Range.ListFormat.ApplyOutlineNumberDefault();
            },
            Body()),

        new Scenario("list-apply-bullet",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyBulletDefault();
            },
            ListParts()),

        new Scenario("list-apply-number",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyNumberDefault();
            },
            ListParts()),

        new Scenario("list-apply-outline",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyOutlineNumberDefault();
            },
            ListParts()),

        new Scenario("list-apply-on-heading",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Paragraphs[0].Range.Style = "Heading 1";
                d.Paragraphs[0].Range.ListFormat.ApplyBulletDefault();
            },
            ListParts()),

        new Scenario("list-apply-on-indented",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Paragraphs[0].Range.ParagraphFormat.LeftIndent = 72;
                d.Paragraphs[0].Range.ListFormat.ApplyNumberDefault();
            },
            ListParts()),

        new Scenario("list-indent-outdent",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyOutlineNumberDefault();
                d.Paragraphs[1].Range.ListFormat.ListIndent();
                d.Paragraphs[2].Range.ListFormat.ListIndent();
                d.Paragraphs[2].Range.ListFormat.ListOutdent();
            },
            ListParts()),

        new Scenario("list-outdent-first-level",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyOutlineNumberDefault();
                d.Paragraphs[0].Range.ListFormat.ListOutdent();
            },
            ListParts()),

        new Scenario("list-level-number",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyOutlineNumberDefault();
                d.Paragraphs[2].Range.ListFormat.ListLevelNumber = 4;
            },
            ListParts()),

        new Scenario("list-remove-numbers",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyOutlineNumberDefault();
                d.Paragraphs[1].Range.ListFormat.ListLevelNumber = 2;
                d.Paragraphs[1].Range.ListFormat.RemoveNumbers();
            },
            ListParts()),

        new Scenario("list-convert-bullet-to-text",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyBulletDefault();
                d.Paragraphs[1].Range.ListFormat.ConvertNumbersToText();
            },
            ListParts()),

        new Scenario("list-convert-number-to-text",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyNumberDefault();
                d.Paragraphs[0].Range.ListFormat.ConvertNumbersToText();
            },
            ListParts()),

        new Scenario("list-apply-template-continue",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Paragraphs[0].Range.ListFormat.ApplyNumberDefault();
                d.Paragraphs[1].Range.ListFormat.ApplyListTemplate(d.ListTemplates[0], true);
            },
            ListParts()),
        new Scenario("list-template-add",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Content.ListFormat.ApplyListTemplate(d.ListTemplates.Add(false, "Mine"));
            },
            ListParts()),
        new Scenario("list-template-add-outline",
            d =>
            {
                d.Content.InsertAfter("One\rTwo");
                d.Content.ListFormat.ApplyListTemplate(d.ListTemplates.Add(true, "Outline"));
            },
            ListParts()),
        new Scenario("list-template-add-unnamed",
            d =>
            {
                d.Content.InsertAfter("One");
                d.ListTemplates.Add();
            },
            ListParts()),
        new Scenario("list-level-font-tab",
            d =>
            {
                d.Content.InsertAfter("One\rTwo");
                ListTemplateWrapper template = d.ListTemplates.Add(false, "Styled");
                ListLevelWrapper level = template.ListLevels[0];
                level.NumberFormat = "%1.";
                level.NumberStyle = ListNumberStyle.Arabic;
                level.Font.Bold = true;
                level.Font.Size = 14;
                level.Font.Name = "Arial";
                level.Font.Color = 255;
                level.TabPosition = 100;
                d.Content.ListFormat.ApplyListTemplate(template);
            },
            ListParts()),
        new Scenario("list-apply-with-level",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                d.Paragraphs[0].Range.ListFormat.ApplyNumberDefault();
                ListTemplateWrapper template = d.ListTemplates[0];
                RangeWrapper range = d.Range(d.Paragraphs[1].Range.Start, d.Paragraphs[2].Range.End);
                range.ListFormat.ApplyListTemplateWithLevel(template, true, ListApplyTo.WholeList, DefaultListBehavior.Word10, 2);
            },
            ListParts()),
        new Scenario("list-apply-template-in-use",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                ListTemplateWrapper template = d.ListTemplates.Add(false, "A");
                d.Paragraphs[0].Range.ListFormat.ApplyListTemplate(template);
                d.Paragraphs[1].Range.ListFormat.ApplyListTemplate(template, false);
                d.Paragraphs[2].Range.ListFormat.ApplyListTemplate(template, false);
            },
            ListParts()),        new Scenario("selection-insert-rows-below",
            d =>
            {
                d.Content.InsertAfter("Before\r");
                TableWrapper table = d.Tables.Add(d.Range(d.Content.End - 1, d.Content.End - 1), 2, 2);
                table.Cell(0, 0).Range.Text = "a";
                table.Cell(1, 0).Range.Text = "c";
                table.Cell(0, 0).Select();
                d.Selection.InsertRowsBelow(2);
            },
            Body()),
        new Scenario("selection-insert-rows-above",
            d =>
            {
                d.Content.InsertAfter("Before\r");
                TableWrapper table = d.Tables.Add(d.Range(d.Content.End - 1, d.Content.End - 1), 2, 2);
                table.Cell(0, 0).Range.Text = "a";
                table.Cell(1, 0).Range.Text = "c";
                table.Cell(1, 0).Select();
                d.Selection.InsertRowsAbove(1);
            },
            Body()),
        new Scenario("selection-clear-character",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Range(0, 5).Style = "Strong";
                d.Range(0, 5).Font.Size = 20;
                d.Range(6, 11).Style = "Strong";
                d.Range(6, 11).Font.Size = 20;
                d.Range(0, 5).Select();
                d.Selection.ClearCharacterStyle();
                d.Range(6, 11).Select();
                d.Selection.ClearCharacterDirectFormatting();
            },
            Body(), Part("styles.xml")),
        new Scenario("selection-clear-character-all",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Range(0, 5).Style = "Strong";
                d.Range(0, 5).Font.Size = 20;
                d.Range(0, 5).Select();
                d.Selection.ClearCharacterAllFormatting();
            },
            Body()),
        new Scenario("selection-clear-paragraph",
            d =>
            {
                d.Content.InsertAfter("One\rTwo\rThree");
                for (int i = 0; i < 3; i++)
                {
                    d.Paragraphs[i].Style = "Heading 1";
                    d.Paragraphs[i].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                }
                d.Paragraphs[0].Range.Select();
                d.Selection.ClearParagraphStyle();
                d.Paragraphs[1].Range.Select();
                d.Selection.ClearParagraphDirectFormatting();
                d.Paragraphs[2].Range.Select();
                d.Selection.ClearParagraphAllFormatting();
            },
            Body()),
        new Scenario("selection-bold-run",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Range(0, 5).Select();
                d.Selection.BoldRun();
                d.Range(6, 11).Select();
                d.Selection.BoldRun();
                d.Selection.BoldRun();
                d.Range(12, 17).Select();
                d.Selection.ItalicRun();
            },
            Body()),        new Scenario("footnote-add",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "The note text");
            },
            NoteParts()),

        new Scenario("footnote-add-to-range",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 10));
            },
            NoteParts()),

        new Scenario("footnote-add-custom-mark",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), "ref*", "custom mark");
            },
            NoteParts()),

        new Scenario("footnote-add-before-another",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(11, 11), null, "second in the file");
                d.Footnotes.Add(d.Range(5, 5), null, "first in the text");
            },
            NoteParts()),

        new Scenario("footnote-delete",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Footnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes[0].Delete();
            },
            NoteParts()),

        new Scenario("endnote-add",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Endnotes.Add(d.Range(5, 5), null, "The end note");
            },
            NoteParts()),

        new Scenario("footnote-and-endnote",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Endnotes.Add(d.Range(11, 11), null, "The end note");
                d.Footnotes.Add(d.Range(5, 5), null, "The footnote");
            },
            NoteParts()),
        new Scenario("note-option-footnote-style",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.NumberStyle = NoteNumberStyle.LowercaseRoman;
            },
            NoteOptionParts()),

        new Scenario("note-option-footnote-start",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.StartingNumber = 5;
            },
            NoteOptionParts()),

        new Scenario("note-option-footnote-restart-section",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.NumberingRule = NumberingRule.RestartSection;
            },
            NoteOptionParts()),

        new Scenario("note-option-footnote-restart-page",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.NumberingRule = NumberingRule.RestartPage;
            },
            NoteOptionParts()),

        new Scenario("note-option-footnote-location",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.Location = FootnoteLocation.BeneathText;
            },
            NoteOptionParts()),

        new Scenario("note-option-footnote-symbol",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.NumberStyle = NoteNumberStyle.Symbol;
            },
            NoteOptionParts()),

        new Scenario("note-option-endnote-section-end",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Endnotes.Location = EndnoteLocation.EndOfSection;
            },
            NoteOptionParts()),

        new Scenario("note-option-endnote-decimal-start",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Endnotes.NumberStyle = NoteNumberStyle.Arabic;
                d.Endnotes.StartingNumber = 3;
            },
            NoteOptionParts()),

        new Scenario("note-convert-footnotes",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.Convert();
            },
            NoteOptionParts()),

        new Scenario("note-convert-endnotes",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Endnotes.Convert();
            },
            NoteOptionParts()),

        new Scenario("note-swap",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Footnotes.Add(d.Range(5, 5), null, "one");
                d.Endnotes.Add(d.Range(11, 11), null, "two");
                d.Footnotes.SwapWithEndnotes();
            },
            NoteOptionParts()),

        new Scenario("content-control-placeholders",
            d =>
            {
                d.Content.InsertAfter("A\rB\rC");
                ContentControlType[] types = { ContentControlType.Text, ContentControlType.DropdownList, ContentControlType.Date };
                for (int i = 0; i < 3; i++)
                {
                    RangeWrapper place = d.Paragraphs[i].Range;
                    place.Collapse();
                    d.ContentControls.Add(types[i], place);
                }
            },
            ControlParts()),

        new Scenario("content-control-rich-text",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
            },
            ControlParts()),

        new Scenario("content-control-text",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.Text, d.Range(6, 11));
            },
            ControlParts()),

        new Scenario("content-control-combo-box",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.ComboBox, d.Range(6, 11));
            },
            ControlParts()),

        new Scenario("content-control-dropdown-list",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.DropdownList, d.Range(6, 11));
            },
            ControlParts()),

        new Scenario("content-control-date",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.Date, d.Range(6, 11));
            },
            ControlParts()),

        new Scenario("content-control-check-box",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.CheckBox, d.Range(6, 6));
            },
            Body()),

        new Scenario("content-control-collapsed-text",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.Text, d.Range(5, 5));
            },
            ControlParts()),

        new Scenario("content-control-collapsed-date",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.Date, d.Range(5, 5));
            },
            ControlParts()),

        new Scenario("content-control-collapsed-dropdown",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.ContentControls.Add(ContentControlType.DropdownList, d.Range(5, 5));
            },
            ControlParts()),

        new Scenario("content-control-properties",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
                cc.Title = "My title";
                cc.Tag = "my-tag";
                cc.LockContentControl = true;
                cc.LockContents = true;
                cc.Appearance = ContentControlAppearance.Hidden;
            },
            ControlParts()),

        new Scenario("content-control-lock-control-only",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
                cc.LockContentControl = true;
            },
            ControlParts()),

        new Scenario("content-control-lock-contents-only",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
                cc.LockContents = true;
            },
            ControlParts()),

        new Scenario("content-control-delete-keep",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
                cc.Delete(false);
            },
            Body()),

        new Scenario("content-control-delete-contents",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
                cc.Delete(true);
            },
            Body()),

        new Scenario("content-control-check-box-checked",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.CheckBox, d.Range(6, 6));
                cc.Checked = true;
            },
            Body()),

        new Scenario("content-control-check-box-symbols",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.CheckBox, d.Range(6, 6));
                cc.SetCheckedSymbol(10004, "Segoe UI Symbol");
                cc.SetUncheckedSymbol(10006, "Segoe UI Symbol");
                cc.Checked = true;
            },
            Body()),

        new Scenario("content-control-dropdown-entries",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.DropdownList, d.Range(5, 5));
                cc.DropdownListEntries.Add("One", "1");
                cc.DropdownListEntries.Add("Two");
                cc.DropdownListEntries.Add("Zero", "0", 0);
            },
            ControlParts()),

        new Scenario("content-control-dropdown-select",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.DropdownList, d.Range(5, 5));
                cc.DropdownListEntries.Add("One", "1");
                cc.DropdownListEntries.Add("Two", "2");
                cc.DropdownListEntries[1].Select();
            },
            ControlParts()),

        new Scenario("content-control-dropdown-clear",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.DropdownList, d.Range(5, 5));
                cc.DropdownListEntries.Add("One", "1");
                cc.DropdownListEntries.Add("Two", "2");
                cc.DropdownListEntries[0].Delete();
                cc.DropdownListEntries.Add("Three", "3");
                cc.DropdownListEntries[0].MoveDown();
            },
            ControlParts()),

        new Scenario("content-control-combo-entries",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.ComboBox, d.Range(5, 5));
                cc.DropdownListEntries.Add("One", "1");
                cc.DropdownListEntries.Add("Two", "2");
                cc.DropdownListEntries.Clear();
                cc.DropdownListEntries.Add("Only", "only");
            },
            ControlParts()),

        new Scenario("content-control-date-format",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.Date, d.Range(5, 5));
                cc.DateDisplayFormat = "dddd, d MMMM yyyy";
                cc.DateStorageFormat = ContentControlDateStorageFormat.Date;
                cc.DateDisplayLocale = 1041;
                cc.DateCalendarType = CalendarType.Japan;
            },
            ControlParts()),

        new Scenario("content-control-multiline",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.Text, d.Range(5, 5));
                cc.MultiLine = true;
            },
            ControlParts()),

        new Scenario("content-control-color-temporary",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.RichText, d.Range(6, 11));
                cc.Color = 255;
                cc.Temporary = true;
            },
            ControlParts()),

        new Scenario("content-control-default-text-style",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.Text, d.Range(6, 11));
                cc.DefaultTextStyle = "Strong";
            },
            Body()),

        new Scenario("custom-xml-part-add",
            d =>
            {
                d.Content.InsertAfter("Hello");
                d.CustomXMLParts.Add("<root xmlns=\"urn:test\"><name>Bob</name></root>");
            },
            CustomXmlParts()),

        new Scenario("custom-xml-part-delete",
            d =>
            {
                d.Content.InsertAfter("Hello");
                d.CustomXMLParts.Add("<root xmlns=\"urn:test\"><name>Bob</name></root>").Delete();
            },
            Part("document.xml.rels"), Part("content-types.xml")),

        new Scenario("content-control-xml-mapping",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                CustomXMLPart part = d.CustomXMLParts.Add("<root xmlns=\"urn:test\"><name>Bob</name></root>");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.Text, d.Range(5, 5));
                cc.XMLMapping.SetMapping("/ns:root/ns:name", "xmlns:ns='urn:test'", part);
            },
            Body(), Part("glossary.document.xml"), Part("customXml.itemProps1.xml")),

        new Scenario("content-control-xml-mapping-delete",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                CustomXMLPart part = d.CustomXMLParts.Add("<root xmlns=\"urn:test\"><name>Bob</name></root>");
                ContentControlWrapper cc = d.ContentControls.Add(ContentControlType.Text, d.Range(5, 5));
                cc.XMLMapping.SetMapping("/ns:root/ns:name", "xmlns:ns='urn:test'", part);
                cc.XMLMapping.Delete();
            },
            Body(), Part("glossary.document.xml")),

        new Scenario("inline-picture-add",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5));
            },
            PictureParts()),

        new Scenario("inline-picture-twice",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5));
                d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(0, 0));
            },
            PictureParts()),

        new Scenario("inline-picture-properties",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                InlineShapeWrapper shape = d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5));
                shape.AlternativeText = "A blue box";
                shape.Title = "Box";
                shape.Width = 60;
                shape.LockAspectRatio = false;
                shape.Height = 30;
            },
            Body()),

        new Scenario("inline-picture-scale",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                InlineShapeWrapper shape = d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5));
                shape.ScaleWidth = 200;
            },
            Body()),

        new Scenario("inline-picture-lock",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                InlineShapeWrapper shape = d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5));
                shape.LockAspectRatio = true;
                shape.Width = 60;
            },
            Body()),

        new Scenario("inline-picture-linked",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.InlineShapes.AddPicture(PicturePath, true, false, d.Range(5, 5));
            },
            PictureParts()),

        new Scenario("inline-picture-delete",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5)).Delete();
            },
            PictureParts()),

        new Scenario("inline-picture-over-text",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(6, 11));
            },
            Body()),

        new Scenario("toc-add",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rSecond\rMore text\rSub heading\rEnd", (0, "Heading 1"), (2, "Heading 1"), (4, "Heading 2"));
                d.TablesOfContents.Add(d.Range(0, 0));
            },
            Body()),

        new Scenario("toc-added-styles",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rTitled\rMore text\rSub\rEnd", (0, "Heading 1"), (2, "Title"), (4, "Subtitle"));
                TableOfContentsWrapper toc = d.TablesOfContents.Add(d.Range(0, 0), true, 1, 3, false, null, true, true, "Title,1,Subtitle,2");
                Assert.Equal(2, toc.HeadingStyles.Count);
            },
            Body()),

        new Scenario("toc-added-styles-nopages",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rTitled\rEnd", (0, "Heading 1"), (2, "Title"));
                d.TablesOfContents.Add(d.Range(0, 0), true, 1, 3, false, null, false, false, "Title,2");
            },
            Body()),

        new Scenario("toc-added-styles-outline",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rTitled\rEnd", (0, "Heading 1"), (2, "Title"));
                d.TablesOfContents.Add(d.Range(0, 0), true, 1, 3, false, null, true, true, "Title,2", false, false, true);
            },
            Body()),
        new Scenario("caption-add",
            d =>
            {
                d.Content.InsertAfter("Intro\rPicture one\rEnd");
                d.Paragraphs[1].Range.InsertCaption("Figure", ": First");
            },
            Body()),

        new Scenario("caption-above",
            d =>
            {
                d.Content.InsertAfter("Intro\rPicture one\rEnd");
                d.Paragraphs[1].Range.InsertCaption("Figure", ": First", string.Empty, CaptionPosition.Above);
            },
            Body()),

        new Scenario("caption-exclude-label",
            d =>
            {
                d.Content.InsertAfter("Intro\rPicture one\rEnd");
                d.Paragraphs[1].Range.InsertCaption("Figure", ": First", string.Empty, CaptionPosition.Above, true);
            },
            Body()),

        new Scenario("tof-add",
            d =>
            {
                d.Content.InsertAfter("Intro\rPicture one\rMiddle\rPicture two\rEnd");
                d.Paragraphs[3].Range.InsertCaption("Figure", ": Second");
                d.Paragraphs[1].Range.InsertCaption("Figure", ": First");
                TableOfFiguresWrapper table = d.TablesOfFigures.Add(d.Range(0, 0), "Figure");
                Assert.Equal("Figure", table.Caption);
                Assert.True(table.IncludeLabel);
            },
            Body()),

        new Scenario("tof-options",
            d =>
            {
                d.Content.InsertAfter("Intro\rPicture one\rEnd");
                d.Paragraphs[1].Range.InsertCaption("Figure", ": First");
                d.TablesOfFigures.Add(d.Range(0, 0), "Figure", false, false, 1, 9, false, null, false, false, null, true);
            },
            Body()),

        new Scenario("toc-mark-entry",
            d =>
            {
                d.Content.InsertAfter("Intro text here\rSecond bit");
                d.TablesOfContents.MarkEntry(d.Range(16, 22), "Second", null, null, 1);
                d.TablesOfContents.MarkEntry(d.Range(0, 5), "Marked one", null, "A", 2);
                TableOfContentsWrapper toc = d.TablesOfContents.Add(d.Range(0, 0), false, 1, 3, true, "A");
                Assert.True(toc.UseFields);
            },
            Body()),
        new Scenario("toc-options",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rSecond\rMore text\rSub heading\rEnd", (0, "Heading 1"), (2, "Heading 1"), (4, "Heading 2"));
                d.TablesOfContents.Add(d.Range(0, 0), true, 1, 1, false, null, false, false, null, false, true);
            },
            Body()),

        new Scenario("toc-hyperlinks",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rSub heading", (0, "Heading 1"), (2, "Heading 2"));
                d.TablesOfContents.Add(d.Range(0, 0), true, 1, 3, false, null, true, true, null, true);
            },
            Body()),

        new Scenario("toc-flags",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rSub heading", (0, "Heading 1"), (2, "Heading 2"));
                d.TablesOfContents.Add(d.Range(0, 0), true, 2, 3, true, "A", true, true, null, true, false, true);
            },
            Body()),

        new Scenario("toc-margins",
            d =>
            {
                d.PageSetup.LeftMargin = 100;
                HeadedDocument(d, "First\rSome text", (0, "Heading 1"));
                d.TablesOfContents.Add(d.Range(0, 0));
            },
            Body()),

        new Scenario("toc-empty",
            d =>
            {
                d.Content.InsertAfter("Plain\rtext");
                d.TablesOfContents.Add(d.Range(0, 0));
            },
            Body()),

        new Scenario("toc-delete",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rSecond", (0, "Heading 1"), (2, "Heading 1"));
                d.TablesOfContents.Add(d.Range(0, 0)).Delete();
            },
            Body()),

        new Scenario("toc-update",
            d =>
            {
                HeadedDocument(d, "First\rSome text\rSecond", (0, "Heading 1"));
                TableOfContentsWrapper toc = d.TablesOfContents.Add(d.Range(0, 0));
                d.Paragraphs[d.Paragraphs.Count - 1].Style = "Heading 1";
                toc.Update();
            },
            Body()),

        new Scenario("shape-add-picture",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5));
            },
            PictureParts()),

        new Scenario("shape-two-pictures",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ShapeWrapper a = d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5));
                ShapeWrapper b = d.Shapes.AddPicture(PicturePath, false, true, 10, 10, 60, 30, d.Range(5, 5));
                b.ZOrder(ZOrderCommand.SendToBack);
            },
            Body()),

        new Scenario("shape-picture-properties",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ShapeWrapper s = d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5));
                s.Name = "Box";
                s.AlternativeText = "A box";
                s.Title = "Title";
                s.Rotation = 30;
                s.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                s.RelativeVerticalPosition = RelativeVerticalPosition.Margin;
                s.Left = 40;
                s.Top = 50;
                s.WrapFormat.DistanceLeft = 10;
                s.WrapFormat.AllowOverlap = false;
                s.LockAnchor = true;
                s.LockAspectRatio = true;
                s.Visible = false;
            },
            Body()),

        new Scenario("shape-picture-wraps",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                foreach (WrapType type in new[] { WrapType.Square, WrapType.Tight, WrapType.Through, WrapType.TopBottom, WrapType.Behind })
                {
                    d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5)).WrapFormat.Type = type;
                }
                ShapeWrapper left = d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5));
                left.WrapFormat.Type = WrapType.Square;
                left.WrapFormat.Side = WrapSideType.Left;
                ShapeWrapper largest = d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5));
                largest.WrapFormat.Type = WrapType.Square;
                largest.WrapFormat.Side = WrapSideType.Largest;
            },
            DrawingElements("wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "wrapNone", "wrapPolygon")),

        new Scenario("shape-add-textbox",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ShapeWrapper s = d.Shapes.AddTextbox(TextOrientation.Horizontal, 20, 30, 120, 60, d.Range(5, 5));
                s.TextFrame.TextRange.Text = "Boxed";
            },
            DrawingElements("wsp", "docPr")),

        new Scenario("shape-add-rectangle",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Shapes.AddShape(AutoShapeType.Rectangle, 20, 30, 80, 40, d.Range(5, 5));
            },
            DrawingElements("wsp", "docPr")),

        new Scenario("shape-add-line",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Shapes.AddLine(20, 30, 100, 70, d.Range(5, 5));
            },
            DrawingElements("wsp", "docPr")),

        new Scenario("shape-autoshapes",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                foreach (AutoShapeType type in new[] { AutoShapeType.Rectangle, AutoShapeType.Diamond, AutoShapeType.RoundedRectangle, AutoShapeType.IsoscelesTriangle,
                    AutoShapeType.Oval, AutoShapeType.Hexagon, AutoShapeType.RightArrow })
                {
                    d.Shapes.AddShape(type, 20, 30, 80, 40, d.Range(5, 5));
                }
            },
            DrawingElements("wsp", "docPr")),

        new Scenario("shape-fill-line",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                ShapeWrapper a = d.Shapes.AddShape(AutoShapeType.Rectangle, 20, 30, 80, 40, d.Range(5, 5));
                a.Fill.ForeColor.RGB = 255;
                a.Line.ForeColor.RGB = 65280;
                a.Line.Weight = 3;
                ShapeWrapper c = d.Shapes.AddTextbox(TextOrientation.Horizontal, 20, 150, 80, 40, d.Range(5, 5));
                c.Fill.ForeColor.RGB = 16711680;
                c.Fill.Transparency = 0.5;
            },
            DrawingElements("wsp", "docPr")),
        new Scenario("shape-delete",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5)).Delete();
            },
            PictureParts()),

        new Scenario("shape-to-inline",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.Shapes.AddPicture(PicturePath, false, true, 20, 30, 60, 30, d.Range(5, 5)).ConvertToInlineShape();
            },
            Body()),

        new Scenario("shape-from-inline",
            d =>
            {
                d.Content.InsertAfter("Hello world again");
                d.InlineShapes.AddPicture(PicturePath, false, true, d.Range(5, 5)).ConvertToShape();
            },
            DrawingElements("docPr", "extent", "wrapNone")),

        new Scenario("document-settings",
            d =>
            {
                d.DefaultTabStop = 54;
                d.AutoHyphenation = true;
                d.HyphenateCaps = false;
                d.ConsecutiveHyphensLimit = 2;
                d.UpdateStylesOnOpen = true;
                d.EmbedTrueTypeFonts = true;
                d.SaveSubsetFonts = true;
                d.DoNotEmbedSystemFonts = false;
                d.RemovePersonalInformation = true;
                d.RemoveDateAndTime = true;
            },
            Elements("settings.xml", "removePersonalInformation", "removeDateAndTime", "embedTrueTypeFonts", "embedSystemFonts", "saveSubsetFonts", "linkStyles",
                "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "doNotHyphenateCaps")),

        new Scenario("default-table-style",
            d => d.SetDefaultTableStyle("Table Grid", false),
            Elements("settings.xml", "defaultTableStyle")),

        new Scenario("heading-insert-paragraph-after",
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("built-in-styles", ApplyEveryBuiltInStyle, Part("styles.xml")),

        new Scenario("font-properties-on",
            d =>
            {
                d.Content.InsertAfter(FontWords());
                for (int i = 0; i < 20; i++)
                {
                    SetFontProperty(d.Range(4 * i, (4 * i) + 3).Font, i, true);
                }
            },
            Body()),

        new Scenario("font-properties-on-then-off",
            d =>
            {
                d.Content.InsertAfter(FontWords());
                for (int i = 0; i < 20; i++)
                {
                    SetFontProperty(d.Range(4 * i, (4 * i) + 3).Font, i, true);
                    SetFontProperty(d.Range(4 * i, (4 * i) + 3).Font, i, false);
                }
            },
            Body()),

        new Scenario("bold-off-in-a-bold-style",
            d =>
            {
                d.Content.InsertAfter("Strong text");
                d.Range(0, 6).Style = "Strong";
                d.Range(0, 6).Font.Bold = false;
            },
            Body()),

        new Scenario("runs-merge-when-formatting-matches",
            d =>
            {
                d.Content.InsertAfter("One two three");
                d.Range(0, 3).Font.Bold = true;
                d.Range(3, 7).Font.Bold = true;
            },
            Body()),

        new Scenario("paragraph-properties-on",
            d =>
            {
                AddParagraphs(d, 17);
                for (int i = 0; i < 17; i++)
                {
                    SetParagraphProperty(d.Paragraphs[i].Range.ParagraphFormat, i, true);
                }
            },
            Body()),

        new Scenario("paragraph-properties-on-then-off",
            d =>
            {
                AddParagraphs(d, 17);
                for (int i = 0; i < 17; i++)
                {
                    SetParagraphProperty(d.Paragraphs[i].Range.ParagraphFormat, i, true);
                    SetParagraphProperty(d.Paragraphs[i].Range.ParagraphFormat, i, false);
                }
            },
            Body()),

        new Scenario("paragraph-style-over-most-of-a-format",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 6).Font.Italic = true;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-over-less-of-a-format",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 5).Font.Italic = true;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-counts-the-mark",
            d =>
            {
                d.Content.InsertAfter("012345678");
                d.Range(0, 4).Font.Italic = true;
                d.Range(9, 10).Font.Italic = true;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-on-a-new-paragraph",
            d =>
            {
                d.Content.InsertAfter("First");
                d.Paragraphs[0].Range.Font.Italic = true;
                d.Content.InsertParagraphAfter();
                d.Paragraphs[1].Range.InsertBefore("Contents");
                d.Paragraphs[1].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-by-property",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 7).Font.Bold = true;
                d.Range(0, 3).Font.Italic = true;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-by-value",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 3).Font.Size = 30;
                d.Range(3, 7).Font.Size = 22;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-first-of-equal-values",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 3).Font.Size = 30;
                d.Range(3, 6).Font.Size = 22;
                d.Range(6, 9).Font.Size = 18;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-value-of-the-style-counts",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 3).Font.Size = 20;
                d.Range(3, 7).Font.Size = 22;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("insert-symbol-ansi",
            d =>
            {
                d.Content.InsertAfter("ab");
                d.Range(1, 1).InsertSymbol(171, "Symbol", false);
            },
            Body()),

        new Scenario("insert-symbol-unicode",
            d =>
            {
                d.Content.InsertAfter("ab");
                d.Range(1, 1).InsertSymbol(8364, "Calibri", true);
            },
            Body()),

        new Scenario("insert-symbol-unicode-in-bold-text",
            d =>
            {
                d.Content.InsertAfter("ab");
                d.Range(0, 2).Font.Bold = true;
                d.Range(1, 1).InsertSymbol(169, "Arial", true);
            },
            Body()),

        new Scenario("insert-symbol-replaces-range",
            d =>
            {
                d.Content.InsertAfter("ab");
                d.Range(0, 1).InsertSymbol(169, "Times New Roman", true);
            },
            Body()),

        new Scenario("insert-alignment-tab-right-margin",
            d =>
            {
                d.Content.InsertAfter("ab");
                d.Range(1, 1).InsertAlignmentTab(AlignmentTabAlignment.Right, AlignmentTabRelative.Margin);
            },
            Body()),

        new Scenario("insert-alignment-tab-center-indent",
            d =>
            {
                d.Content.InsertAfter("ab");
                d.Range(1, 1).InsertAlignmentTab(AlignmentTabAlignment.Center, AlignmentTabRelative.Indent);
            },
            Body()),

        new Scenario("case-upper-part-of-bold-run",
            d =>
            {
                d.Content.InsertAfter("abc DEF");
                d.Range(0, 3).Font.Bold = true;
                d.Range(0, 3).Case = CharacterCase.UpperCase;
            },
            Body()),

        new Scenario("case-upper-inside-a-run",
            d =>
            {
                d.Content.InsertAfter("abcdef");
                d.Range(2, 4).Case = CharacterCase.UpperCase;
            },
            Body()),

        new Scenario("case-title-sentence",
            d =>
            {
                d.Content.InsertAfter("hello wORLD. second sentence here");
                d.Range(0, 33).Case = CharacterCase.TitleSentence;
            },
            Body()),

        new Scenario("formatted-text-copy-bold-to-end",
            d =>
            {
                d.Content.InsertAfter("abc def ghi");
                d.Range(4, 7).Font.Bold = true;
                RangeWrapper target = d.Range(11, 11);
                target.FormattedText = d.Range(4, 7).FormattedText;
            },
            Body()),

        new Scenario("formatted-text-copy-across-paragraphs",
            d =>
            {
                d.Content.InsertAfter("aaa\rbbb\rccc");
                d.Paragraphs[1].Format.Alignment = ParagraphAlignment.Center;
                RangeWrapper target = d.Range(11, 11);
                target.FormattedText = d.Range(2, 6).FormattedText;
            },
            Body()),

        new Scenario("formatted-text-copy-paragraph-with-mark",
            d =>
            {
                d.Content.InsertAfter("aaa\rbbb\rccc");
                d.Paragraphs[1].Style = "Heading 1";
                RangeWrapper target = d.Range(0, 0);
                target.FormattedText = d.Paragraphs[1].Range.FormattedText;
            },
            Body()),

        new Scenario("paragraph-style-same-value",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 3).Font.Size = 28;
                d.Range(3, 6).Font.Size = 30;
                d.Paragraphs[0].Style = "Title";
            },
            Body()),

        new Scenario("paragraph-style-clears-paragraph-format",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                ParagraphFormatWrapper format = d.Paragraphs[0].Range.ParagraphFormat;
                format.Alignment = ParagraphAlignment.Center;
                format.SpaceAfter = 30;
                format.LeftIndent = 20;
                format.KeepWithNext = true;
                format.PageBreakBefore = true;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-keeps-some-format",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 7).Font.Superscript = true;
                d.Range(0, 7).Font.Scaling = 150;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-underline-color",
            d =>
            {
                d.Content.InsertAfter("0123456789");
                d.Range(0, 7).Font.Underline = UnderlineStyle.SingleLine;
                d.Range(0, 7).Font.UnderlineColor = "FF0000";
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("paragraph-style-keeps-a-section",
            d =>
            {
                d.Content.InsertAfter("First");
                d.Sections.Add();
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs[0].Style = "Heading 1";
            },
            Body()),

        new Scenario("normal-style-applied",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Style = "Normal";
            },
            Body()),

        new Scenario("insert-paragraph-after-formatted",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Range(0, 4).Font.Bold = true;
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("insert-paragraph-before-formatted",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Range(0, 4).Font.Bold = true;
                d.Paragraphs[0].Range.InsertParagraphBefore();
            },
            Body()),

        new Scenario("paragraphs-add",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs.Add();
            },
            Body()),

        new Scenario("split-a-formatted-paragraph",
            d =>
            {
                d.Content.InsertAfter("One two");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Range(3, 3).InsertParagraphAfter();
            },
            Body()),

        new Scenario("insert-paragraph-after-middle",
            d =>
            {
                d.Content.InsertAfter("One");
                d.Content.InsertParagraphAfter();
                d.Content.InsertAfter("Two");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs[1].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("heading-with-format-insert-paragraph-after",
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("split-a-heading-at-its-end",
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Range(7, 7).InsertParagraphAfter();
            },
            Body()),

        new Scenario("split-a-heading-in-the-middle",
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Range(3, 3).InsertParagraphAfter();
            },
            Body()),

        new Scenario("content-text-with-paragraphs", d => d.Content.Text = "One\rTwo", Body()),

        new Scenario("insert-before-with-a-paragraph",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(0, 0).InsertBefore("A\rB");
            },
            Body()),

        new Scenario("join-formatted-paragraphs",
            d =>
            {
                d.Content.InsertAfter("First");
                d.Content.InsertParagraphAfter();
                d.Content.InsertAfter("Second");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs[1].Range.ParagraphFormat.Alignment = ParagraphAlignment.Right;
                TextUnitCollection characters = d.Paragraphs[0].Range.Characters;
                characters[characters.Count - 1].Delete();
            },
            Body()),

        new Scenario("delete-a-section-break",
            d =>
            {
                d.Content.InsertAfter("First");
                d.Sections.Add();
                d.Content.InsertAfter("Second");
                d.Sections[0].PageSetup.LeftMargin = 36;
                TextUnitCollection characters = d.Sections[0].Range.Characters;
                characters[characters.Count - 1].Delete();
            },
            WholeBody()),

        new Scenario("breaks",
            d =>
            {
                d.Content.InsertAfter("Line Page Column");
                d.Range(11, 11).InsertBreak(BreakType.ColumnBreak);
                d.Range(5, 5).InsertBreak(BreakType.PageBreak);
                d.Range(0, 0).InsertBreak(BreakType.LineBreak);
            },
            Body()),

        new Scenario("page-break-at-paragraph-start",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(0, 0).InsertBreak(BreakType.PageBreak);
            },
            Body()),

        new Scenario("page-break-at-paragraph-end",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(4, 4).InsertBreak(BreakType.PageBreak);
            },
            Body()),

        new Scenario("column-break-at-paragraph-start",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(0, 0).InsertBreak(BreakType.ColumnBreak);
            },
            Body()),

        new Scenario("column-break-at-paragraph-end",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(4, 4).InsertBreak(BreakType.ColumnBreak);
            },
            Body()),

        new Scenario("line-break-at-paragraph-end",
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(4, 4).InsertBreak(BreakType.LineBreak);
            },
            Body()),

        new Scenario("delete-text-within-a-run",
            d =>
            {
                d.Content.InsertAfter("One two three");
                d.Range(3, 7).Delete();
            },
            Body()),

        new Scenario("style-borders-paragraph",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Paragraph);
                s.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;
                s.Borders.OutsideColor = ColorHelper.RGB(255, 0, 0);
            },
            RefStyles()),
        new Scenario("style-borders-top",
            d => d.Styles.Add("RefStyle", StyleType.Paragraph).Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine,
            RefStyles()),
        new Scenario("style-shading-paragraph",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Paragraph);
                s.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
                s.Shading.Texture = ShadingTexture.Percent10;
            },
            RefStyles()),
        new Scenario("style-borders-shading-character",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Character);
                s.Font.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
                s.Font.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);
            },
            RefStyles()),
        new Scenario("style-font-borders-shading",
            d =>
            {
                StyleWrapper s = d.Styles.Add("RefStyle", StyleType.Paragraph);
                s.Font.Borders[BorderPosition.Top].LineStyle = BorderLineStyle.SingleLine;
                s.Font.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);
            },
            RefStyles()),
        new Scenario("table-style-add", d => d.Styles.Add("RefTable", StyleType.Table), RefStyles()),
        new Scenario("table-style-members",
            d =>
            {
                TableStyleWrapper t = d.Styles.Add("RefTable", StyleType.Table).Table;
                t.Alignment = RowAlignment.Center;
                t.AllowBreakAcrossPage = false;
                t.TopPadding = 5;
                t.BottomPadding = 6;
                t.LeftPadding = 7;
                t.RightPadding = 8;
                t.RowStripe = 2;
                t.ColumnStripe = 3;
                t.Spacing = 2;
                t.LeftIndent = 10;
            },
            RefStyles()),
        new Scenario("table-style-padding-top", d => d.Styles.Add("RefTable", StyleType.Table).Table.TopPadding = 5, RefStyles()),
        new Scenario("table-style-borders-shading",
            d =>
            {
                TableStyleWrapper t = d.Styles.Add("RefTable", StyleType.Table).Table;
                t.Borders.OutsideLineStyle = BorderLineStyle.SingleLine;
                t.Borders.InsideLineStyle = BorderLineStyle.SingleLine;
                t.Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
            },
            RefStyles()),
        new Scenario("table-style-conditional-mixed",
            d =>
            {
                ConditionalStyleWrapper c = d.Styles.Add("RefTable", StyleType.Table).Table.Condition(ConditionCode.FirstRow);
                c.Font.Bold = true;
                c.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                c.Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0);
            },
            RefStyles()),
        new Scenario("table-style-conditional-banding",
            d =>
            {
                TableStyleWrapper t = d.Styles.Add("RefTable", StyleType.Table).Table;
                t.Condition(ConditionCode.OddRowBanding).Shading.BackgroundPatternColor = ColorHelper.RGB(255, 0, 0);
                t.Condition(ConditionCode.LastColumn).Font.Italic = true;
                t.Condition(ConditionCode.SouthWestCell).Font.Bold = true;
            },
            RefStyles()),
        new Scenario("table-style-cond-shading-first", d => d.Styles.Add("RefTable", StyleType.Table).Table.Condition(ConditionCode.FirstRow).Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0), RefStyles()),
        new Scenario("table-style-cond-shading-band2", d => d.Styles.Add("RefTable", StyleType.Table).Table.Condition(ConditionCode.EvenRowBanding).Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0), RefStyles()),
        new Scenario("table-style-cond-shading-vband", d => d.Styles.Add("RefTable", StyleType.Table).Table.Condition(ConditionCode.OddColumnBanding).Shading.BackgroundPatternColor = ColorHelper.RGB(0, 255, 0), RefStyles()),
        new Scenario("table-style-cond-borders-band", d => d.Styles.Add("RefTable", StyleType.Table).Table.Condition(ConditionCode.OddRowBanding).Borders[BorderPosition.Bottom].LineStyle = BorderLineStyle.SingleLine, RefStyles()),
        new Scenario("table-style-cond-padding-only", d => d.Styles.Add("RefTable", StyleType.Table).Table.Condition(ConditionCode.FirstRow).TopPadding = 3, RefStyles()),
        new Scenario("find-match-prefix",
            d =>
            {
                d.Content.InsertAfter("pre prefix preach unpre repre");
                RangeWrapper r = d.Content;
                r.Find.MatchPrefix = true;
                r.Find.Execute("pre", replaceWith: "X", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-match-suffix",
            d =>
            {
                d.Content.InsertAfter("ing sing singing ingot");
                RangeWrapper r = d.Content;
                r.Find.MatchSuffix = true;
                r.Find.Execute("ing", replaceWith: "X", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-match-prefix-suffix",
            d =>
            {
                d.Content.InsertAfter("cat concat cats cat");
                RangeWrapper r = d.Content;
                r.Find.MatchPrefix = true;
                r.Find.MatchSuffix = true;
                r.Find.Execute("cat", replaceWith: "X", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-ignore-punct",
            d =>
            {
                d.Content.InsertAfter("co-op coop co,op co op");
                RangeWrapper r = d.Content;
                r.Find.IgnorePunct = true;
                r.Find.Execute("coop", replaceWith: "X", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-ignore-space",
            d =>
            {
                d.Content.InsertAfter("a b ab a  b ab");
                RangeWrapper r = d.Content;
                r.Find.IgnoreSpace = true;
                r.Find.Execute("ab", replaceWith: "X", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-match-phrase",
            d =>
            {
                d.Content.InsertAfter("a  b\tc\vd e and two  three");
                RangeWrapper r = d.Content;
                r.Find.MatchPhrase = true;
                r.Find.Execute("a b c d e", replaceWith: "X", replace: ReplaceMode.All);
                r.Find.Execute("two three", replaceWith: "Y", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-language",
            d =>
            {
                d.Content.InsertAfter("One two three");
                RangeWrapper r = d.Content;
                r.Find.Replacement.LanguageID = 1036;
                r.Find.Replacement.NoProofing = true;
                r.Find.Execute("two", format: true, replaceWith: "^&", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-font",
            d =>
            {
                d.Content.InsertAfter("One two three two");
                RangeWrapper r = d.Content;
                r.Find.Replacement.Font.Bold = true;
                r.Find.Replacement.Font.Size = 14;
                r.Find.Execute("two", format: true, replaceWith: "2", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-font-text-unchanged",
            d =>
            {
                d.Content.InsertAfter("One two three two");
                RangeWrapper r = d.Content;
                r.Find.Replacement.Font.Italic = true;
                r.Find.Execute("two", format: true, replaceWith: "^&", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-empty-with-format",
            d =>
            {
                d.Content.InsertAfter("One two three");
                RangeWrapper r = d.Content;
                r.Find.Replacement.Font.Italic = true;
                r.Find.Execute("two", format: true, replaceWith: string.Empty, replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-format-only",
            d =>
            {
                d.Content.InsertAfter("plain bold plain bold");
                d.Range(6, 10).Font.Bold = true;
                d.Range(17, 21).Font.Bold = true;
                RangeWrapper r = d.Content;
                r.Find.Font.Bold = true;
                r.Find.Replacement.Font.Italic = true;
                r.Find.Replacement.Font.Bold = false;
                r.Find.Execute(string.Empty, format: true, replaceWith: string.Empty, replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-style",
            d =>
            {
                d.Content.InsertAfter("One");
                d.Content.InsertParagraphAfter();
                d.Content.InsertAfter("Two");
                RangeWrapper r = d.Content;
                r.Find.Replacement.Style = "Heading 1";
                r.Find.Execute("One", format: true, replaceWith: "Uno", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-style-criteria",
            d =>
            {
                d.Content.InsertAfter("One");
                d.Content.InsertParagraphAfter();
                d.Content.InsertAfter("Two");
                d.Paragraphs[0].Style = "Heading 1";
                RangeWrapper r = d.Content;
                r.Find.Style = "Heading 1";
                r.Find.Replacement.Font.Italic = true;
                r.Find.Execute(string.Empty, format: true, replaceWith: string.Empty, replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-paragraphformat",
            d =>
            {
                d.Content.InsertAfter("One");
                d.Content.InsertParagraphAfter();
                d.Content.InsertAfter("Two");
                RangeWrapper r = d.Content;
                r.Find.Replacement.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                r.Find.Execute("Two", format: true, replaceWith: "Two", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("find-replace-highlight",
            d =>
            {
                d.Content.InsertAfter("One two three");
                RangeWrapper r = d.Content;
                r.Find.Replacement.Highlight = true;
                r.Find.Execute("two", format: true, replaceWith: "^&", replace: ReplaceMode.All);
            },
            Body()),

        new Scenario("replace-text-within-a-run",
            d =>
            {
                d.Content.InsertAfter("One two three");
                d.Range(4, 7).Text = "2";
            },
            Body()),

        //The results of these need page layout: Word has worked them out and the library leaves them for Word, so only their code is compared.
        new Scenario("field-from-code", d => d.Fields.Add(d.Range(0, 0), FieldType.Empty, "PAGE", false), Elements("document.xml", "instrText", "fldSimple")),

        new Scenario("field-page-with-merge-format", d => d.Fields.Add(d.Range(0, 0), FieldType.Page), Elements("document.xml", "instrText", "fldSimple")),

        new Scenario("field-date-without-merge-format", d => d.Fields.Add(d.Range(0, 0), FieldType.Date, string.Empty, false), Body()),

        new Scenario("field-date-from-code", d => d.Fields.Add(d.Range(0, 0), FieldType.Empty, "DATE", false), Body()),

        new Scenario("field-within-text",
            d =>
            {
                d.Content.InsertAfter("Before after");
                d.Fields.Add(d.Range(7, 7), FieldType.Date);
            },
            Body()),

        new Scenario("field-within-bold-text",
            d =>
            {
                d.Content.InsertAfter("Before after");
                d.Range(0, 12).Font.Bold = true;
                d.Fields.Add(d.Range(7, 7), FieldType.Date);
            },
            Body()),

        new Scenario("table-2x4-word9-behavior", d => d.Tables.Add(d.Range(0, 0), 2, 4, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed), Body()),

        new Scenario("table-autofit-content", d => d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word8TableBehavior, AutoFitBehavior.Content), Body()),

        new Scenario("table-autofit-window", d => d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word8TableBehavior, AutoFitBehavior.Window), Body()),

        new Scenario("table-word9-autofit-content", d => d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Content), Body()),

        new Scenario("table-word9-autofit-window", d => d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Window), Body()),

        new Scenario("table-nested",
            d =>
            {
                TableWrapper outer = d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                d.Tables.Add(outer.Cell(0, 0).Range, 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
            },
            Body()),

        new Scenario("table-in-text",
            d =>
            {
                d.Content.InsertAfter("Before after");
                d.Tables.Add(d.Range(7, 7), 1, 1);
            },
            Body()),

        new Scenario("cell-vertical-merge",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 3, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Cell(0, 0).Merge(table.Cell(1, 0));
            },
            Body()),

        new Scenario("cell-merge-block",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 3, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Cell(0, 0).Merge(table.Cell(1, 1));
            },
            Body()),

        new Scenario("cell-split",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Cell(0, 0).Split(1, 2);
            },
            Body()),

        new Scenario("row-properties",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 4, 1, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Rows[0].HeightRule = RowHeightRule.AtLeast;
                table.Rows[0].Height = 20;
                table.Rows[1].HeightRule = RowHeightRule.Exactly;
                table.Rows[1].Height = 20;
                table.Rows[2].HeadingFormat = true;
                table.Rows[3].AllowBreakAcrossPages = false;
            },
            Body()),

        new Scenario("heading-rows-off",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 3, 1, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Rows[1].HeadingFormat = true;
                table.Rows[0].HeadingFormat = false;
            },
            Body()),

        new Scenario("rows-add",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Cell(0, 0).Range.Text = "a";
                table.Rows.Add();
            },
            Body()),

        new Scenario("columns-add",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Columns.Add();
            },
            Body()),

        new Scenario("column-delete",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 1, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Columns[1].Delete();
            },
            Body()),

        new Scenario("column-width-narrower",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Columns[0].Width = 100;
            },
            Body()),

        new Scenario("column-width-wider",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Columns[0].Width = 200;
            },
            Body()),

        new Scenario("column-width-preferred-points",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.PreferredWidthType = PreferredWidthType.Points;
                table.PreferredWidth = 300;
                table.Columns[0].Width = 200;
            },
            Body()),

        new Scenario("column-width-after-add",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Columns.Add();
                table.Columns[0].Width = 120;
                table.Columns[3].Width = 70;
            },
            Body()),

        new Scenario("columns-add-within-space",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.Columns[0].Width = 40;
                table.Columns[1].Width = 40;
                table.Columns.Add();
            },
            Body()),

        new Scenario("columns-add-preferred-points",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.PreferredWidthType = PreferredWidthType.Points;
                table.PreferredWidth = 300;
                table.Columns.Add();
            },
            Body()),

        new Scenario("column-delete-preferred-points",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.PreferredWidthType = PreferredWidthType.Points;
                table.PreferredWidth = 300;
                table.Columns[1].Delete();
            },
            Body()),

        new Scenario("table-preferred-percent",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 2, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.PreferredWidthType = PreferredWidthType.Percent;
                table.PreferredWidth = 80;
            },
            Body()),

        new Scenario("table-borders-enable",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 1, 1);
                table.Borders.Enable = true;
            },
            Body()),

        new Scenario("table-properties",
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 1, 2, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed);
                table.PreferredWidthType = PreferredWidthType.Points;
                table.PreferredWidth = 200;
                table.Rows.Alignment = RowAlignment.Center;
                table.Rows.LeftIndent = 0;
                table.Cell(0, 0).Shading.BackgroundPatternColor = "FF0000";
                table.Cell(0, 1).VerticalAlignment = CellVerticalAlignment.Center;
            },
            Body()),

        new Scenario("section-properties",
            d =>
            {
                PageSetupWrapper setup = d.Sections[0].PageSetup;
                setup.Orientation = PageOrientation.Landscape;
                setup.Orientation = PageOrientation.Portrait;
                setup.SectionStart = SectionStart.Continuous;
                setup.SectionStart = SectionStart.NewPage;
            },
            WholeBody()),

        new Scenario("section-landscape", d => d.Sections[0].PageSetup.Orientation = PageOrientation.Landscape, WholeBody()),

        new Scenario("section-continuous", d => d.Sections[0].PageSetup.SectionStart = SectionStart.Continuous, WholeBody()),

        new Scenario("line-numbers", d => d.Sections[0].PageSetup.LineNumbering.Active = true, WholeBody()),

        new Scenario("text-columns", d => d.Sections[0].PageSetup.TextColumns.SetCount(2), WholeBody()),

        new Scenario("different-first-page", d => d.Sections[0].PageSetup.DifferentFirstPageHeaderFooter = true, WholeBody()),

        new Scenario("page-numbers-add", d => d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers.Add(), PageNumber()),

        new Scenario("page-numbers-add-left", d => d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers.Add(PageNumberAlignment.Left), PageNumber()),

        new Scenario("page-numbers-add-center", d => d.Sections[0].Footers[HeaderFooterIndex.Primary].PageNumbers.Add(PageNumberAlignment.Center), PageNumber()),

        new Scenario("page-numbers-add-to-text",
            d =>
            {
                HeaderFooterWrapper footer = d.Sections[0].Footers[HeaderFooterIndex.Primary];
                footer.Range.Text = "Footer";
                footer.PageNumbers.Add();
            },
            PageNumber()),

        new Scenario("two-sections-with-headers",
            d =>
            {
                d.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Header";
                d.Content.InsertAfter("First");
                d.Sections.Add();
            },
            BodyWithoutHeaderReferences()),

        new Scenario("header-unlink",
            d =>
            {
                d.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Header";
                d.Sections.Add();
                d.Sections[1].Headers[HeaderFooterIndex.Primary].LinkToPrevious = false;
            },
            HeaderFooter("header", "default")),

        new Scenario("comment-only",
            d =>
            {
                d.Content.InsertAfter("Commented text");
                AsReference(d.Comments.Add(d.Range(0, 9), "A comment"));
            },
            CommentParts()),

        new Scenario("comment-on-insertion-point",
            d =>
            {
                d.Content.InsertAfter("Text");
                AsReference(d.Comments.Add(d.Range(2, 2), "A comment"));
            },
            CommentParts()),

        new Scenario("comment-in-empty-paragraph", d => AsReference(d.Comments.Add(d.Range(0, 0), "A comment")), CommentParts()),

        new Scenario("comment-deleted",
            d =>
            {
                d.Content.InsertAfter("Commented text");
                CommentWrapper comment = d.Comments.Add(d.Range(0, 9), "A comment");
                AsReference(comment);
                AsReference(d.Comments.Add(d.Range(10, 14), "Another"));
                comment.Delete();
            },
            CommentParts()),

        new Scenario("join-paragraphs",
            d =>
            {
                d.Content.InsertAfter("First paragraph");
                d.Paragraphs[0].Range.InsertParagraphAfter();
                d.Paragraphs[1].Range.InsertBefore("Second paragraph");
                TextUnitCollection characters = d.Paragraphs[0].Range.Characters;
                characters[characters.Count - 1].Delete();
            },
            Body()),
    };

    //Word takes the author from Application.UserName, which the generator sets to Reference; the library takes the user's login name.
    private static void AsReference(CommentWrapper comment)
    {
        comment.Author = "Reference";
        comment.Initial = "R";
    }

    //Twenty three-letter words, one for each property below, as in scenarios.ps1: word i is at positions 4i to 4i + 3.
    private static string FontWords()
    {
        return string.Join(" ", Enumerable.Range(1, 20).Select(i => "w" + i.ToString("00", CultureInfo.InvariantCulture)));
    }

    private static void SetFontProperty(FontWrapper font, int index, bool on)
    {
        switch (index)
        {
            case 0: font.Italic = on; break;
            case 1: font.StrikeThrough = on; break;
            case 2: font.DoubleStrikeThrough = on; break;
            case 3: font.AllCaps = on; break;
            case 4: font.SmallCaps = on; break;
            case 5: font.Hidden = on; break;
            case 6: font.Emboss = on; break;
            case 7: font.Engrave = on; break;
            case 8: font.Outline = on; break;
            case 9: font.Shadow = on; break;
            case 10: font.Superscript = on; break;
            case 11: font.Subscript = on; break;
            case 12: font.Underline = on ? UnderlineStyle.SingleLine : UnderlineStyle.None; break;
            case 13: font.Size = on ? 14 : 12; break;
            case 14: font.Color = on ? WordColor.FromRgb(255, 0, 0) : WordColor.Automatic; break;
            case 15: font.Spacing = on ? 1.5 : 0; break;
            case 16: font.Scaling = on ? 150 : 100; break;
            case 17: font.Position = on ? 3 : 0; break;
            case 18: font.Kerning = on ? 14 : 1; break;
            default: font.Bold = on; break;
        }
    }

    private static void AddParagraphs(WordprocessingDocumentWrapper document, int count)
    {
        for (int i = 0; i < count; i++)
        {
            document.Content.InsertAfter("p");
            document.Content.InsertParagraphAfter();
        }
    }

    private static void SetParagraphProperty(ParagraphFormatWrapper format, int index, bool on)
    {
        switch (index)
        {
            case 0: format.Alignment = on ? ParagraphAlignment.Center : ParagraphAlignment.Left; break;
            case 1: format.Alignment = on ? ParagraphAlignment.Right : ParagraphAlignment.Left; break;
            case 2: format.Alignment = on ? ParagraphAlignment.Justify : ParagraphAlignment.Left; break;
            case 3: format.LeftIndent = on ? 18 : 0; break;
            case 4: format.RightIndent = on ? 18 : 0; break;
            case 5: format.FirstLineIndent = on ? 18 : 0; break;
            case 6: format.FirstLineIndent = on ? -18 : 0; break;
            case 7: format.SpaceBefore = on ? 6 : 0; break;
            case 8: format.SpaceAfter = on ? 6 : 8; break;
            case 9: format.LineSpacingRule = on ? LineSpacingRule.OnePointFive : LineSpacingRule.SingleSpacing; break;
            case 10: format.KeepWithNext = on; break;
            case 11: format.KeepTogether = on; break;
            case 12: format.PageBreakBefore = on; break;
            case 13: format.WidowControl = !on; break;
            case 14: format.OutlineLevel = on ? ParagraphOutlineLevel.Level1 : ParagraphOutlineLevel.BodyText; break;
            case 15: format.NoLineNumber = on; break;
            default: format.Hyphenation = !on; break;
        }
    }

    //The styles of the built-in-styles scenario in scenarios.ps1, in the same order: every BuiltInStyle a new document does not have, then the
    //built-in styles that have no BuiltInStyle value, by name.
    private static readonly int[] BuiltInStylesScenario =
    {
        -2, -3, -4, -5, -6, -7, -8, -9, -10, -20, -21, -22, -23, -24, -25, -26, -27, -28, -30, -31, -32, -33, -35, -39, -40, -41,
        -42, -43, -44, -63, -64, -65, -67, -68, -75, -76, -77, -81, -82, -86, -87, -88, -89, -91, -180, -181, -182, -261, -262, -263, -264, -265,
        -11, -12, -13, -14, -15, -16, -17, -18, -19, -29, -34, -36, -37, -38, -45, -46, -47, -48, -51, -52, -53, -54,
        -69, -70, -71, -72, -73, -74, -78, -79, -80, -83, -84, -85, -90, -95, -96, -97, -98, -99, -100, -101, -102,
        -103, -104, -105, -158,
    };

    //The table styles, applied to one table in turn.
    private static readonly int[] BuiltInTableStylesScenario = { -159, -160, -161, -162, -163, -164, -165, -166, -167, -168, -169, -170, -171, -172, -173, -174, -175, -176, -177, -178 };

    private static readonly string[] BuiltInStylesScenarioByName = { "Comment Subject", "Placeholder Text" };

    private static void ApplyEveryBuiltInStyle(WordprocessingDocumentWrapper document)
    {
        foreach (StyleWrapper style in BuiltInStylesScenario.Select(id => document.Styles[(BuiltInStyle)id])
            .Concat(BuiltInStylesScenarioByName.Select(name => document.Styles[name])))
        {
            document.Content.InsertAfter("x");
            ParagraphWrapper paragraph = document.Paragraphs[document.Paragraphs.Count - 1];
            if (style.Type == StyleType.Character)
            {
                int start = paragraph.Range.Start;
                document.Range(start, start + 1).Style = style.Name;
            }
            else
            {
                paragraph.Style = style.Name;
            }
            document.Content.InsertParagraphAfter();
        }
        TableWrapper table = document.Tables.Add(document.Paragraphs[document.Paragraphs.Count - 1].Range, 1, 1);
        foreach (int id in BuiltInTableStylesScenario)
        {
            table.Style = document.Styles[(BuiltInStyle)id].Name;
        }
        table.Style = "Table Grid";
    }

    //The scenario folders, so that a reference without a case shows up (as a failure) rather than being ignored.
    public static TheoryData<string> Names
    {
        get
        {
            TheoryData<string> data = new TheoryData<string>();
            foreach (string folder in Directory.GetDirectories(ReferenceFolder).Select(Path.GetFileName).OrderBy(n => n, StringComparer.Ordinal))
            {
                data.Add(folder);
            }
            return data;
        }
    }

    [Fact]
    public void EveryScenarioHasAWordReference()
    {
        string[] missing = Scenarios.Select(s => s.Name).Where(n => !Directory.Exists(Path.Combine(ReferenceFolder, n))).ToArray();
        Assert.True(missing.Length == 0, "No Word reference for: " + string.Join(", ", missing) + ". Run tools/WordReference/New-WordReferences.ps1.");
    }

    [Theory]
    [MemberData(nameof(Names))]
    public void LibraryWritesWhatWordWrites(string name)
    {
        Scenario scenario = Scenarios.SingleOrDefault(s => s.Name == name);
        Assert.True(scenario != null, "There is a Word reference for '" + name + "' but no case in MarkupConformanceTest.");

        List<string> failures = Compare(scenario);
        if (failures.Count > 0)
        {
            Assert.Fail(string.Format(CultureInfo.InvariantCulture, "The library's markup for '{0}' differs from Word's (- Word, + library):\n\n{1}", name,
                string.Join("\n", failures)));
        }
    }

    private static List<string> Compare(Scenario scenario)
    {
        SortedDictionary<string, string> word = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (string path in Directory.GetFiles(Path.Combine(ReferenceFolder, scenario.Name)))
        {
            word.Add(Path.GetFileName(path), File.ReadAllText(path));
        }
        SortedDictionary<string, string> library = WriteWithLibrary(scenario.Act);

        List<string> failures = new List<string>();
        foreach (Comparison comparison in scenario.Comparisons)
        {
            comparison.Compare(word, library, failures);
        }
        return failures;
    }

    private static SortedDictionary<string, string> WriteWithLibrary(Action<WordprocessingDocumentWrapper> act)
    {
        //The test assembly runs in the culture of the Word that wrote the references (see TestCulture), which the new document's defaults follow.
        using (MemoryStream stream = new MemoryStream())
        {
            using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
            {
                act(document);
            }
            stream.Position = 0;
            return MarkupNormalizer.NormalizePackage(stream);
        }
    }

    //Every part, and the same set of parts.
    private static Comparison AllParts()
    {
        return new Comparison("all parts", null, null);
    }

    private static Comparison Part(string part)
    {
        return new Comparison(part, parts => part, root => new[] { root });
    }

    //The content of the body, without the section properties of the last section (whose page size and margins are the new document's).
    private static Comparison Body()
    {
        return new Comparison("document.xml body", parts => "document.xml", root => root.Element(W + "body").Elements().Where(e => e.Name != W + "sectPr"));
    }

    //The content of the body and the section properties of the last section.
    private static Comparison WholeBody()
    {
        return new Comparison("document.xml body", parts => "document.xml", root => root.Element(W + "body").Elements());
    }

    //The content of the body and all the section properties, without the references to headers and footers, whose ids Word and the library number
    //differently because the library has no empty parts for even pages and the first page.
    private static Comparison BodyWithoutHeaderReferences()
    {
        return new Comparison("document.xml body", parts => "document.xml", root =>
        {
            root.Descendants().Where(e => e.Name == W + "headerReference" || e.Name == W + "footerReference").Remove();
            return root.Element(W + "body").Elements();
        });
    }

    //The paragraph properties and field codes of the default footer. The library leaves the result of the field for Word to work out.
    private static Comparison PageNumber()
    {
        return new Comparison("default footer", parts => FindHeaderFooter(parts, "footer", "default"),
            root => root.Descendants().Where(e => e.Name == W + "pPr" || e.Name == W + "instrText" || e.Name == W + "t"
                ? e.Name != W + "t" || !IsFieldResult(e)
                : false));
    }

    private static bool IsFieldResult(XElement text)
    {
        return text.Parent.Element(W + "rPr")?.Element(W + "noProof") != null;
    }

    //The text of the body, the hyperlink relationships and the Hyperlink character style.
    private static Comparison[] HyperlinkParts()
    {
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        return new[]
        {
            Body(),
            new Comparison("hyperlink relationships", parts => "document.xml.rels",
                root => root.Elements(packageRelationships + "Relationship").Where(e => ((string)e.Attribute("Type") ?? string.Empty).EndsWith("/hyperlink", StringComparison.Ordinal))),
            new Comparison("styles.xml Hyperlink style", parts => "styles.xml",
                root => root.Elements(W + "style").Where(e => (string)e.Attribute(W + "styleId") == "Hyperlink"))
        };
    }

    //The text of the body and the numbering definitions.
    private static Comparison[] ListParts()
    {
        return new[] { Body(), Part("numbering.xml") };
    }

    //The text of the body, the notes and the settings that name their separators.
    private static Comparison[] NoteParts()
    {
        return new[]
        {
            Body(),
            Part("footnotes.xml"),
            Part("endnotes.xml"),
            Elements("settings.xml", "footnotePr", "endnotePr", "footnote", "endnote")
        };
    }

    //The notes, the settings that say how they are numbered, and the section properties that do too.
    private static Comparison[] NoteOptionParts()
    {
        return new[]
        {
            WholeBody(),
            Part("footnotes.xml"),
            Part("endnotes.xml"),
            Elements("settings.xml", "footnotePr", "endnotePr", "footnote", "endnote", "pos", "numFmt", "numStart", "numRestart")
        };
    }

    //The text of the body, and the glossary document that keeps the placeholders of content controls.
    private static Comparison[] ControlParts()
    {
        return new[] { Body(), Part("glossary.document.xml") };
    }

    //Fills a document with the text and gives paragraphs (0-based) styles.
    private static void HeadedDocument(WordprocessingDocumentWrapper d, string text, params (int Paragraph, string Style)[] styles)
    {
        d.Content.InsertAfter(text);
        foreach ((int paragraph, string style) in styles)
        {
            d.Paragraphs[paragraph].Style = style;
        }
    }

    private static readonly string PicturePath = System.IO.Path.GetFullPath(TestFile.Picture);

    //Elements of the drawing part of the document in any namespace, by local name.
    private static Comparison DrawingElements(params string[] localNames)
    {
        return new Comparison("drawing " + string.Join(", ", localNames), parts => "document.xml",
            root => root.Descendants().Where(e => localNames.Contains(e.Name.LocalName)));
    }

    private static Comparison[] PictureParts()
    {
        return new[] { Body(), Part("document.xml.rels"), Part("content-types.xml") };
    }

    private static Comparison[] CustomXmlParts()
    {
        return new[]
        {
            Body(), Part("customXml.item1.xml"), Part("customXml.itemProps1.xml"), Part("customXml.item1.xml.rels"),
            Part("document.xml.rels"), Part("content-types.xml")
        };
    }

    private static Comparison[] CommentParts()
    {
        return new[] { Body(), Part("comments.xml"), Part("commentsExtended.xml"), Part("commentsIds.xml"), Part("commentsExtensible.xml"), Part("people.xml") };
    }

    //The styles that the scenarios add, whose ids start with Ref.
    private static Comparison RefStyles()
    {
        return new Comparison("styles.xml Ref styles", parts => "styles.xml",
            root => root.Elements(W + "style").Where(e => ((string)e.Attribute(W + "styleId") ?? string.Empty).StartsWith("Ref", StringComparison.Ordinal)));
    }

    private static Comparison Elements(string part, params string[] localNames)
    {
        return new Comparison(part + " " + string.Join(", ", localNames.Select(n => "w:" + n)), parts => part,
            root => root.Descendants().Where(e => e.Name.Namespace == W && localNames.Contains(e.Name.LocalName)));
    }

    //The header or footer part of the given type (default, even or first) of the last section, found through its reference, because Word and
    //the library number the parts differently.
    private static Comparison HeaderFooter(string kind, string type)
    {
        return new Comparison(type + " " + kind, parts => FindHeaderFooter(parts, kind, type), root => new[] { root });
    }

    private static string FindHeaderFooter(SortedDictionary<string, string> parts, string kind, string type)
    {
        if (!parts.TryGetValue("document.xml", out string document) || !parts.TryGetValue("document.xml.rels", out string relationships))
        {
            return null;
        }
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement reference = XElement.Parse(document).Element(W + "body").Element(W + "sectPr")?.Elements(W + kind + "Reference")
            .FirstOrDefault(e => (string)e.Attribute(W + "type") == type);
        string target = XElement.Parse(relationships).Elements(packageRelationships + "Relationship")
            .Where(e => reference != null && (string)e.Attribute("Id") == (string)reference.Attribute(r + "id"))
            .Select(e => (string)e.Attribute("Target"))
            .FirstOrDefault();
        return target == null ? null : MarkupNormalizer.GetPartFileName(target.TrimStart('/'));
    }

    private sealed class Scenario
    {
        public Scenario(string name, Action<WordprocessingDocumentWrapper> act, params Comparison[] comparisons)
        {
            Name = name;
            Act = act;
            Comparisons = comparisons;
        }

        public string Name { get; }

        public Action<WordprocessingDocumentWrapper> Act { get; }

        public Comparison[] Comparisons { get; }
    }

    private sealed class Comparison
    {
        private readonly string _label;
        private readonly Func<SortedDictionary<string, string>, string> _findPart;
        private readonly Func<XElement, IEnumerable<XElement>> _select;

        //findPart picks the part from the parts of one document (null when it has none); a null findPart compares every part.
        public Comparison(string label, Func<SortedDictionary<string, string>, string> findPart, Func<XElement, IEnumerable<XElement>> select)
        {
            _label = label;
            _findPart = findPart;
            _select = select;
        }

        public void Compare(SortedDictionary<string, string> word, SortedDictionary<string, string> library, List<string> failures)
        {
            if (_findPart != null)
            {
                CompareOne(_label, _findPart, _select, word, library, failures);
                return;
            }
            string[] missing = word.Keys.Except(library.Keys).ToArray();
            string[] extra = library.Keys.Except(word.Keys).ToArray();
            if (missing.Length > 0 || extra.Length > 0)
            {
                failures.Add(string.Format(CultureInfo.InvariantCulture, "Parts: the library does not write {0} and writes {1} that Word does not.\n",
                    missing.Length == 0 ? "nothing missing" : string.Join(", ", missing), extra.Length == 0 ? "nothing" : string.Join(", ", extra)));
            }
            foreach (string part in word.Keys.Intersect(library.Keys))
            {
                CompareOne(part, parts => part, root => new[] { root }, word, library, failures);
            }
        }

        private static void CompareOne(string label, Func<SortedDictionary<string, string>, string> findPart, Func<XElement, IEnumerable<XElement>> select,
            SortedDictionary<string, string> word, SortedDictionary<string, string> library, List<string> failures)
        {
            string wordPart = findPart(word);
            if (wordPart == null || !word.TryGetValue(wordPart, out string wordXml))
            {
                failures.Add(label + ": the Word reference does not have it. Regenerate the references.\n");
                return;
            }
            string libraryPart = findPart(library);
            if (libraryPart == null || !library.TryGetValue(libraryPart, out string libraryXml))
            {
                failures.Add(string.Format(CultureInfo.InvariantCulture, "{0}: the library does not write it (Word writes {1}; the library writes {2}).\n", label,
                    wordPart, string.Join(", ", library.Keys)));
                return;
            }
            string expected = MarkupNormalizer.Select(wordXml, select);
            string actual = MarkupNormalizer.Select(libraryXml, select);
            if (expected != actual)
            {
                StringBuilder builder = new StringBuilder();
                builder.Append(label).Append(wordPart == libraryPart ? string.Empty : " (" + wordPart + " in Word, " + libraryPart + " in the library)")
                    .Append(":\n").Append(TextDiff.Lines(expected, actual));
                failures.Add(builder.ToString());
            }
        }
    }
}
