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
    };

    private static readonly string[] BuiltInStylesScenarioByName = { "Comment Subject" };

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

    private static Comparison[] CommentParts()
    {
        return new[] { Body(), Part("comments.xml"), Part("commentsExtended.xml"), Part("commentsIds.xml"), Part("commentsExtensible.xml"), Part("people.xml") };
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
