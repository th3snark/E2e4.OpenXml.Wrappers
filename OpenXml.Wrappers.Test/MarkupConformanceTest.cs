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
/// A scenario the library cannot match yet is skipped, under the step of the markup plan that will make it match (the last one, when several
/// are needed). When a step lands, move its scenarios to <see cref="PlanStep.None"/>. Regenerate the references with
/// <c>tools/WordReference/New-WordReferences.ps1</c>.
/// </remarks>
public class MarkupConformanceTest
{
    private const string Step5Reason = "Plan step 5 (tables, sections, comments and package): table, section, header, comment and protection markup differs from Word's.";

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string ReferenceFolder = Path.Combine(AppContext.BaseDirectory, "TestFiles", "WordReference");

    /// <summary>The plan step a scenario waits for before the library matches Word.</summary>
    private enum PlanStep
    {
        None = 0,
        TablesSectionsCommentsPackage = 5,
    }

    private static readonly Scenario[] Scenarios = new[]
    {
        new Scenario("blank-document", PlanStep.None, d => { }, AllParts()),

        new Scenario("insert-after-tab-and-line-break", PlanStep.None,
            d => d.Content.InsertAfter("Before\tafter the tab\vafter the break"),
            Body()),

        new Scenario("bold-on", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Bold text");
                d.Range(0, 4).Font.Bold = true;
            },
            Body()),

        new Scenario("bold-on-then-off", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Bold text");
                d.Range(0, 4).Font.Bold = true;
                d.Range(0, 4).Font.Bold = false;
            },
            Body()),

        new Scenario("font-name", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Arial text");
                d.Range(0, 5).Font.Name = "Arial";
            },
            Body()),

        new Scenario("table-3x3", PlanStep.TablesSectionsCommentsPackage, d => d.Tables.Add(d.Range(0, 0), 3, 3), Body()),

        new Scenario("table-3x3-word9-behavior", PlanStep.TablesSectionsCommentsPackage,
            d => d.Tables.Add(d.Range(0, 0), 3, 3, DefaultTableBehavior.Word9TableBehavior, AutoFitBehavior.Fixed),
            Body()),

        new Scenario("cell-merge", PlanStep.TablesSectionsCommentsPackage,
            d =>
            {
                TableWrapper table = d.Tables.Add(d.Range(0, 0), 3, 3);
                table.Cell(0, 0).Merge(table.Cell(0, 1));
            },
            Body()),

        new Scenario("comment-with-reply", PlanStep.TablesSectionsCommentsPackage,
            d =>
            {
                d.Content.InsertAfter("Commented text");
                CommentWrapper comment = d.Comments.Add(d.Range(0, 9), "A comment");
                CommentWrapper reply = comment.Replies.Add("A reply");

                //Word takes the author from Application.UserName, which the generator sets to Reference; the library takes the user's
                //login name.
                foreach (CommentWrapper each in new[] { comment, reply })
                {
                    each.Author = "Reference";
                    each.Initial = "R";
                }
            },
            Body(), Part("comments.xml"), Part("commentsExtended.xml"), Part("commentsIds.xml"), Part("commentsExtensible.xml"), Part("people.xml")),

        new Scenario("bookmark", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Bookmarked text");
                d.Bookmarks.Add("Mark", d.Range(0, 10));
            },
            Body()),

        new Scenario("date-field", PlanStep.None, d => d.Fields.Add(d.Range(0, 0), FieldType.Date), Body()),

        new Scenario("header-footer-text", PlanStep.TablesSectionsCommentsPackage,
            d =>
            {
                SectionWrapper section = d.Sections[0];
                section.Headers[HeaderFooterIndex.Primary].Range.Text = "Header text";
                section.Footers[HeaderFooterIndex.Primary].Range.Text = "Footer text";
            },
            Elements("document.xml", "headerReference", "footerReference"),
            HeaderFooter("header", "default"), HeaderFooter("header", "even"), HeaderFooter("header", "first"),
            HeaderFooter("footer", "default"), HeaderFooter("footer", "even"), HeaderFooter("footer", "first")),

        new Scenario("section-add", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("First section");
                d.Sections.Add();
            },
            Body()),

        new Scenario("page-borders", PlanStep.TablesSectionsCommentsPackage, d => d.Sections[0].Borders.Enable = true, Elements("document.xml", "pgBorders")),

        new Scenario("protect-read-only", PlanStep.TablesSectionsCommentsPackage,
            d => d.Protect(ProtectionType.AllowOnlyReading, "password"),
            Elements("settings.xml", "documentProtection")),

        new Scenario("heading-insert-paragraph-after", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("built-in-styles", PlanStep.None, ApplyEveryBuiltInStyle, Part("styles.xml")),

        new Scenario("font-properties-on", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter(FontWords());
                for (int i = 0; i < 20; i++)
                {
                    SetFontProperty(d.Range(4 * i, (4 * i) + 3).Font, i, true);
                }
            },
            Body()),

        new Scenario("font-properties-on-then-off", PlanStep.None,
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

        new Scenario("bold-off-in-a-bold-style", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Strong text");
                d.Range(0, 6).Style = "Strong";
                d.Range(0, 6).Font.Bold = false;
            },
            Body()),

        new Scenario("runs-merge-when-formatting-matches", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("One two three");
                d.Range(0, 3).Font.Bold = true;
                d.Range(3, 7).Font.Bold = true;
            },
            Body()),

        new Scenario("paragraph-properties-on", PlanStep.None,
            d =>
            {
                AddParagraphs(d, 17);
                for (int i = 0; i < 17; i++)
                {
                    SetParagraphProperty(d.Paragraphs[i].Range.ParagraphFormat, i, true);
                }
            },
            Body()),

        new Scenario("paragraph-properties-on-then-off", PlanStep.None,
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

        new Scenario("normal-style-applied", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Style = "Normal";
            },
            Body()),

        new Scenario("insert-paragraph-after-formatted", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Range(0, 4).Font.Bold = true;
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("insert-paragraph-before-formatted", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Range(0, 4).Font.Bold = true;
                d.Paragraphs[0].Range.InsertParagraphBefore();
            },
            Body()),

        new Scenario("paragraphs-add", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs.Add();
            },
            Body()),

        new Scenario("split-a-formatted-paragraph", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("One two");
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Range(3, 3).InsertParagraphAfter();
            },
            Body()),

        new Scenario("insert-paragraph-after-middle", PlanStep.None,
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

        new Scenario("heading-with-format-insert-paragraph-after", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Range.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("split-a-heading-at-its-end", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Range(7, 7).InsertParagraphAfter();
            },
            Body()),

        new Scenario("split-a-heading-in-the-middle", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Range(3, 3).InsertParagraphAfter();
            },
            Body()),

        new Scenario("content-text-with-paragraphs", PlanStep.None, d => d.Content.Text = "One\rTwo", Body()),

        new Scenario("insert-before-with-a-paragraph", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(0, 0).InsertBefore("A\rB");
            },
            Body()),

        new Scenario("join-formatted-paragraphs", PlanStep.None,
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

        new Scenario("delete-a-section-break", PlanStep.None,
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

        new Scenario("breaks", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Line Page Column");
                d.Range(11, 11).InsertBreak(BreakType.ColumnBreak);
                d.Range(5, 5).InsertBreak(BreakType.PageBreak);
                d.Range(0, 0).InsertBreak(BreakType.LineBreak);
            },
            Body()),

        new Scenario("page-break-at-paragraph-start", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(0, 0).InsertBreak(BreakType.PageBreak);
            },
            Body()),

        new Scenario("page-break-at-paragraph-end", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(4, 4).InsertBreak(BreakType.PageBreak);
            },
            Body()),

        new Scenario("column-break-at-paragraph-start", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(0, 0).InsertBreak(BreakType.ColumnBreak);
            },
            Body()),

        new Scenario("column-break-at-paragraph-end", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(4, 4).InsertBreak(BreakType.ColumnBreak);
            },
            Body()),

        new Scenario("line-break-at-paragraph-end", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Text");
                d.Range(4, 4).InsertBreak(BreakType.LineBreak);
            },
            Body()),

        new Scenario("delete-text-within-a-run", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("One two three");
                d.Range(3, 7).Delete();
            },
            Body()),

        new Scenario("replace-text-within-a-run", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("One two three");
                d.Range(4, 7).Text = "2";
            },
            Body()),

        //The results of these need page layout: Word has worked them out and the library leaves them for Word, so only their code is compared.
        new Scenario("field-from-code", PlanStep.None, d => d.Fields.Add(d.Range(0, 0), FieldType.Empty, "PAGE", false), Elements("document.xml", "instrText", "fldSimple")),

        new Scenario("field-page-with-merge-format", PlanStep.None, d => d.Fields.Add(d.Range(0, 0), FieldType.Page), Elements("document.xml", "instrText", "fldSimple")),

        new Scenario("field-date-without-merge-format", PlanStep.None, d => d.Fields.Add(d.Range(0, 0), FieldType.Date, string.Empty, false), Body()),

        new Scenario("field-date-from-code", PlanStep.None, d => d.Fields.Add(d.Range(0, 0), FieldType.Empty, "DATE", false), Body()),

        new Scenario("field-within-text", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Before after");
                d.Fields.Add(d.Range(7, 7), FieldType.Date);
            },
            Body()),

        new Scenario("field-within-bold-text", PlanStep.None,
            d =>
            {
                d.Content.InsertAfter("Before after");
                d.Range(0, 12).Font.Bold = true;
                d.Fields.Add(d.Range(7, 7), FieldType.Date);
            },
            Body()),

        new Scenario("join-paragraphs", PlanStep.None,
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

    public static TheoryData<string> Matching => ScenariosWaitingFor(PlanStep.None);

    public static TheoryData<string> WaitingForStep5 => ScenariosWaitingFor(PlanStep.TablesSectionsCommentsPackage);

    public static TheoryData<string> Waiting
    {
        get
        {
            TheoryData<string> data = new TheoryData<string>();
            foreach (Scenario scenario in Scenarios.Where(s => s.Step != PlanStep.None))
            {
                data.Add(scenario.Name);
            }
            return data;
        }
    }

    //The scenario folders, so that a reference without a case shows up (as a failure) rather than being ignored.
    private static TheoryData<string> ScenariosWaitingFor(PlanStep step)
    {
        TheoryData<string> data = new TheoryData<string>();
        foreach (string folder in Directory.GetDirectories(ReferenceFolder).Select(Path.GetFileName).OrderBy(n => n, StringComparer.Ordinal))
        {
            Scenario scenario = Scenarios.SingleOrDefault(s => s.Name == folder);
            if ((scenario == null ? PlanStep.None : scenario.Step) == step)
            {
                data.Add(folder);
            }
        }
        return data;
    }

    [Fact]
    public void EveryScenarioHasAWordReference()
    {
        string[] missing = Scenarios.Select(s => s.Name).Where(n => !Directory.Exists(Path.Combine(ReferenceFolder, n))).ToArray();
        Assert.True(missing.Length == 0, "No Word reference for: " + string.Join(", ", missing) + ". Run tools/WordReference/New-WordReferences.ps1.");
    }

    [Theory]
    [MemberData(nameof(Matching))]
    [MemberData(nameof(WaitingForStep5), Skip = Step5Reason)]
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

    /// <summary>
    /// Keeps the skipped cases honest: each still runs through the library and is compared, and one that now matches Word fails here so that
    /// it is moved to <see cref="PlanStep.None"/>, where <see cref="LibraryWritesWhatWordWrites"/> then guards it.
    /// </summary>
    [Theory]
    [MemberData(nameof(Waiting))]
    public void ASkippedScenarioStillDiffersFromWord(string name)
    {
        Scenario scenario = Scenarios.Single(s => s.Name == name);

        List<string> failures = Compare(scenario);

        Assert.True(failures.Count > 0, string.Format(CultureInfo.InvariantCulture,
            "The library's markup for '{0}' now matches Word's: move the scenario from {1} to PlanStep.None.", name, scenario.Step));
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
        public Scenario(string name, PlanStep step, Action<WordprocessingDocumentWrapper> act, params Comparison[] comparisons)
        {
            Name = name;
            Step = step;
            Act = act;
            Comparisons = comparisons;
        }

        public string Name { get; }

        public PlanStep Step { get; }

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
