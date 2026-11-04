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
    private const string Step4Reason = "Plan step 4 (run and paragraph content): runs, on/off properties, xml:space and paragraph edits differ from Word's.";
    private const string Step5Reason = "Plan step 5 (tables, sections, comments and package): table, section, header, comment and protection markup differs from Word's.";

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string ReferenceFolder = Path.Combine(AppContext.BaseDirectory, "TestFiles", "WordReference");

    /// <summary>The plan step a scenario waits for before the library matches Word.</summary>
    private enum PlanStep
    {
        None = 0,
        RunsAndParagraphs = 4,
        TablesSectionsCommentsPackage = 5,
    }

    private static readonly Scenario[] Scenarios = new[]
    {
        new Scenario("blank-document", PlanStep.None, d => { }, AllParts()),

        new Scenario("insert-after-tab-and-line-break", PlanStep.RunsAndParagraphs,
            d => d.Content.InsertAfter("Before\tafter the tab\vafter the break"),
            Body()),

        new Scenario("bold-on", PlanStep.RunsAndParagraphs,
            d =>
            {
                d.Content.InsertAfter("Bold text");
                d.Range(0, 4).Font.Bold = true;
            },
            Body()),

        new Scenario("bold-on-then-off", PlanStep.RunsAndParagraphs,
            d =>
            {
                d.Content.InsertAfter("Bold text");
                d.Range(0, 4).Font.Bold = true;
                d.Range(0, 4).Font.Bold = false;
            },
            Body()),

        new Scenario("font-name", PlanStep.RunsAndParagraphs,
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

        new Scenario("bookmark", PlanStep.RunsAndParagraphs,
            d =>
            {
                d.Content.InsertAfter("Bookmarked text");
                d.Bookmarks.Add("Mark", d.Range(0, 10));
            },
            Body()),

        new Scenario("date-field", PlanStep.RunsAndParagraphs, d => d.Fields.Add(d.Range(0, 0), FieldType.Date), Body()),

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

        new Scenario("section-add", PlanStep.RunsAndParagraphs,
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

        new Scenario("heading-insert-paragraph-after", PlanStep.RunsAndParagraphs,
            d =>
            {
                d.Content.InsertAfter("Heading");
                d.Paragraphs[0].Style = "Heading 1";
                d.Paragraphs[0].Range.InsertParagraphAfter();
            },
            Body()),

        new Scenario("built-in-styles", PlanStep.None, ApplyEveryBuiltInStyle, Part("styles.xml")),

        new Scenario("join-paragraphs", PlanStep.RunsAndParagraphs,
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

    public static TheoryData<string> WaitingForStep4 => ScenariosWaitingFor(PlanStep.RunsAndParagraphs);

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
    [MemberData(nameof(WaitingForStep4), Skip = Step4Reason)]
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
