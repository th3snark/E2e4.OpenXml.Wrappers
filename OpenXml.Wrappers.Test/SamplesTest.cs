using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXml.Wrappers.Samples;
using OpenXml.Wrappers.Test.Utils;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Runs every example and checks the document it writes with each <see cref="DocumentIntegrity"/> check: valid Open XML, references that
/// resolve and unique ids, which is what Word needs to open it without repairing it.
/// </summary>
/// <remarks>
/// A check that fails today because of a known library bug is listed in <see cref="KnownIssues"/> with the step of the markup plan that fixes
/// it, and is reported as skipped rather than failed. When the step lands, remove its entries.
/// </remarks>
public class SamplesTest : IClassFixture<SamplesTest.SampleDocuments>
{
    private const string Step3Reason = "Known issue, fixed by plan step 3 (new documents and styles): the styles the library refers to are not created.";

    //Example file name and check, and the plan step that fixes the problem. Comments refer to the CommentText and CommentReference styles,
    //which the library does not create.
    private static readonly Dictionary<string, int> KnownIssues = new Dictionary<string, int>(StringComparer.Ordinal)
    {
        [Key("basic-07-comment.docx", DocumentIntegrity.Styles)] = 3,
        [Key("basic-13-characters.docx", DocumentIntegrity.Styles)] = 3,
        [Key("intermediate-03-minutes.docx", DocumentIntegrity.Styles)] = 3,
        [Key("intermediate-05-editing.docx", DocumentIntegrity.Styles)] = 3,
        [Key("advanced-01-quarterly-report.docx", DocumentIntegrity.Styles)] = 3,
        [Key("advanced-03-agreement.docx", DocumentIntegrity.Styles)] = 3,
    };

    private readonly SampleDocuments _documents;

    public SamplesTest(SampleDocuments documents)
    {
        _documents = documents;
    }

    public static TheoryData<string, string> Clean => Checks(0);

    public static TheoryData<string, string> KnownForStep3 => Checks(3);

    private static TheoryData<string, string> Checks(int step)
    {
        TheoryData<string, string> data = new TheoryData<string, string>();
        foreach (Example example in Examples.All)
        {
            foreach (string check in DocumentIntegrity.All)
            {
                if ((KnownIssues.TryGetValue(Key(example.FileName, check), out int known) ? known : 0) == step)
                {
                    data.Add(example.FileName, check);
                }
            }
        }
        return data;
    }

    private static string Key(string fileName, string check)
    {
        return fileName + "/" + check;
    }

    [Fact]
    public void ThereAreThirteenBasicFiveIntermediateAndThreeAdvancedExamples()
    {
        Assert.Equal(13, Examples.All.Count(e => e.Category == "Basic"));
        Assert.Equal(5, Examples.All.Count(e => e.Category == "Intermediate"));
        Assert.Equal(3, Examples.All.Count(e => e.Category == "Advanced"));
        Assert.Equal(Examples.All.Count, Examples.All.Select(e => e.FileName).Distinct().Count());
    }

    [Fact]
    public void EveryKnownIssueNamesAnExampleAndACheck()
    {
        HashSet<string> keys = new HashSet<string>(Examples.All.SelectMany(e => DocumentIntegrity.All.Select(c => Key(e.FileName, c))), StringComparer.Ordinal);
        Assert.DoesNotContain(KnownIssues.Keys, k => !keys.Contains(k));
        Assert.All(KnownIssues.Values, step => Assert.Equal(3, step));
    }

    [Theory]
    [MemberData(nameof(Clean))]
    [MemberData(nameof(KnownForStep3), Skip = Step3Reason)]
    public void EveryExamplePassesEveryIntegrityCheck(string fileName, string check)
    {
        string path = _documents.Get(fileName);

        using (WordprocessingDocument document = WordprocessingDocument.Open(path, false))
        {
            Assert.NotNull(document.MainDocumentPart.Document.Body);
            IList<string> problems = DocumentIntegrity.Check(document, check);
            Assert.True(problems.Count == 0, string.Format(CultureInfo.InvariantCulture, "{0} fails the {1} check:\n{2}", fileName, check,
                string.Join("\n", problems.Take(15))));
        }
    }

    //Keeps KnownIssues honest: an entry whose check now passes fails here, so that it is removed and the check guards the example again.
    [Theory]
    [MemberData(nameof(KnownForStep3))]
    public void AKnownIssueStillFails(string fileName, string check)
    {
        using (WordprocessingDocument document = WordprocessingDocument.Open(_documents.Get(fileName), false))
        {
            Assert.True(DocumentIntegrity.Check(document, check).Count > 0, string.Format(CultureInfo.InvariantCulture,
                "{0} now passes the {1} check: remove it from KnownIssues.", fileName, check));
        }
    }

    /// <summary>Writes each example once, the first time a test asks for it, to a temporary folder deleted at the end.</summary>
    public sealed class SampleDocuments : IDisposable
    {
        private readonly string _folder = Path.Combine(Path.GetTempPath(), "samples-" + Path.GetRandomFileName());
        private readonly Dictionary<string, string> _paths = new Dictionary<string, string>(StringComparer.Ordinal);

        public SampleDocuments()
        {
            Directory.CreateDirectory(_folder);
        }

        public string Get(string fileName)
        {
            if (!_paths.TryGetValue(fileName, out string path))
            {
                path = Path.Combine(_folder, fileName);
                Examples.All.Single(e => e.FileName == fileName).Create(path);
                Assert.True(File.Exists(path));
                _paths.Add(fileName, path);
            }
            return path;
        }

        public void Dispose()
        {
            Directory.Delete(_folder, true);
        }
    }
}
