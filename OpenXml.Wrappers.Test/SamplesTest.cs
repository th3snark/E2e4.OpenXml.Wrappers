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
public class SamplesTest : IClassFixture<SamplesTest.SampleDocuments>
{
    private readonly SampleDocuments _documents;

    public SamplesTest(SampleDocuments documents)
    {
        _documents = documents;
    }

    public static TheoryData<string, string> Checks
    {
        get
        {
            TheoryData<string, string> data = new TheoryData<string, string>();
            foreach (Example example in Examples.All)
            {
                foreach (string check in DocumentIntegrity.All)
                {
                    data.Add(example.FileName, check);
                }
            }
            return data;
        }
    }

    [Fact]
    public void ThereAreThirteenBasicFiveIntermediateAndThreeAdvancedExamples()
    {
        Assert.Equal(13, Examples.All.Count(e => e.Category == "Basic"));
        Assert.Equal(5, Examples.All.Count(e => e.Category == "Intermediate"));
        Assert.Equal(3, Examples.All.Count(e => e.Category == "Advanced"));
        Assert.Equal(Examples.All.Count, Examples.All.Select(e => e.FileName).Distinct().Count());
    }

    [Theory]
    [MemberData(nameof(Checks))]
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
