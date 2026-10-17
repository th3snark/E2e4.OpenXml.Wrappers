using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OpenXml.Wrappers.Samples;
using Xunit;

namespace OpenXml.Wrappers.Test;

/// <summary>
/// Runs every example and checks that the document it writes is valid Open XML, which is what Word needs to open it without repairing it.
/// </summary>
public class SamplesTest : IDisposable
{
    private readonly string _folder = Path.Combine(Path.GetTempPath(), "samples-" + Path.GetRandomFileName());

    public SamplesTest()
    {
        Directory.CreateDirectory(_folder);
    }

    public void Dispose()
    {
        Directory.Delete(_folder, true);
        GC.SuppressFinalize(this);
    }

    public static TheoryData<string> ExampleFiles
    {
        get
        {
            TheoryData<string> data = new TheoryData<string>();
            foreach (Example example in Examples.All)
            {
                data.Add(example.FileName);
            }
            return data;
        }
    }

    [Fact]
    public void ThereAreTwelveBasicFiveIntermediateAndThreeAdvancedExamples()
    {
        Assert.Equal(12, Examples.All.Count(e => e.Category == "Basic"));
        Assert.Equal(5, Examples.All.Count(e => e.Category == "Intermediate"));
        Assert.Equal(3, Examples.All.Count(e => e.Category == "Advanced"));
        Assert.Equal(Examples.All.Count, Examples.All.Select(e => e.FileName).Distinct().Count());
    }

    [Theory]
    [MemberData(nameof(ExampleFiles))]
    public void EveryExampleWritesAValidDocument(string fileName)
    {
        Example example = Examples.All.Single(e => e.FileName == fileName);
        string path = Path.Combine(_folder, fileName);

        example.Create(path);

        Assert.True(File.Exists(path));
        using (WordprocessingDocument document = WordprocessingDocument.Open(path, false))
        {
            OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
            string[] errors = validator.Validate(document)
                .Select(e => string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0} at {1}: {2}", e.Part?.Uri, e.Path?.XPath, e.Description))
                .ToArray();
            Assert.True(errors.Length == 0, fileName + Environment.NewLine + string.Join(Environment.NewLine, errors.Take(15)));
            Assert.NotNull(document.MainDocumentPart.Document.Body);
        }
    }
}
