using System.Globalization;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.WordReference;

/// <summary>
/// <c>dotnet run --project tools/WordReference/Tool -- normalize &lt;docx folder&gt; &lt;output folder&gt;</c> writes the normalised parts of
/// each .docx to <c>&lt;output folder&gt;/&lt;file name&gt;/&lt;part&gt;</c>, replacing what was there.
/// <c>dotnet run --project tools/WordReference/Tool -- text &lt;docx folder&gt;</c> writes the library's <c>Content.Text</c> of each .docx to
/// <c>&lt;file name&gt;.library.txt</c> beside it (none for a document the library cannot read).
/// </summary>
internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length == 3 && args[0] == "normalize")
        {
            Normalize(args[1], args[2]);
            return 0;
        }
        if (args.Length == 2 && args[0] == "text")
        {
            WriteText(args[1]);
            return 0;
        }
        Console.Error.WriteLine("Usage: normalize <docx folder> <output folder> | text <docx folder>");
        return 2;
    }

    private static void Normalize(string inputFolder, string outputFolder)
    {
        foreach (string path in Directory.GetFiles(inputFolder, "*.docx").OrderBy(p => p, StringComparer.Ordinal))
        {
            string folder = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(path));
            if (Directory.Exists(folder))
            {
                Directory.Delete(folder, true);
            }
            Directory.CreateDirectory(folder);
            SortedDictionary<string, string> parts = MarkupNormalizer.NormalizePackage(path);
            foreach (KeyValuePair<string, string> part in parts)
            {
                File.WriteAllText(Path.Combine(folder, part.Key), part.Value, new UTF8Encoding(false));
            }
            Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "{0}: {1} parts", Path.GetFileName(folder), parts.Count));
        }
    }

    private static void WriteText(string folder)
    {
        foreach (string path in Directory.GetFiles(folder, "*.docx").OrderBy(p => p, StringComparer.Ordinal))
        {
            string text;
            try
            {
                using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open(path))
                {
                    text = document.Content.Text;
                }
            }
            catch (Exception exception) when (exception is XmlException || exception is OpenXmlPackageException || exception is InvalidDataException)
            {
                //No text file: the script reports that the library could not read the document.
                Console.Error.WriteLine(string.Format(CultureInfo.InvariantCulture, "{0}: {1}", Path.GetFileName(path), exception.Message));
                continue;
            }
            File.WriteAllText(Path.ChangeExtension(path, ".library.txt"), text, new UTF8Encoding(false));
        }
    }
}
