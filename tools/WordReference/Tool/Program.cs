using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXml.Wrappers.Test.Utils;
using OpenXml.Wrappers.Word;

namespace OpenXml.Wrappers.WordReference;

/// <summary>
/// <c>dotnet run --project tools/WordReference/Tool -- normalize &lt;docx folder&gt; &lt;output folder&gt;</c> writes the normalised parts of
/// each .docx to <c>&lt;output folder&gt;/&lt;file name&gt;/&lt;part&gt;</c>, replacing what was there.
/// <c>dotnet run --project tools/WordReference/Tool -- text &lt;docx folder&gt;</c> writes the library's <c>Content.Text</c> of each .docx to
/// <c>&lt;file name&gt;.library.txt</c> beside it (none for a document the library cannot read).
/// <c>dotnet run --project tools/WordReference/Tool -- templates &lt;docx folder&gt; &lt;templates folder&gt;</c> writes the parts the library
/// starts a new document from (from <c>blank-document.docx</c>) and its built-in style definitions (from <c>built-in-styles.docx</c>) to
/// <c>OpenXml.Wrappers/Word/Templates</c>, as Word wrote them but without revision ids, document ids and proofing state.
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
        if (args.Length == 3 && args[0] == "templates")
        {
            WriteTemplates(args[1], args[2]);
            return 0;
        }
        Console.Error.WriteLine("Usage: normalize <docx folder> <output folder> | text <docx folder> | templates <docx folder> <templates folder>");
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

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";

    private static readonly string[] LeftOutProperties =
    {
        "Application", "AppVersion", "TotalTime", "Pages", "Words", "Characters", "Lines", "Paragraphs", "CharactersWithSpaces", "Company",
    };

    private static void WriteTemplates(string inputFolder, string templatesFolder)
    {
        string blank = Path.Combine(inputFolder, "blank-document.docx");
        if (File.Exists(blank))
        {
            using (ZipArchive package = ZipFile.OpenRead(blank))
            {
                WriteTemplate(package, "word/document.xml", templatesFolder, "document.xml");
                WriteTemplate(package, "word/styles.xml", templatesFolder, "styles.xml");
                WriteTemplate(package, "word/settings.xml", templatesFolder, "settings.xml");
                WriteTemplate(package, "word/fontTable.xml", templatesFolder, "fontTable.xml");
                WriteTemplate(package, "word/webSettings.xml", templatesFolder, "webSettings.xml");
                WriteTemplate(package, "word/theme/theme1.xml", templatesFolder, "theme1.xml");
                WriteTemplate(package, "docProps/app.xml", templatesFolder, "app.xml");
            }
        }
        WriteListTemplates(inputFolder, templatesFolder);
        WriteGlossaryTemplates(inputFolder, templatesFolder);
        WriteTocFormatTemplates(inputFolder, templatesFolder);
        WriteListGalleryTemplates(inputFolder, templatesFolder);
        string styles = Path.Combine(inputFolder, "built-in-styles.docx");
        if (File.Exists(styles))
        {
            using (ZipArchive package = ZipFile.OpenRead(styles))
            {
                WriteTemplate(package, "word/styles.xml", templatesFolder, "BuiltInStyles.xml", root =>
                {
                    //Only the style definitions: the document defaults and latent styles are the new document's.
                    root.Elements().Where(e => e.Name != W + "style").Remove();
                });
            }
        }
    }

    private static void WriteListGalleryTemplates(string inputFolder, string templatesFolder)
    {
        string galleries = Path.Combine(inputFolder, "application-list-galleries-apply.docx");
        if (!File.Exists(galleries))
        {
            return;
        }
        using (ZipArchive package = ZipFile.OpenRead(galleries))
        {
            WriteTemplate(package, "word/numbering.xml", templatesFolder, "ListGalleries.xml", root =>
            {
                //The definitions of the 21 templates of the list galleries, in the order of the galleries: the bullets, the numbers and the
                //outlines, seven of each. The list instances and the ids each definition was given are the document's own.
                //Word writes them in an order of its own, so they are put in the order they were applied in, which the lists say: list 1
                //uses the first, list 2 the second, and so on. The picture bullet of the last bullet template stays out (the picture is a part
                //of its own), leaving the character it falls back on.
                List<XElement> definitions = root.Elements(W + "abstractNum").ToList();
                List<XElement> ordered = root.Elements(W + "num").OrderBy(n => (int)n.Attribute(W + "numId"))
                    .Select(n => definitions.First(a => (string)a.Attribute(W + "abstractNumId") == (string)n.Element(W + "abstractNumId").Attribute(W + "val"))).ToList();
                root.Elements().Remove();
                root.Add(ordered);
                root.Descendants(W + "nsid").Remove();
                root.Descendants(W + "lvlPicBulletId").Remove();
            });
        }
    }

    private static void WriteTocFormatTemplates(string inputFolder, string templatesFolder)
    {
        //The nine styles of the table of contents that each of the formats of the Table of Contents dialog gives (TablesOfContents.Format 1 to 6).
        XElement formats = new XElement("formats", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName));
        for (int value = 1; value <= 6; value++)
        {
            string path = Path.Combine(inputFolder, "toc-format-" + value.ToString(CultureInfo.InvariantCulture) + ".docx");
            if (!File.Exists(path))
            {
                return;
            }
            using (ZipArchive package = ZipFile.OpenRead(path))
            {
                XDocument styles;
                using (Stream stream = package.GetEntry("word/styles.xml").Open())
                {
                    styles = XDocument.Load(stream);
                }
                XElement format = new XElement("format", new XAttribute("value", value));
                foreach (XElement style in styles.Root.Elements(W + "style").Where(e => ((string)e.Attribute(W + "styleId") ?? string.Empty).StartsWith("TOC", StringComparison.Ordinal)
                    && ((string)e.Attribute(W + "type")) == "paragraph"))
                {
                    XElement copy = new XElement(style);
                    copy.DescendantsAndSelf().Attributes().Where(a => a.Name.Namespace == W && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).Remove();
                    copy.Descendants().Where(e => e.Name == W + "rsid").Remove();
                    format.Add(copy);
                }
                formats.Add(format);
            }
        }
        Directory.CreateDirectory(templatesFolder);
        XmlWriterSettings settings = new XmlWriterSettings() { Encoding = new UTF8Encoding(false), Indent = false };
        using (XmlWriter writer = XmlWriter.Create(Path.Combine(templatesFolder, "TocFormats.xml"), settings))
        {
            new XDocument(formats).Save(writer);
        }
        Console.WriteLine("Template TocFormats.xml from toc-format-1 to toc-format-6");
    }

    private static void WriteGlossaryTemplates(string inputFolder, string templatesFolder)
    {
        string controls = Path.Combine(inputFolder, "content-control-placeholders.docx");
        if (!File.Exists(controls))
        {
            return;
        }
        using (ZipArchive package = ZipFile.OpenRead(controls))
        {
            //The glossary document Word keeps the placeholders of content controls in: the three default placeholders and the parts around them.
            WriteTemplate(package, "word/glossary/document.xml", templatesFolder, "GlossaryDocument.xml");
            WriteTemplate(package, "word/glossary/styles.xml", templatesFolder, "GlossaryStyles.xml");
            WriteTemplate(package, "word/glossary/settings.xml", templatesFolder, "GlossarySettings.xml");
            WriteTemplate(package, "word/glossary/webSettings.xml", templatesFolder, "GlossaryWebSettings.xml");
            WriteTemplate(package, "word/glossary/fontTable.xml", templatesFolder, "GlossaryFontTable.xml");
        }
    }

    private static void WriteListTemplates(string inputFolder, string templatesFolder)
    {
        string lists = Path.Combine(inputFolder, "list-defaults.docx");
        if (!File.Exists(lists))
        {
            return;
        }
        using (ZipArchive package = ZipFile.OpenRead(lists))
        {
            WriteTemplate(package, "word/numbering.xml", templatesFolder, "ListDefaults.xml", root =>
            {
                //Only the definitions of the default bullet, number and outline lists: the list instances and the id each definition was
                //given are the document's own.
                root.Elements().Where(e => e.Name != W + "abstractNum").Remove();
                root.Descendants(W + "nsid").Remove();
            });
        }
    }

    private static void WriteTemplate(ZipArchive package, string entryName, string templatesFolder, string fileName, Action<XElement> change = null)
    {
        XDocument document;
        using (Stream stream = package.GetEntry(entryName).Open())
        {
            document = XDocument.Load(stream);
        }
        XElement root = document.Root;

        //What changes on every save, or says which program wrote the file: revision ids, document ids, proofing state, the application.
        //Also the statistics, which would go stale because the library does not recount them, and the empty company, so that it reads as not set.
        root.DescendantsAndSelf().Attributes().Where(a => (a.Name.Namespace == W && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)) || a.Name == W14 + "paraId" || a.Name == W14 + "textId").Remove();
        root.Descendants().Where(e => e.Name == W + "rsids" || e.Name == W + "rsid" || e.Name == W + "proofState" || e.Name == W14 + "docId"
            || e.Name == W15 + "docId" || (e.Name.Namespace == ExtendedProperties && LeftOutProperties.Contains(e.Name.LocalName))).Remove();
        change?.Invoke(root);

        Directory.CreateDirectory(templatesFolder);
        XmlWriterSettings settings = new XmlWriterSettings() { Encoding = new UTF8Encoding(false), Indent = false };
        using (XmlWriter writer = XmlWriter.Create(Path.Combine(templatesFolder, fileName), settings))
        {
            document.Save(writer);
        }
        Console.WriteLine(string.Format(CultureInfo.InvariantCulture, "Template {0} from {1}", fileName, entryName));
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
