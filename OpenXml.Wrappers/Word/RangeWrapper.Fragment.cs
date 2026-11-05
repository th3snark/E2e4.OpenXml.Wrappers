using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Copying content: the formatted text of a range, a file or XML put into a range, and a range put into a file or XML.
/// </summary>
public partial class RangeWrapper
{
    /// <summary>
    /// The text of the range with its formatting, as <c>Range.FormattedText</c> in Word. Assigning one range's to another copies the content
    /// (paragraphs, tables, fields, pictures and links, with their formatting and the styles they use) over the content of the other, from
    /// this document or from another:
    /// <code>target.FormattedText = source.FormattedText;</code>
    /// The target covers the copy afterwards. A range that ends before the end of a paragraph copies only the text in the range, and its
    /// last paragraph joins the text after the target; one that includes the paragraph mark brings the mark, and with it the paragraph's
    /// properties, which a paragraph mark carries.
    /// </summary>
    /// <remarks>
    /// Bookmarks stay with the original, as in Word. Comments are not copied (Word copies them). A table that the range only touches is copied
    /// whole. Lists are not copied to another document (the paragraphs lose their numbering there), and a section break is not copied.
    /// </remarks>
    public RangeWrapper FormattedText
    {
        get { return Duplicate(); }
        set
        {
            Guard.ThrowIfNull(value, nameof(value));
            Fragment fragment = Fragment.From(value);
            ReplaceWithFragment(fragment, fragment.HasFinalMark, true);
        }
    }

    /// <summary>
    /// The range as a package of Office Open XML in one piece (a "flat OPC" document, as Word gives): its document part with the styles,
    /// theme and settings that go with it, as <c>Range.WordOpenXML</c> in Word.
    /// </summary>
    public string WordOpenXML
    {
        get
        {
            Fragment fragment = Fragment.From(this);
            using (MemoryStream stream = new MemoryStream())
            {
                using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(stream, true, false))
                {
                    FillDocument(document, fragment);
                    return document.WordprocessingDocument.ToFlatOpcString();
                }
            }
        }
    }

    /// <summary>
    /// Replaces the range with the content of Office Open XML in one piece, as <c>Range.InsertXML</c> in Word: the XML that
    /// <see cref="WordOpenXML"/> gives. The last paragraph mark of the XML comes with it, so the text after the range starts a new paragraph.
    /// The range is collapsed at the start of what was put in.
    /// </summary>
    /// <exception cref="ArgumentException">The text is not a flat OPC document of a Word document.</exception>
    public void InsertXML(string xml)
    {
        Guard.ThrowIfNull(xml, nameof(xml));
        using (MemoryStream stream = new MemoryStream())
        {
            WordprocessingDocument package;
            try
            {
                if (System.Xml.Linq.XDocument.Parse(xml).Root?.Name != System.Xml.Linq.XName.Get("package", "http://schemas.microsoft.com/office/2006/xmlPackage"))
                {
                    throw new ArgumentException("The text is not a flat OPC document.");
                }
                package = WordprocessingDocument.FromFlatOpcString(xml, stream, true);
            }
            catch (Exception exception) when (exception is System.Xml.XmlException || exception is OpenXmlPackageException || exception is ArgumentException || exception is InvalidOperationException)
            {
                throw new ArgumentException("The text is not a flat OPC document of a Word document.", nameof(xml), exception);
            }
            using (package)
            {
                ReplaceWithFragment(Fragment.From(package), true, false);
            }
        }
    }

    /// <summary>
    /// Replaces the range with the whole content of a file (a document) including its last paragraph mark, as <c>Range.InsertFile</c> in Word.
    /// The range is collapsed at the start of what was put in.
    /// </summary>
    public void InsertFile(string fileName)
    {
        Guard.ThrowIfNull(fileName, nameof(fileName));
        using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, false))
        {
            ReplaceWithFragment(Fragment.From(package), true, false);
        }
    }

    /// <summary>
    /// Replaces the range with the content of a file (a document) without its last paragraph mark, so that the last paragraph joins the text
    /// after the range, as <c>Range.ImportFragment</c> in Word. The range is collapsed at the start of what was put in.
    /// </summary>
    public void ImportFragment(string fileName)
    {
        Guard.ThrowIfNull(fileName, nameof(fileName));
        using (WordprocessingDocument package = WordprocessingDocument.Open(fileName, false))
        {
            ReplaceWithFragment(Fragment.From(package), false, false);
        }
    }

    /// <summary>
    /// Saves the range as a document, as <c>Range.ExportFragment</c> in Word. A last paragraph that is only partly in the range becomes a
    /// whole paragraph of the file.
    /// </summary>
    /// <exception cref="NotSupportedException">The format is not <see cref="SaveFormat.XmlDocument"/>.</exception>
    public void ExportFragment(string fileName, SaveFormat format = SaveFormat.XmlDocument)
    {
        Guard.ThrowIfNull(fileName, nameof(fileName));
        if (format != SaveFormat.XmlDocument)
        {
            throw new NotSupportedException("A fragment can only be saved as a document (.docx).");
        }

        Fragment fragment = Fragment.From(this);
        using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Create(fileName, true, false))
        {
            FillDocument(document, fragment);
        }
    }

    //Makes the content of a new document the blocks of the fragment (a last paragraph that stopped short becomes a whole one).
    private static void FillDocument(WordprocessingDocumentWrapper document, Fragment fragment)
    {
        Body body = document.WordprocessingDocument.MainDocumentPart.Document.Body;
        fragment.PrepareFor(body);
        body.RemoveAllChildren<Paragraph>();
        SectionProperties section = body.Elements<SectionProperties>().FirstOrDefault();
        foreach (OpenXmlElement block in fragment.Blocks)
        {
            if (section != null)
            {
                section.InsertBeforeSelf(block);
            }
            else
            {
                body.AppendChild(block);
            }
        }
        if (fragment.IsEmpty || fragment.Blocks.Last() is Table)
        {
            Paragraph last = new Paragraph();
            if (section != null)
            {
                section.InsertBeforeSelf(last);
            }
            else
            {
                body.AppendChild(last);
            }
        }
    }

    /// <summary>
    /// Replaces the range with the content of a fragment, by the rules of Word: the text of the first paragraph joins the text before the
    /// range, taking the paragraph properties of the paragraph it came from when its mark came with it; the paragraphs in between are put in
    /// whole; and a last paragraph without its mark joins the text after the range. A range of elements becomes a span.
    /// </summary>
    /// <param name="fragment">The content to put in.</param>
    /// <param name="withFinalMark">Whether the last paragraph of the fragment comes with its mark.</param>
    /// <param name="coverInserted">True to make the range what was put in, false to collapse it at the start.</param>
    internal void ReplaceWithFragment(Fragment fragment, bool withFinalMark, bool coverInserted)
    {
        GetPositions(out StoryMap map, out int start, out int end);
        MakeSpan(map, start, end);
        if (end > start)
        {
            Delete();
        }
        if (fragment.IsEmpty)
        {
            return;
        }

        GetPositions(out map, out start, out _);
        OpenXmlElement story = StoryMap.FindRoot(AnchorElement);
        fragment.PrepareFor(story);
        int lengthBefore = map.Length;
        List<Paragraph> touched = new List<Paragraph>();

        bool lastIsParagraph = fragment.Blocks[fragment.Blocks.Count - 1] is Paragraph;
        if (fragment.Blocks.Count == 1 && lastIsParagraph && !withFinalMark)
        {
            //Only some text: it goes into the paragraph.
            Paragraph holder = ParagraphForPosition(map, start, out int local);
            List<OpenXmlElement> content = fragment.Blocks[0].ChildElements.Where(e => !(e is ParagraphProperties)).ToList();
            foreach (OpenXmlElement element in content)
            {
                element.Remove();
            }
            InsertElementsAt(holder, local, content);
            touched.Add(holder);
        }
        else
        {
            Paragraph head = ParagraphForPosition(map, start, out int local);
            Paragraph tail = SplitParagraphAt(head, local);
            touched.Add(head);
            touched.Add(tail);

            int next = 0;
            if (fragment.Blocks[0] is Paragraph first)
            {
                foreach (OpenXmlElement element in first.ChildElements.Where(e => !(e is ParagraphProperties)).ToList())
                {
                    element.Remove();
                    head.AppendChild(element);
                }

                //The mark that ends the first paragraph is the one of the paragraph it came from.
                head.ParagraphProperties?.Remove();
                if (first.ParagraphProperties != null)
                {
                    head.PrependChild(first.ParagraphProperties.CloneNode(true));
                }
                next = 1;
            }

            OpenXmlElement last = fragment.Blocks.Count > next ? fragment.Blocks[fragment.Blocks.Count - 1] : null;
            bool joinsTail = last is Paragraph && !withFinalMark;
            OpenXmlElement anchor = head;
            for (int i = next; i < fragment.Blocks.Count - (joinsTail ? 1 : 0); i++)
            {
                OpenXmlElement block = fragment.Blocks[i];
                anchor.InsertAfterSelf(block);
                anchor = block;
                if (block is Paragraph inserted)
                {
                    touched.Add(inserted);
                }
            }
            if (joinsTail)
            {
                List<OpenXmlElement> content = last.ChildElements.Where(e => !(e is ParagraphProperties)).ToList();
                OpenXmlElement before = tail.ParagraphProperties;
                foreach (OpenXmlElement element in content)
                {
                    element.Remove();
                    if (before == null)
                    {
                        tail.PrependChild(element);
                    }
                    else
                    {
                        before.InsertAfterSelf(element);
                    }
                    before = element;
                }
            }
        }

        foreach (Paragraph paragraph in touched)
        {
            paragraph.CoalesceRuns();
        }

        StoryMap changed = StoryMap.Build(story);
        int added = changed.Length - lengthBefore;
        MakeSpan(changed, start, coverInserted ? start + added : start);
    }

    //Puts elements of a paragraph (runs, links, fields) at a position in it, straight after the character before it when there is one in the
    //paragraph, else before the character after it, or at the end of the paragraph.
    private void InsertElementsAt(Paragraph paragraph, int position, List<OpenXmlElement> elements)
    {
        StoryMap map = StoryMap.Build(paragraph).EnsureBoundary(position);
        int index = map.IndexAt(position);
        Atom next = index < 0 ? null : map.Atoms[index];
        int previousIndex = (index < 0 ? map.Atoms.Count : index) - 1;
        Atom previous = previousIndex >= 0 && !map.Atoms[previousIndex].IsMark ? map.Atoms[previousIndex] : null;
        if (previous != null)
        {
            OpenXmlElement anchor = previous.Top;
            for (int i = elements.Count - 1; i >= 0; i--)
            {
                anchor.InsertAfterSelf(elements[i]);
            }
        }
        else if (next != null && !next.IsMark)
        {
            OpenXmlElement anchor = next.Top;
            foreach (OpenXmlElement element in elements)
            {
                anchor.InsertBeforeSelf(element);
            }
        }
        else
        {
            foreach (OpenXmlElement element in elements)
            {
                paragraph.AppendChild(element);
            }
        }
    }
}
