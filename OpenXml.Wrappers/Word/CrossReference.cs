using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>What a cross reference points at, as <c>WdReferenceType</c> in Word (with the same values). A caption label is given as its name.</summary>
public enum CrossReferenceType
{
    /// <summary>A numbered item; not supported.</summary>
    NumberedItem = 0,

    /// <summary>A heading.</summary>
    Heading = 1,

    /// <summary>A bookmark.</summary>
    Bookmark = 2,

    /// <summary>A footnote; not supported.</summary>
    Footnote = 3,

    /// <summary>An endnote; not supported.</summary>
    Endnote = 4
}

/// <summary>What a cross reference shows, as <c>WdReferenceKind</c> in Word (with the same values). Only the kinds listed as supported are.</summary>
public enum CrossReferenceKind
{
    /// <summary>The text of the heading or bookmark.</summary>
    ContentText = -1,

    /// <summary>The label and number of a caption with the text after it.</summary>
    EntireCaption = 2,

    /// <summary>The label and number of a caption.</summary>
    OnlyLabelAndNumber = 3,

    /// <summary>The page number of the heading, bookmark or caption.</summary>
    PageNumber = 7
}

public partial class RangeWrapper
{
    private static int _referenceNumber = 244960000 + (Environment.TickCount & 0xFFFF);

    /// <summary>
    /// Inserts a cross reference to a heading or a bookmark at the end of the range, as <c>Range.InsertCrossReference</c> in Word: a hidden
    /// bookmark round the target and a <c>REF</c> field (a <c>PAGEREF</c> field for the page number) that refers to it. The page number is
    /// shown as 1 until Word works it out.
    /// </summary>
    /// <param name="referenceType">Heading or bookmark.</param>
    /// <param name="referenceKind">The text, or the page number.</param>
    /// <param name="referenceItem">The index (0-based) of the item in the list that <see cref="WordprocessingDocumentWrapper.GetCrossReferenceItems(CrossReferenceType)"/> gives.</param>
    /// <param name="insertAsHyperlink">True to make the reference a link to its target.</param>
    /// <exception cref="NotSupportedException">The type or kind is not one of those supported.</exception>
    /// <exception cref="ArgumentOutOfRangeException">There is no item at the index.</exception>
    public void InsertCrossReference(CrossReferenceType referenceType, CrossReferenceKind referenceKind, int referenceItem, bool insertAsHyperlink = true)
    {
        if (referenceType != CrossReferenceType.Heading && referenceType != CrossReferenceType.Bookmark)
        {
            throw new NotSupportedException("Only references to headings, bookmarks and captions are supported.");
        }
        CheckKind(referenceKind, false);
        if (referenceType == CrossReferenceType.Bookmark)
        {
            List<string> names = CrossReferences.BookmarkNames(AnchorElement);
            if (referenceItem < 0 || referenceItem >= names.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(referenceItem));
            }
            InsertCrossReference(names[referenceItem], referenceKind, insertAsHyperlink);
            return;
        }

        List<Paragraph> headings = CrossReferences.Headings(AnchorElement);
        if (referenceItem < 0 || referenceItem >= headings.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(referenceItem));
        }
        string bookmark = CrossReferences.HiddenBookmark(headings[referenceItem], null);
        AddReference(bookmark, referenceKind == CrossReferenceKind.PageNumber, insertAsHyperlink);
    }

    /// <summary>
    /// Inserts a cross reference to the bookmark with a name at the end of the range, as <c>Range.InsertCrossReference</c> in Word. With a
    /// reference that is not a link the field is written the short way Word writes it.
    /// </summary>
    /// <param name="bookmarkName">The name of a bookmark of the document.</param>
    /// <param name="referenceKind">The text, or the page number.</param>
    /// <param name="insertAsHyperlink">True to make the reference a link to its target.</param>
    /// <exception cref="ArgumentException">There is no such bookmark.</exception>
    public void InsertCrossReference(string bookmarkName, CrossReferenceKind referenceKind, bool insertAsHyperlink = true)
    {
        Guard.ThrowIfNull(bookmarkName);
        CheckKind(referenceKind, false);
        if (!CrossReferences.BookmarkNames(AnchorElement).Contains(bookmarkName, StringComparer.OrdinalIgnoreCase))
        {
            throw new ArgumentException("There is no bookmark with that name.", nameof(bookmarkName));
        }
        AddReference(bookmarkName, referenceKind == CrossReferenceKind.PageNumber, insertAsHyperlink);
    }

    /// <summary>
    /// Inserts a cross reference to a caption at the end of the range, as <c>Range.InsertCrossReference</c> in Word with the name of a caption
    /// label (<c>Figure</c>, <c>Table</c>) as the type.
    /// </summary>
    /// <param name="captionLabel">The label of the caption.</param>
    /// <param name="referenceKind">The whole caption, the label and number, or the page number.</param>
    /// <param name="referenceItem">The index (0-based) in the list that <see cref="WordprocessingDocumentWrapper.GetCrossReferenceItems(string)"/> gives.</param>
    /// <param name="insertAsHyperlink">True to make the reference a link to its target.</param>
    public void InsertCrossReference(string captionLabel, CrossReferenceKind referenceKind, int referenceItem, bool insertAsHyperlink = true)
    {
        Guard.ThrowIfNull(captionLabel);
        CheckKind(referenceKind, true);
        List<Paragraph> captions = CrossReferences.Captions(AnchorElement, captionLabel);
        if (referenceItem < 0 || referenceItem >= captions.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(referenceItem));
        }
        string bookmark = CrossReferences.HiddenBookmark(captions[referenceItem], referenceKind == CrossReferenceKind.OnlyLabelAndNumber ? captionLabel : null);
        AddReference(bookmark, referenceKind == CrossReferenceKind.PageNumber, insertAsHyperlink, CrossReferences.CaptionParts(captions[referenceItem], referenceKind == CrossReferenceKind.EntireCaption));
    }

    private static void CheckKind(CrossReferenceKind kind, bool caption)
    {
        bool known = kind == CrossReferenceKind.PageNumber || (caption ? kind == CrossReferenceKind.EntireCaption || kind == CrossReferenceKind.OnlyLabelAndNumber
            : kind == CrossReferenceKind.ContentText);
        if (!known)
        {
            throw new NotSupportedException("That kind of reference is not supported.");
        }
    }

    //Adds the field as Word does: a reference that is a link as a field in pieces, one that is not as a single element; the page number is
    //shown as 1 (and not proofed) until Word works it out. The parts, if given, are the pieces of the text the reference shows, which keep the
    //formatting of what they were taken from.
    private void AddReference(string bookmark, bool page, bool hyperlink, IList<KeyValuePair<string, bool>> parts = null)
    {
        string code = " " + (page ? "PAGEREF " : "REF ") + bookmark + (hyperlink ? " \\h" : string.Empty) + " ";
        RangeWrapper at = Duplicate();
        FieldChar begin = new FieldChar() { FieldCharType = FieldCharValues.Begin, Dirty = true };
        at.InsertFieldRuns(
            _ => new List<Run>
            {
                new Run(begin),
                new Run(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
            });
        FieldWrapper field = new FieldWrapper(begin);
        if (page)
        {
            field.Result.Text = "1";
            field.MarkResultNotProofed();
            return;
        }
        field.Update();
        if (parts != null && parts.Count > 0)
        {
            Run separate = ((Run)begin.Parent).NextSibling<Run>();
            while (separate != null && separate.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.Separate)
            {
                separate = separate.NextSibling<Run>();
            }
            Run end = separate.NextSibling<Run>();
            while (end != null && end.GetFirstChild<FieldChar>()?.FieldCharType?.Value != FieldCharValues.End)
            {
                Run old = end;
                end = end.NextSibling<Run>();
                old.Remove();
            }
            foreach (KeyValuePair<string, bool> part in parts)
            {
                Run piece = new Run(OpenXmlExtensions.NewText(part.Key));
                if (part.Value)
                {
                    piece.PrependChild(new RunProperties(new NoProof()));
                }
                end.InsertBeforeSelf(piece);
            }
        }
        if (!hyperlink)
        {
            field.ToSimpleFieldIfPlain();
        }
    }

    internal static string NewReferenceName()
    {
        return "_Ref" + System.Threading.Interlocked.Increment(ref _referenceNumber).ToString(CultureInfo.InvariantCulture);
    }
}

internal static class CrossReferences
{
    private static readonly Regex HeadingName = new Regex(@"^heading [1-9]$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static IEnumerable<Paragraph> Paragraphs(OpenXmlElement context)
    {
        return StoryMap.FindRoot(context).Descendants<Paragraph>();
    }

    public static List<Paragraph> Headings(OpenXmlElement context)
    {
        return Paragraphs(context).Where(p =>
        {
            string id = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            string name = id == null ? null : StyleHelper.GetStyleName(p, id);
            return name != null && HeadingName.IsMatch(name);
        }).ToList();
    }

    private static string Label(Paragraph paragraph)
    {
        string instruction = paragraph.Descendants<SimpleField>().Select(f => f.Instruction?.Value).FirstOrDefault(i => i != null && i.TrimStart().StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase))
            ?? paragraph.Descendants<FieldCode>().Select(c => c.Text).FirstOrDefault(i => i != null && i.TrimStart().StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase));
        if (instruction == null)
        {
            return null;
        }
        string[] parts = instruction.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        return parts.Length > 1 ? parts[1] : null;
    }

    public static List<Paragraph> Captions(OpenXmlElement context, string label)
    {
        return Paragraphs(context).Where(p => string.Equals(Label(p), label, StringComparison.OrdinalIgnoreCase)).ToList();
    }

    public static List<string> BookmarkNames(OpenXmlElement context)
    {
        return StoryMap.FindRoot(context).Descendants<BookmarkStart>().Select(b => b.Name?.Value).Where(n => !string.IsNullOrEmpty(n) && n[0] != '_').ToList();
    }

    //The pieces of the text of a caption (or of its label and number): the text before the number, the number, which is not proofed, and the rest.
    public static IList<KeyValuePair<string, bool>> CaptionParts(Paragraph paragraph, bool entire)
    {
        SimpleField number = paragraph.Elements<SimpleField>().FirstOrDefault(f => f.Instruction?.Value != null && f.Instruction.Value.TrimStart().StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase));
        List<KeyValuePair<string, bool>> parts = new List<KeyValuePair<string, bool>>();
        if (number == null)
        {
            return parts;
        }
        string before = string.Concat(number.ElementsBefore().OfType<Run>().SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
        string digits = string.Concat(number.Descendants<Text>().Select(t => t.Text));
        string after = string.Concat(number.ElementsAfter().OfType<Run>().SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
        if (before.Length > 0)
        {
            parts.Add(new KeyValuePair<string, bool>(before, false));
        }
        parts.Add(new KeyValuePair<string, bool>(digits, true));
        if (entire && after.Length > 0)
        {
            parts.Add(new KeyValuePair<string, bool>(after, false));
        }
        return parts;
    }

    public static string Text(Paragraph paragraph)
    {
        return new RangeWrapper(paragraph).Text.TrimEnd('\r');
    }

    //The hidden bookmark that stands for a heading's text, a whole caption or (with a label) its label and number, made if there is none yet.
    public static string HiddenBookmark(Paragraph paragraph, string labelAndNumberOf)
    {
        bool labelOnly = labelAndNumberOf != null;
        OpenXmlElement start = paragraph.ChildElements.First(e => !(e is ParagraphProperties));
        OpenXmlElement end = null;
        if (labelOnly)
        {
            end = paragraph.ChildElements.FirstOrDefault(e => e is SimpleField f && f.Instruction?.Value != null && f.Instruction.Value.TrimStart().StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase));
            if (end == null)
            {
                throw new NotSupportedException("The caption has no number to refer to.");
            }
        }
        else
        {
            end = paragraph.ChildElements.Last(e => !(e is BookmarkEnd));
        }

        string name = RangeWrapper.NewReferenceName();
        MainDocumentPartHolder.Allocate(paragraph, out int id);
        BookmarkStart open = new BookmarkStart() { Id = id.ToString(CultureInfo.InvariantCulture), Name = name };
        BookmarkEnd finish = new BookmarkEnd() { Id = id.ToString(CultureInfo.InvariantCulture) };
        //A later bookmark opens after the ones that are open already at the same place and closes before the ones that close at the same place.
        OpenXmlElement firstContent = start;
        while (firstContent is BookmarkStart && firstContent.NextSibling<OpenXmlElement>() != null)
        {
            firstContent = firstContent.NextSibling<OpenXmlElement>();
        }
        firstContent.InsertBeforeSelf(open);
        end.InsertAfterSelf(finish);
        return name;
    }
}

internal static class MainDocumentPartHolder
{
    //A bookmark id that no bookmark of the document has yet.
    public static void Allocate(OpenXmlElement context, out int id)
    {
        DocumentFormat.OpenXml.Packaging.MainDocumentPart main = StyleHelper.GetMainDocumentPart(context);
        IEnumerable<OpenXmlElement> stories = main != null ? main.GetStoryRoots().Where(r => r != null) : new[] { StoryMap.FindRoot(context) };
        id = 0;
        foreach (BookmarkStart other in stories.SelectMany(r => r.Descendants<BookmarkStart>()))
        {
            id = Math.Max(id, Convert.ToInt32(other.Id.Value, CultureInfo.InvariantCulture) + 1);
        }
    }
}

public partial class WordprocessingDocumentWrapper
{
    /// <summary>
    /// The things a cross reference of a type can point at, as <c>Document.GetCrossReferenceItems</c> in Word: the text of the headings, or the
    /// names of the bookmarks (without the hidden ones), in the order of the document.
    /// </summary>
    /// <exception cref="NotSupportedException">The type is not heading or bookmark.</exception>
    public IList<string> GetCrossReferenceItems(CrossReferenceType referenceType)
    {
        if (referenceType == CrossReferenceType.Heading)
        {
            return CrossReferences.Headings(Content.AnchorElement).Select(CrossReferences.Text).ToList();
        }
        if (referenceType == CrossReferenceType.Bookmark)
        {
            return CrossReferences.BookmarkNames(Content.AnchorElement);
        }
        throw new NotSupportedException("Only headings, bookmarks and captions are supported.");
    }

    /// <summary>The captions with a label, as <c>Document.GetCrossReferenceItems</c> in Word with the name of a label: the text of each caption.</summary>
    public IList<string> GetCrossReferenceItems(string captionLabel)
    {
        Guard.ThrowIfNull(captionLabel);
        return CrossReferences.Captions(Content.AnchorElement, captionLabel).Select(CrossReferences.Text).ToList();
    }
}
