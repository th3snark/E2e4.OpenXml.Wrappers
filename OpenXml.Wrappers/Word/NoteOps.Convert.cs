using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

internal static partial class NoteOps
{
    private const long TemporaryId = 1000000;

    /// <summary>
    /// Makes the notes of one kind notes of the other kind, in the order of the text: the mark in the text, the note itself and the styles of
    /// both. Ids are given after all the notes have moved.
    /// </summary>
    private static void MoveToOtherKind(MainDocumentPart main, bool fromEndnote, long idBase, List<FootnoteEndnoteReferenceType> references, List<FootnoteEndnoteType> notes)
    {
        bool toEndnote = !fromEndnote;
        OpenXmlCompositeElement target = GetRoot(main, toEndnote);
        long next = idBase;
        foreach (FootnoteEndnoteReferenceType reference in references)
        {
            FootnoteEndnoteType note = notes.FirstOrDefault(n => n.Id?.Value == reference.Id?.Value);
            if (note == null)
            {
                continue;
            }
            long id = next++;

            //The mark in the text, in its own run, with the style of the other kind.
            FootnoteEndnoteReferenceType moved = toEndnote ? (FootnoteEndnoteReferenceType)new EndnoteReference() { Id = id } : new FootnoteReference() { Id = id };
            foreach (OpenXmlAttribute attribute in reference.GetAttributes().Where(a => a.LocalName == "customMarkFollows"))
            {
                moved.SetAttribute(attribute);
            }
            reference.InsertBeforeSelf(moved);
            reference.Remove();
            RestyleRun(moved.Parent as Run, fromEndnote, toEndnote);

            //The note, with the paragraph style and the mark of the other kind.
            FootnoteEndnoteType result = toEndnote ? (FootnoteEndnoteType)new Endnote() { Id = id } : new Footnote() { Id = id };
            foreach (OpenXmlElement child in note.ChildElements.ToList())
            {
                child.Remove();
                result.AppendChild(child);
            }
            foreach (Paragraph paragraph in result.Elements<Paragraph>())
            {
                ParagraphStyleId style = paragraph.ParagraphProperties?.ParagraphStyleId;
                if (style?.Val?.Value == TextStyle(fromEndnote))
                {
                    style.Val = TextStyle(toEndnote);
                }
                foreach (Run run in paragraph.Elements<Run>())
                {
                    RestyleRun(run, fromEndnote, toEndnote);
                    FootnoteReferenceMark footnoteMark = run.GetFirstChild<FootnoteReferenceMark>();
                    EndnoteReferenceMark endnoteMark = run.GetFirstChild<EndnoteReferenceMark>();
                    if (toEndnote && footnoteMark != null)
                    {
                        footnoteMark.InsertBeforeSelf(new EndnoteReferenceMark());
                        footnoteMark.Remove();
                    }
                    else if (!toEndnote && endnoteMark != null)
                    {
                        endnoteMark.InsertBeforeSelf(new FootnoteReferenceMark());
                        endnoteMark.Remove();
                    }
                }
            }
            note.Remove();
            target.AppendChild(result);
        }
    }

    /// <summary>The range of a separator note of the document (the line above the notes), as <c>Footnotes.Separator</c> in Word.</summary>
    internal static RangeWrapper GetSeparator(MainDocumentPart main, bool endnote, FootnoteEndnoteValues type)
    {
        EnsureParts(main);
        FootnoteEndnoteType note = GetRoot(main, endnote).Elements<FootnoteEndnoteType>().First(n => n.Type?.Value == type);
        StoryMap map = StoryMap.Build(note);
        return RangeWrapper.FromMap(map, 0, map.Length);
    }

    private static void RestyleRun(Run run, bool fromEndnote, bool toEndnote)
    {
        RunStyle style = run?.RunProperties?.RunStyle;
        if (style?.Val?.Value == ReferenceStyle(fromEndnote))
        {
            style.Val = ReferenceStyle(toEndnote);
        }
    }

    private static void PrepareKind(MainDocumentPart main, bool endnote)
    {
        EnsureParts(main);
        BuiltInStyles.GetOrCreate(main, TextStyle(endnote));
        BuiltInStyles.GetOrCreate(main, ReferenceStyle(endnote));
    }

    /// <summary>Makes all the notes of one kind notes of the other, as <c>Footnotes.Convert</c> and <c>Endnotes.Convert</c> in Word.</summary>
    internal static void Convert(MainDocumentPart main, bool fromEndnote)
    {
        if (GetNotes(main, fromEndnote).Count == 0)
        {
            return;
        }
        PrepareKind(main, !fromEndnote);
        MoveToOtherKind(main, fromEndnote, TemporaryId, GetReferences(main, fromEndnote), GetNotes(main, fromEndnote));
        Renumber(main, true);
        Renumber(main, false);
    }

    /// <summary>Makes the footnotes endnotes and the endnotes footnotes.</summary>
    internal static void Swap(MainDocumentPart main)
    {
        if (GetNotes(main, false).Count == 0 && GetNotes(main, true).Count == 0)
        {
            return;
        }
        PrepareKind(main, false);
        PrepareKind(main, true);
        //Both lists are taken before either kind moves, or the notes that arrive would move back.
        List<FootnoteEndnoteReferenceType> footnoteReferences = GetReferences(main, false);
        List<FootnoteEndnoteType> footnotes = GetNotes(main, false);
        List<FootnoteEndnoteReferenceType> endnoteReferences = GetReferences(main, true);
        List<FootnoteEndnoteType> endnotes = GetNotes(main, true);
        MoveToOtherKind(main, false, TemporaryId, footnoteReferences, footnotes);
        MoveToOtherKind(main, true, TemporaryId * 2, endnoteReferences, endnotes);
        Renumber(main, true);
        Renumber(main, false);
    }
}
