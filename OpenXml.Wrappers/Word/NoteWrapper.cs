using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// What footnotes and endnotes have in common: the parts that hold them, their ids and the marks that refer to them. Footnotes and endnotes
/// are the same in the file but for their names, so one set of methods works on both.
/// </summary>
internal static partial class NoteOps
{
    private const string Main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static string TextStyle(bool endnote)
    {
        return endnote ? "EndnoteText" : "FootnoteText";
    }

    internal static string ReferenceStyle(bool endnote)
    {
        return endnote ? "EndnoteReference" : "FootnoteReference";
    }

    internal static OpenXmlCompositeElement GetRoot(MainDocumentPart main, bool endnote)
    {
        return endnote ? (OpenXmlCompositeElement)main.EndnotesPart?.Endnotes : main.FootnotesPart?.Footnotes;
    }

    /// <summary>The notes of the document that are not separators, in the order of the file.</summary>
    internal static List<FootnoteEndnoteType> GetNotes(MainDocumentPart main, bool endnote)
    {
        OpenXmlCompositeElement root = GetRoot(main, endnote);
        if (root == null)
        {
            return new List<FootnoteEndnoteType>();
        }
        return root.Elements<FootnoteEndnoteType>().Where(n => n.Type == null || !n.Type.HasValue || n.Type.Value == FootnoteEndnoteValues.Normal).ToList();
    }

    /// <summary>The references to notes in the main text, in the order of the text.</summary>
    internal static List<FootnoteEndnoteReferenceType> GetReferences(MainDocumentPart main, bool endnote)
    {
        Body body = main.Document?.Body;
        if (body == null)
        {
            return new List<FootnoteEndnoteReferenceType>();
        }
        return endnote ? body.Descendants<EndnoteReference>().Cast<FootnoteEndnoteReferenceType>().ToList() : body.Descendants<FootnoteReference>().Cast<FootnoteEndnoteReferenceType>().ToList();
    }

    internal static FootnoteEndnoteType FindNote(MainDocumentPart main, bool endnote, FootnoteEndnoteReferenceType reference)
    {
        return GetNotes(main, endnote).FirstOrDefault(n => n.Id?.Value == reference.Id?.Value);
    }

    internal static FootnoteEndnoteReferenceType FindReference(MainDocumentPart main, bool endnote, FootnoteEndnoteType note)
    {
        return GetReferences(main, endnote).FirstOrDefault(r => r.Id?.Value == note.Id?.Value);
    }

    private static Settings GetSettings(MainDocumentPart main)
    {
        if (main.DocumentSettingsPart == null)
        {
            main.AddNewPart<DocumentSettingsPart>();
        }
        if (main.DocumentSettingsPart.Settings == null)
        {
            main.DocumentSettingsPart.Settings = new Settings();
        }
        return main.DocumentSettingsPart.Settings;
    }

    private static Paragraph SeparatorParagraph(OpenXmlElement mark)
    {
        return new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
            new Run(mark));
    }

    /// <summary>
    /// Makes the footnotes and endnotes parts when the document has none, with the separators, and says so in the settings, as Word does for
    /// both at once.
    /// </summary>
    internal static void EnsureParts(MainDocumentPart main)
    {
        if (main.FootnotesPart != null && main.EndnotesPart != null)
        {
            return;
        }

        Settings settings = GetSettings(main);
        if (main.FootnotesPart == null)
        {
            FootnotesPart part = main.AddNewPart<FootnotesPart>();
            part.Footnotes = new Footnotes(
                new Footnote(SeparatorParagraph(new SeparatorMark())) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Footnote(SeparatorParagraph(new ContinuationSeparatorMark())) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
            if (settings.GetFirstChild<FootnoteDocumentWideProperties>() == null)
            {
                settings.AddChild(new FootnoteDocumentWideProperties(new FootnoteSpecialReference() { Id = -1 }, new FootnoteSpecialReference() { Id = 0 }));
            }
        }
        if (main.EndnotesPart == null)
        {
            EndnotesPart part = main.AddNewPart<EndnotesPart>();
            part.Endnotes = new Endnotes(
                new Endnote(SeparatorParagraph(new SeparatorMark())) { Type = FootnoteEndnoteValues.Separator, Id = -1 },
                new Endnote(SeparatorParagraph(new ContinuationSeparatorMark())) { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 });
            if (settings.GetFirstChild<EndnoteDocumentWideProperties>() == null)
            {
                settings.AddChild(new EndnoteDocumentWideProperties(new EndnoteSpecialReference() { Id = -1 }, new EndnoteSpecialReference() { Id = 0 }));
            }
        }
    }

    /// <summary>Gives the references and the notes they refer to the numbers 1, 2, 3 in the order of the text, as Word does when it saves.</summary>
    internal static void Renumber(MainDocumentPart main, bool endnote)
    {
        List<FootnoteEndnoteReferenceType> references = GetReferences(main, endnote);
        List<FootnoteEndnoteType> notes = GetNotes(main, endnote);
        List<(FootnoteEndnoteReferenceType Reference, FootnoteEndnoteType Note)> pairs = references
            .Select(r => (Reference: r, Note: notes.FirstOrDefault(n => n.Id?.Value == r.Id?.Value)))
            .Where(p => p.Note != null).ToList();

        for (int i = 0; i < pairs.Count; i++)
        {
            pairs[i].Reference.Id = i + 1;
        }
        for (int i = 0; i < pairs.Count; i++)
        {
            pairs[i].Note.Id = i + 1;
        }
        //The notes follow in the order of their marks.
        OpenXmlCompositeElement root = GetRoot(main, endnote);
        foreach ((FootnoteEndnoteReferenceType Reference, FootnoteEndnoteType Note) pair in pairs)
        {
            pair.Note.Remove();
            root.AppendChild(pair.Note);
        }
    }

    internal static Run NewReferenceRun(bool endnote, long id, string customMark)
    {
        Run run = new Run(new RunProperties(new RunStyle() { Val = ReferenceStyle(endnote) }));
        if (endnote)
        {
            EndnoteReference reference = new EndnoteReference() { Id = id };
            AddCustomMark(reference, customMark);
            run.AppendChild(reference);
        }
        else
        {
            FootnoteReference reference = new FootnoteReference() { Id = id };
            AddCustomMark(reference, customMark);
            run.AppendChild(reference);
        }
        if (!string.IsNullOrEmpty(customMark))
        {
            run.AppendChild(OpenXmlExtensions.NewText(customMark));
        }
        return run;
    }

    private static void AddCustomMark(OpenXmlElement reference, string customMark)
    {
        if (!string.IsNullOrEmpty(customMark))
        {
            //Word writes 1 where the SDK would write true.
            reference.SetAttribute(new OpenXmlAttribute("w", "customMarkFollows", Main, "1"));
        }
    }

    internal static Paragraph NewNoteParagraph(bool endnote, string customMark, string text)
    {
        Run mark = new Run(new RunProperties(new RunStyle() { Val = ReferenceStyle(endnote) }));
        if (!string.IsNullOrEmpty(customMark))
        {
            mark.AppendChild(OpenXmlExtensions.NewText(customMark));
        }
        else if (endnote)
        {
            mark.AppendChild(new EndnoteReferenceMark());
        }
        else
        {
            mark.AppendChild(new FootnoteReferenceMark());
        }
        return new Paragraph(
            new ParagraphProperties(new ParagraphStyleId() { Val = TextStyle(endnote) }),
            mark,
            new Run(OpenXmlExtensions.NewText(" " + (text ?? string.Empty))));
    }

    internal static NoteAdded Add(RangeWrapper range, bool endnote, string reference, string text)
    {
        Guard.ThrowIfNull(range);
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(range.AnchorElement);
        if (main == null || StoryMap.FindRoot(range.AnchorElement) is Footnote || StoryMap.FindRoot(range.AnchorElement) is Endnote || StoryMap.FindRoot(range.AnchorElement) is Comment)
        {
            throw new NotSupportedException("A note can be added only to the main text of a document.");
        }

        EnsureParts(main);
        BuiltInStyles.GetOrCreate(main, TextStyle(endnote));
        BuiltInStyles.GetOrCreate(main, ReferenceStyle(endnote));

        OpenXmlCompositeElement root = GetRoot(main, endnote);
        long id = root.Elements<FootnoteEndnoteType>().Select(n => n.Id?.Value ?? 0).DefaultIfEmpty(0).Max() + 1;
        FootnoteEndnoteType note = endnote ? (FootnoteEndnoteType)new Endnote() { Id = id } : new Footnote() { Id = id };
        note.AppendChild(NewNoteParagraph(endnote, reference, text));
        root.AppendChild(note);

        //The mark goes at the end of the range, which Word leaves as it was.
        RangeWrapper at = range.Duplicate();
        at.Collapse(CollapseDirection.End);
        at.InsertNoteReference(paragraph => new List<Run> { NewReferenceRun(endnote, id, reference) });

        Renumber(main, endnote);
        return new NoteAdded() { Note = note, Main = main };
    }

    internal sealed class NoteAdded
    {
        public FootnoteEndnoteType Note;
        public MainDocumentPart Main;
    }

    /// <summary>The span of the mark in the text, with the characters of a custom mark, as <c>Footnote.Reference</c> in Word.</summary>
    internal static RangeWrapper GetReferenceRange(MainDocumentPart main, bool endnote, FootnoteEndnoteType note)
    {
        FootnoteEndnoteReferenceType reference = FindReference(main, endnote, note);
        if (reference == null)
        {
            throw new InvalidOperationException("The note has no mark in the text.");
        }
        StoryMap map = StoryMap.Build(StoryMap.FindRoot(reference));
        int at = map.Atoms.FindIndex(a => ReferenceEquals(a.Element, reference));
        int start = map.Atoms[at].Start;
        int end = map.Atoms[at].End;
        //A custom mark is text in the same run.
        for (int i = at + 1; i < map.Atoms.Count && ReferenceEquals(map.Atoms[i].Element.Parent, reference.Parent); i++)
        {
            end = map.Atoms[i].End;
        }
        return RangeWrapper.FromMap(map, start, end);
    }

    /// <summary>The text of the note without its mark and the space after it, as <c>Footnote.Range</c> in Word.</summary>
    internal static RangeWrapper GetTextRange(FootnoteEndnoteType note)
    {
        StoryMap map = StoryMap.Build(note);
        int start = 0;
        Paragraph first = note.Elements<Paragraph>().FirstOrDefault();
        Run markRun = first?.Elements<Run>().FirstOrDefault();
        if (markRun != null && map.TryGetAtom(markRun.Elements().FirstOrDefault(e => !(e is RunProperties)) ?? markRun, out Atom markAtom))
        {
            int lastOfMark = map.Atoms.FindLastIndex(a => a.Element != null && ReferenceEquals(a.Element.Parent, markRun));
            start = lastOfMark >= 0 ? map.Atoms[lastOfMark].End : markAtom.End;
        }
        //The space that Word puts after the mark.
        if (map.Length > start && map.GetText(start, start + 1) == " ")
        {
            start++;
        }
        int end = map.Length;
        if (end > start && map.Atoms.Count > 0 && map.Atoms[map.Atoms.Count - 1].Kind == AtomKind.ParagraphMark)
        {
            end--;
        }
        return RangeWrapper.FromMap(map, start, Math.Max(start, end));
    }

    internal static void Delete(MainDocumentPart main, bool endnote, FootnoteEndnoteType note)
    {
        FootnoteEndnoteReferenceType reference = FindReference(main, endnote, note);
        if (reference != null)
        {
            GetReferenceRange(main, endnote, note).Delete();
        }
        note.Remove();
        Renumber(main, endnote);
    }

    /// <summary>The notes whose marks are in the range, in the order of the text.</summary>
    internal static List<FootnoteEndnoteType> GetNotesInRange(RangeWrapper range, bool endnote)
    {
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(range.AnchorElement);
        List<FootnoteEndnoteType> result = new List<FootnoteEndnoteType>();
        if (main == null || GetRoot(main, endnote) == null)
        {
            return result;
        }
        OpenXmlElement story = StoryMap.FindRoot(range.AnchorElement);
        if (!(story is Body))
        {
            return result;
        }
        StoryMap map = StoryMap.Build(story);
        range.GetStoryPositions(out int rangeStart, out int rangeEnd);
        foreach (Atom atom in map.Atoms)
        {
            bool isReference = endnote ? atom.Element is EndnoteReference : atom.Element is FootnoteReference;
            if (isReference && (rangeStart == rangeEnd ? atom.Start >= rangeStart && atom.Start <= rangeEnd : atom.Start >= rangeStart && atom.Start < rangeEnd))
            {
                FootnoteEndnoteType note = FindNote(main, endnote, (FootnoteEndnoteReferenceType)atom.Element);
                if (note != null)
                {
                    result.Add(note);
                }
            }
        }
        return result;
    }
}

/// <summary>
/// A footnote, as <c>Footnote</c> in Word. Each footnote is a story of its own here, where Word joins them into one, so positions in
/// <see cref="Range"/> count from the start of the footnote.
/// </summary>
public class FootnoteWrapper : IEquatable<FootnoteWrapper>
{
    private readonly Footnote _note;

    internal FootnoteWrapper(Footnote note)
    {
        _note = note;
    }

    private MainDocumentPart Main
    {
        get { return StyleHelper.GetMainDocumentPart(_note); }
    }

    /// <summary>The text of the footnote, as <c>Footnote.Range</c> in Word: without its mark and the space after it.</summary>
    public RangeWrapper Range
    {
        get { return NoteOps.GetTextRange(_note); }
    }

    /// <summary>The mark of the footnote in the text, as <c>Footnote.Reference</c> in Word.</summary>
    public RangeWrapper Reference
    {
        get { return NoteOps.GetReferenceRange(Main, false, _note); }
    }

    /// <summary>The 0-based place of the footnote among the footnotes of the document (Word counts from 1).</summary>
    public int Index
    {
        get { return NoteOps.GetReferences(Main, false).FindIndex(r => r.Id?.Value == _note.Id?.Value); }
    }

    /// <summary>The range the footnote's mark is in, as <c>Footnote.Parent</c> in Word for text.</summary>
    public RangeWrapper Parent
    {
        get { return new RangeWrapper(NoteOps.FindReference(Main, false, _note).Ancestors<Paragraph>().First()); }
    }

    /// <summary>Deletes the footnote and its mark, as <c>Footnote.Delete</c> in Word. The footnotes after it are numbered again.</summary>
    public void Delete()
    {
        NoteOps.Delete(Main, false, _note);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as FootnoteWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(FootnoteWrapper other)
    {
        return !(other is null) && ReferenceEquals(_note, other._note);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _note.GetHashCode();
    }

    public static bool operator ==(FootnoteWrapper left, FootnoteWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(FootnoteWrapper left, FootnoteWrapper right)
    {
        return !(left == right);
    }
}

/// <summary>
/// An endnote, as <c>Endnote</c> in Word. See <see cref="FootnoteWrapper"/> for how its range is counted.
/// </summary>
public class EndnoteWrapper : IEquatable<EndnoteWrapper>
{
    private readonly Endnote _note;

    internal EndnoteWrapper(Endnote note)
    {
        _note = note;
    }

    private MainDocumentPart Main
    {
        get { return StyleHelper.GetMainDocumentPart(_note); }
    }

    /// <summary>The text of the endnote, as <c>Endnote.Range</c> in Word: without its mark and the space after it.</summary>
    public RangeWrapper Range
    {
        get { return NoteOps.GetTextRange(_note); }
    }

    /// <summary>The mark of the endnote in the text, as <c>Endnote.Reference</c> in Word.</summary>
    public RangeWrapper Reference
    {
        get { return NoteOps.GetReferenceRange(Main, true, _note); }
    }

    /// <summary>The 0-based place of the endnote among the endnotes of the document (Word counts from 1).</summary>
    public int Index
    {
        get { return NoteOps.GetReferences(Main, true).FindIndex(r => r.Id?.Value == _note.Id?.Value); }
    }

    /// <summary>The range the endnote's mark is in, as <c>Endnote.Parent</c> in Word for text.</summary>
    public RangeWrapper Parent
    {
        get { return new RangeWrapper(NoteOps.FindReference(Main, true, _note).Ancestors<Paragraph>().First()); }
    }

    /// <summary>Deletes the endnote and its mark, as <c>Endnote.Delete</c> in Word. The endnotes after it are numbered again.</summary>
    public void Delete()
    {
        NoteOps.Delete(Main, true, _note);
    }

    /// <inheritdoc/>
    public override bool Equals(object obj)
    {
        return Equals(obj as EndnoteWrapper);
    }

    /// <inheritdoc/>
    public bool Equals(EndnoteWrapper other)
    {
        return !(other is null) && ReferenceEquals(_note, other._note);
    }

    /// <inheritdoc/>
    public override int GetHashCode()
    {
        return _note.GetHashCode();
    }

    public static bool operator ==(EndnoteWrapper left, EndnoteWrapper right)
    {
        return ReferenceEquals(left, right) || (!(left is null) && left.Equals(right));
    }

    public static bool operator !=(EndnoteWrapper left, EndnoteWrapper right)
    {
        return !(left == right);
    }
}

/// <summary>
/// The footnotes of a range or a document, as <c>Footnotes</c> in Word.
/// </summary>
public partial class FootnoteWrapperCollection : IEnumerable<FootnoteWrapper>
{
    private readonly RangeWrapper _range;

    internal FootnoteWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the footnotes are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<FootnoteWrapper> GetItems()
    {
        return NoteOps.GetNotesInRange(_range, false).Cast<Footnote>().Select(n => new FootnoteWrapper(n)).ToList();
    }

    /// <summary>The number of footnotes.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The footnote at the index (0-based), as <c>Footnotes.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no footnote at the index.</exception>
    public FootnoteWrapper this[int index]
    {
        get
        {
            List<FootnoteWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Adds a footnote, as <c>Footnotes.Add</c> in Word: its mark goes at the end of <paramref name="range"/> (the text of the range is kept).
    /// </summary>
    /// <param name="range">Where the mark goes.</param>
    /// <param name="reference">Text to use for the mark in place of the number; without it the mark is numbered by Word.</param>
    /// <param name="text">The text of the footnote.</param>
    /// <exception cref="NotSupportedException">The range is not in the main text.</exception>
    public FootnoteWrapper Add(RangeWrapper range, string reference = null, string text = null)
    {
        return new FootnoteWrapper((Footnote)NoteOps.Add(range, false, reference, text).Note);
    }

    /// <inheritdoc/>
    public IEnumerator<FootnoteWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}

/// <summary>
/// The endnotes of a range or a document, as <c>Endnotes</c> in Word.
/// </summary>
public partial class EndnoteWrapperCollection : IEnumerable<EndnoteWrapper>
{
    private readonly RangeWrapper _range;

    internal EndnoteWrapperCollection(RangeWrapper range)
    {
        _range = range;
    }

    /// <summary>The range the endnotes are in.</summary>
    public RangeWrapper Parent
    {
        get { return _range; }
    }

    private List<EndnoteWrapper> GetItems()
    {
        return NoteOps.GetNotesInRange(_range, true).Cast<Endnote>().Select(n => new EndnoteWrapper(n)).ToList();
    }

    /// <summary>The number of endnotes.</summary>
    public int Count
    {
        get { return GetItems().Count; }
    }

    /// <summary>The endnote at the index (0-based), as <c>Endnotes.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no endnote at the index.</exception>
    public EndnoteWrapper this[int index]
    {
        get
        {
            List<EndnoteWrapper> items = GetItems();
            if (index < 0 || index >= items.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return items[index];
        }
    }

    /// <summary>
    /// Adds an endnote, as <c>Endnotes.Add</c> in Word: its mark goes at the end of <paramref name="range"/> (the text of the range is kept).
    /// </summary>
    /// <param name="range">Where the mark goes.</param>
    /// <param name="reference">Text to use for the mark in place of the number; without it the mark is numbered by Word.</param>
    /// <param name="text">The text of the endnote.</param>
    /// <exception cref="NotSupportedException">The range is not in the main text.</exception>
    public EndnoteWrapper Add(RangeWrapper range, string reference = null, string text = null)
    {
        return new EndnoteWrapper((Endnote)NoteOps.Add(range, true, reference, text).Note);
    }

    /// <inheritdoc/>
    public IEnumerator<EndnoteWrapper> GetEnumerator()
    {
        return GetItems().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
