using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Reading, inserting, replacing and deleting the text of a span of characters.
/// </summary>
/// <remarks>
/// Each operation works on a map of the story, or when the span lies inside one paragraph and the operation does not add or remove
/// paragraph marks, on a map of that paragraph alone. The helpers take the map (the "scope") they work in.
/// </remarks>
public partial class RangeWrapper
{
    private string GetSpanText()
    {
        GetScoped(false, out StoryMap map, out int start, out int end);
        return map.GetText(start, end);
    }

    /// <summary>
    /// Replaces the text of the span. The new text takes the formatting of the first character that was replaced (or of the character before
    /// an insertion point). The range then covers the new text.
    /// </summary>
    private void SetSpanText(string text)
    {
        ReplaceSpanText(text);
    }

    /// <summary>
    /// Replaces the text of the span and returns the number of positions the new text takes up.
    /// </summary>
    internal int ReplaceSpanText(string text)
    {
        string newText = text ?? String.Empty;
        GetScoped(MakesParagraphs(newText), out StoryMap map, out int start, out int end);

        //Formatting is taken from the text before it goes.
        RunProperties formatting = FindFormatting(map, start, end);
        OpenXmlElement scope = map.Root;
        if (end > start)
        {
            DeleteSpanContent(map, start, end);
        }

        int length = InsertSpanText(scope, start, newText, formatting);
        MakeSpan(StoryMap.Build(scope), start, start + length);
        return length;
    }

    //Text with a carriage return starts new paragraphs, which a map of one paragraph cannot hold.
    private static bool MakesParagraphs(string text)
    {
        return text.Contains('\r', StringComparison.Ordinal);
    }

    private void InsertIntoSpan(string text, bool after)
    {
        if (String.IsNullOrEmpty(text))
        {
            return;
        }

        GetScoped(MakesParagraphs(text), out StoryMap map, out int start, out int end);
        int position = after ? end : start;
        RunProperties formatting = FindFormatting(map, position, position);
        OpenXmlElement scope = map.Root;
        int length = InsertSpanText(scope, position, text, formatting);

        //As in Word the range grows to include the text that was added to it.
        MakeSpan(StoryMap.Build(scope), start, end + length);
    }

    private void DeleteSpan()
    {
        GetScoped(false, out StoryMap map, out int start, out int end);
        OpenXmlElement scope = map.Root;
        DeleteSpanContent(map, start, end);
        MakeSpan(StoryMap.Build(scope), start, start);
    }

    #region Formatting of new text

    /// <summary>
    /// The run properties that new text at a position should have: those of the first character of the span, or if there is none, those of
    /// the character before the position (in the same paragraph), or of the one after it.
    /// </summary>
    private static RunProperties FindFormatting(StoryMap map, int start, int end)
    {
        Atom atom = null;
        int index = map.IndexAt(start);
        if (end > start)
        {
            for (int i = Math.Max(index, 0); i < map.Atoms.Count && map.Atoms[i].Start < end; i++)
            {
                if (!map.Atoms[i].IsMark)
                {
                    atom = map.Atoms[i];
                    break;
                }
            }
        }
        if (atom == null)
        {
            int before = (index < 0 ? map.Atoms.Count : index) - 1;
            if (before >= 0 && !map.Atoms[before].IsMark && ParagraphOf(map.Atoms[before]) == ParagraphAt(map, start))
            {
                atom = map.Atoms[before];
            }
        }
        if (atom == null && index >= 0 && !map.Atoms[index].IsMark)
        {
            atom = map.Atoms[index];
        }
        return atom == null ? null : ((Run)atom.Element.Parent).RunProperties;
    }

    private static Paragraph ParagraphOf(Atom atom)
    {
        return atom.Kind == AtomKind.ParagraphMark ? (Paragraph)atom.Element : atom.Kind == AtomKind.RowMark ? null : atom.Element.Ancestors<Paragraph>().First();
    }

    //The paragraph a position is in: the paragraph of the character there, or of the last character when the position is the end.
    private static Paragraph ParagraphAt(StoryMap map, int position)
    {
        int index = map.IndexAt(position);
        if (index < 0)
        {
            index = map.Atoms.Count - 1;
        }
        return index < 0 ? null : ParagraphOf(map.Atoms[index]);
    }

    #endregion

    #region Inserting

    /// <summary>
    /// Inserts text at a position. A carriage return in the text starts a new paragraph. Returns the number of positions that were added.
    /// </summary>
    private int InsertSpanText(OpenXmlElement scope, int position, string text, RunProperties formatting)
    {
        int added = 0;
        string[] lines = text.Replace("\a", String.Empty, StringComparison.Ordinal).Replace("\r\n", "\r", StringComparison.Ordinal).Split('\r');
        for (int i = 0; i < lines.Length; i++)
        {
            if (i > 0)
            {
                SplitParagraphAt(scope, position + added);
                added++;
            }
            if (lines[i].Length > 0)
            {
                added += InsertLine(scope, position + added, lines[i], formatting);
            }
        }
        return added;
    }

    private int InsertLine(OpenXmlElement scope, int position, string line, RunProperties formatting)
    {
        InsertRunsAt(scope, position, paragraph => BuildNewRunListFromText(line, paragraph), formatting);
        return line.Length;
    }

    /// <summary>
    /// Inserts the runs made by <paramref name="build"/> at a position: straight after the character before it when that is in the same
    /// paragraph, otherwise before the character after it or at the end of the paragraph. The runs take <paramref name="formatting"/> when it is given.
    /// </summary>
    internal void InsertRunsAt(OpenXmlElement scope, int position, Func<Paragraph, List<Run>> build, RunProperties formatting)
    {
        StoryMap map = StoryMap.Build(scope).EnsureBoundary(position);
        int index = map.IndexAt(position);
        Atom next = index < 0 ? null : map.Atoms[index];
        Paragraph paragraph = ParagraphAt(map, position);
        if (paragraph == null || (next != null && next.Kind == AtomKind.RowMark))
        {
            throw new InvalidOperationException("Text cannot be inserted here, after the last cell of a table row.");
        }

        List<Run> runs = build(paragraph);
        if (formatting != null)
        {
            foreach (Run run in runs)
            {
                run.RunProperties?.Remove();
                run.PrependChild((RunProperties)formatting.CloneNode(true));
            }
        }

        //The text goes straight after the character before it when that is in the same paragraph, otherwise before the character after it
        //or at the end of the paragraph.
        int previousIndex = (index < 0 ? map.Atoms.Count : index) - 1;
        Atom previous = previousIndex >= 0 && !map.Atoms[previousIndex].IsMark ? map.Atoms[previousIndex] : null;
        if (previous != null && ParagraphOf(previous) == paragraph)
        {
            OpenXmlElement anchor = previous.Element.Parent;
            for (int i = runs.Count - 1; i >= 0; i--)
            {
                anchor.InsertAfterSelf(runs[i]);
            }
        }
        else if (next != null && !next.IsMark)
        {
            OpenXmlElement anchor = next.Element.Parent;
            foreach (Run run in runs)
            {
                anchor.InsertBeforeSelf(run);
            }
        }
        else
        {
            //The paragraph has no text yet.
            foreach (Run run in runs)
            {
                paragraph.AppendChild(run);
            }
        }
    }

    /// <summary>
    /// Splits the paragraph that holds the position in two, with a paragraph mark at the position. The second paragraph has the same
    /// properties as the first. Returns the second paragraph.
    /// </summary>
    private Paragraph SplitParagraphAt(OpenXmlElement scope, int position)
    {
        StoryMap map = StoryMap.Build(scope).EnsureBoundary(position);
        int index = map.IndexAt(position);
        Paragraph paragraph = ParagraphAt(map, position);
        if (paragraph == null || (index >= 0 && map.Atoms[index].Kind == AtomKind.RowMark))
        {
            throw new InvalidOperationException("A paragraph cannot be started here, after the last cell of a table row.");
        }

        //Everything in the paragraph from the position on goes into the new paragraph.
        List<OpenXmlElement> moved = new List<OpenXmlElement>();
        Atom first = index >= 0 && !map.Atoms[index].IsMark ? map.Atoms[index] : null;
        if (first != null)
        {
            OpenXmlElement top = first.Element.Parent;
            if (top.Parent != paragraph)
            {
                throw new NotSupportedException("A paragraph cannot be split inside a hyperlink or other element that holds runs.");
            }
            moved.Add(top);
            moved.AddRange(top.ElementsAfter());
        }

        //The mark that ended the paragraph before the split ends the second paragraph afterwards, unless the split is at the very end.
        bool endMarkMoves = first != null;

        Paragraph second = new Paragraph();
        if (paragraph.ParagraphProperties != null)
        {
            ParagraphProperties properties = (ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true);
            //A section break belongs to the end of the paragraph, so it moves to the second paragraph.
            paragraph.ParagraphProperties.RemoveAllChildren<SectionProperties>();
            second.AppendChild(properties);
        }
        foreach (OpenXmlElement element in moved)
        {
            element.Remove();
            second.AppendChild(element);
        }
        paragraph.InsertAfterSelf(second);
        if (endMarkMoves)
        {
            SpanRegistry.ParagraphSplit(paragraph, second);
        }
        return second;
    }

    #endregion

    #region Markers

    /// <summary>
    /// Puts a start marker before the first text of the span and an end marker after the last (bookmark and comment markers). A collapsed
    /// span gets both at its position, and a span with no text, only paragraph marks, covers the content of its first paragraph.
    /// </summary>
    internal void PlaceMarkers(OpenXmlElement startMarker, OpenXmlElement endMarker)
    {
        GetScoped(false, out StoryMap map, out int start, out int end);
        if (end > start)
        {
            map = map.EnsureBoundary(end);
        }
        map = map.EnsureBoundary(start);

        List<Atom> text = map.AtomsIn(start, end).Where(a => !a.IsMark).ToList();
        if (text.Count > 0)
        {
            text[0].Element.Parent.InsertBeforeSelf(startMarker);
            text[text.Count - 1].Element.Parent.InsertAfterSelf(endMarker);
            return;
        }

        if (end > start)
        {
            //Only paragraph marks: the markers go round the content of the first paragraph.
            Paragraph first = (Paragraph)map.AtomsIn(start, end).First(a => a.Kind == AtomKind.ParagraphMark).Element;
            AddAroundContent(first, startMarker, endMarker);
            return;
        }

        //An insertion point: next to the character before it when that is in the same paragraph, otherwise at the start of the paragraph.
        int index = map.IndexAt(start);
        Paragraph paragraph = ParagraphAt(map, start);
        int previousIndex = (index < 0 ? map.Atoms.Count : index) - 1;
        Atom previous = previousIndex >= 0 && !map.Atoms[previousIndex].IsMark ? map.Atoms[previousIndex] : null;
        if (previous != null && ParagraphOf(previous) == paragraph)
        {
            previous.Element.Parent.InsertAfterSelf(endMarker);
            previous.Element.Parent.InsertAfterSelf(startMarker);
        }
        else if (index >= 0 && !map.Atoms[index].IsMark)
        {
            map.Atoms[index].Element.Parent.InsertBeforeSelf(endMarker);
            endMarker.InsertBeforeSelf(startMarker);
        }
        else
        {
            AddAroundContent(paragraph, startMarker, endMarker);
        }
    }

    private static void AddAroundContent(Paragraph paragraph, OpenXmlElement startMarker, OpenXmlElement endMarker)
    {
        if (paragraph.ParagraphProperties != null)
        {
            paragraph.ParagraphProperties.InsertAfterSelf(startMarker);
        }
        else
        {
            paragraph.PrependChild(startMarker);
        }
        paragraph.AppendChild(endMarker);
    }

    #endregion

    #region Deleting

    /// <summary>
    /// Deletes the characters from <paramref name="start"/> up to <paramref name="end"/> in the map. Whole paragraphs (with their marks) go,
    /// as do whole table rows and tables; a range that starts part of the way into a paragraph and takes in its paragraph mark joins that
    /// paragraph to the one after. The end of a table cell cannot be deleted, only the content. Ranges that pointed into the deleted text
    /// are moved to the place where it was.
    /// </summary>
    private static void DeleteSpanContent(StoryMap map, int start, int end)
    {
        if (end <= start)
        {
            return;
        }

        map = map.EnsureBoundary(end);
        map = map.EnsureBoundary(start);
        List<Atom> inRange = map.AtomsIn(start, end).ToList();

        //Ranges that point into the text that is going are moved to where it was, before it is removed.
        SpanRegistry.Removing(map.Root, new HashSet<OpenXmlElement>(inRange.Select(a => a.Element)), Anchor.At(map, end));

        //Tables and rows that are wholly inside the range go. This needs the end of a row in the range.
        List<OpenXmlElement> blocks = inRange.Any(a => a.Kind == AtomKind.RowMark) ? FindCoveredBlocks(map, start, end) : new List<OpenXmlElement>();
        HashSet<OpenXmlElement> removedBlocks = new HashSet<OpenXmlElement>(blocks);
        List<Atom> remaining = inRange.Where(a => removedBlocks.Count == 0 || (!a.Element.Ancestors().Any(removedBlocks.Contains) && !removedBlocks.Contains(a.Element))).ToList();

        //The paragraph marks in the range that end ordinary paragraphs (not table cells).
        List<Paragraph> markParagraphs = remaining
            .Where(a => a.Kind == AtomKind.ParagraphMark && !IsCellEnd((Paragraph)a.Element))
            .Select(a => (Paragraph)a.Element)
            .ToList();
        Dictionary<Paragraph, int> paragraphStarts = markParagraphs.ToDictionary(p => p, p => FirstPositionOf(map, p));

        //The runs that hold the characters, and everything between them.
        List<Run> charRuns = remaining.Where(a => !a.IsMark).Select(a => (Run)a.Element.Parent).Distinct().ToList();
        List<OpenXmlElement> toRemove = new List<OpenXmlElement>();
        if (charRuns.Count > 0)
        {
            List<Run> allRuns = map.Root.Descendants<Run>().ToList();
            int first = allRuns.IndexOf(charRuns.First());
            int last = allRuns.IndexOf(charRuns.Last());
            HashSet<Run> doomed = new HashSet<Run>(allRuns.Skip(first).Take(last - first + 1));
            ExtendToWholeFields(allRuns, doomed);
            toRemove.AddRange(allRuns.Where(doomed.Contains));
        }

        List<OpenXmlElement> containers = toRemove.Select(r => r.Parent).Distinct().ToList();
        foreach (OpenXmlElement block in blocks)
        {
            if (block.Parent != null)
            {
                toRemove.Add(block);
            }
        }
        DeleteElements(toRemove.Where(e => e.Parent != null).ToList());
        RemoveEmptyInlineContainers(containers);

        //Paragraphs. Those that are wholly inside the range go; one that is only partly in it is joined to the paragraph after it.
        Paragraph head = null;
        foreach (Paragraph paragraph in markParagraphs)
        {
            if (paragraph.Parent == null)
            {
                continue;
            }
            if (paragraphStarts[paragraph] >= start)
            {
                DeleteParagraph(paragraph);
            }
            else
            {
                head = paragraph;
            }
        }
        if (head != null && head.Parent != null && head.NextSibling() is Paragraph next && !IsCellEnd(head))
        {
            foreach (OpenXmlElement element in next.ChildElements.Where(e => !(e is ParagraphProperties)).ToList())
            {
                element.Remove();
                head.AppendChild(element);
            }
            next.Remove();
            SpanRegistry.ParagraphMerged(next, head);
        }
    }

    private static bool IsCellEnd(Paragraph paragraph)
    {
        return paragraph.Parent is TableCell && paragraph.Parent.LastChild == paragraph;
    }

    //The position of the first character of a paragraph.
    private static int FirstPositionOf(StoryMap map, Paragraph paragraph)
    {
        if (!map.TryGetAtom(paragraph, out Atom mark))
        {
            return 0;
        }
        int index = mark.Index;
        while (index > 0 && !map.Atoms[index - 1].IsMark && ParagraphOf(map.Atoms[index - 1]) == paragraph)
        {
            index--;
        }
        return map.Atoms[index].Start;
    }

    /// <summary>
    /// The tables, and table rows of tables that are not all covered, whose every character is inside the range.
    /// </summary>
    private static List<OpenXmlElement> FindCoveredBlocks(StoryMap map, int start, int end)
    {
        Dictionary<OpenXmlElement, (int Start, int End)> extents = new Dictionary<OpenXmlElement, (int, int)>();
        foreach (Atom atom in map.Atoms)
        {
            IEnumerable<OpenXmlElement> owners = atom.Element.Ancestors().Where(a => a is Table || a is TableRow);
            if (atom.Kind == AtomKind.RowMark)
            {
                owners = owners.Concat(new[] { atom.Element });
            }
            foreach (OpenXmlElement owner in owners)
            {
                extents[owner] = extents.TryGetValue(owner, out (int Start, int End) current)
                    ? (Math.Min(current.Start, atom.Start), Math.Max(current.End, atom.End))
                    : (atom.Start, atom.End);
            }
        }

        List<OpenXmlElement> covered = new List<OpenXmlElement>();
        foreach (KeyValuePair<OpenXmlElement, (int Start, int End)> pair in extents)
        {
            if (pair.Value.Start >= start && pair.Value.End <= end)
            {
                covered.Add(pair.Key);
            }
        }

        //Keep only the outermost: a covered table takes its rows, and a covered row its cells and nested tables, with it.
        return covered.Where(e => !e.Ancestors().Any(covered.Contains)).ToList();
    }

    /// <summary>
    /// Adds to the runs to be deleted the rest of any field that is only partly in them, so that no field is left without its begin, separator
    /// or end.
    /// </summary>
    private static void ExtendToWholeFields(List<Run> allRuns, HashSet<Run> doomed)
    {
        bool changed = true;
        while (changed)
        {
            changed = false;
            Stack<int> open = new Stack<int>();
            for (int i = 0; i < allRuns.Count; i++)
            {
                FieldChar fieldChar = allRuns[i].GetFirstChild<FieldChar>();
                if (fieldChar == null || fieldChar.FieldCharType == null)
                {
                    continue;
                }
                if (fieldChar.FieldCharType.Value == FieldCharValues.Begin)
                {
                    open.Push(i);
                }
                else if (fieldChar.FieldCharType.Value == FieldCharValues.End && open.Count > 0)
                {
                    int begin = open.Pop();
                    bool any = false;
                    bool all = true;
                    for (int j = begin; j <= i; j++)
                    {
                        if (doomed.Contains(allRuns[j]))
                        {
                            any = true;
                        }
                        else
                        {
                            all = false;
                        }
                    }
                    if (any && !all)
                    {
                        for (int j = begin; j <= i; j++)
                        {
                            changed |= doomed.Add(allRuns[j]);
                        }
                    }
                }
            }
        }
    }

    /// <summary>
    /// Removes hyperlinks, simple fields and similar elements that have lost all their runs.
    /// </summary>
    private static void RemoveEmptyInlineContainers(List<OpenXmlElement> containers)
    {
        foreach (OpenXmlElement container in containers)
        {
            OpenXmlElement current = container;
            while (current != null && current.Parent != null && (current is Hyperlink || current is SimpleField)
                && !current.Descendants<Run>().Any())
            {
                OpenXmlElement parent = current.Parent;
                current.Remove();
                current = parent;
            }
        }
    }

    #endregion
}
