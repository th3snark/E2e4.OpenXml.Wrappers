using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

internal enum AtomKind
{
    Text,
    Tab,
    Break,
    ParagraphMark,
    RowMark
}

/// <summary>
/// A piece of the text of a story: the text of a <c>w:t</c>, one tab or break, or an end of paragraph or end of row marker. Every character of
/// the story belongs to exactly one atom and an atom is <see cref="Length"/> character positions long.
/// </summary>
internal sealed class Atom
{
    public AtomKind Kind;

    /// <summary>The text, tab or break element, the paragraph for a paragraph mark and the table row for a row mark.</summary>
    public OpenXmlElement Element;

    public int Index;
    public int Start;
    public int Length;

    /// <summary>What <c>Range.Text</c> shows for the atom. This is longer than <see cref="Length"/> for the marker that ends a table cell or row.</summary>
    public string Text;

    public int End
    {
        get { return Start + Length; }
    }

    public bool IsMark
    {
        get { return Kind == AtomKind.ParagraphMark || Kind == AtomKind.RowMark; }
    }
}

/// <summary>
/// The text of a story (the body, a header or footer or a comment) as a sequence of character positions, as Word numbers them: the characters
/// of the text, a position for each tab and break, and one for each end of paragraph (which is also the end of a table cell for the last
/// paragraph of a cell) and for each end of table row. Positions of things that are not text (pictures, fields codes) are not counted.
/// </summary>
/// <remarks>
/// A map is a snapshot. Build a new one after the document changes.
/// </remarks>
internal sealed class StoryMap
{
    private StoryMap(OpenXmlElement root)
        : this(root, null)
    {
    }

    //A map of the whole of the root, or when blocks is given of those blocks only (children of the root, in order).
    private StoryMap(OpenXmlElement root, IReadOnlyList<OpenXmlElement> blocks)
    {
        Root = root;
        _blocks = blocks;
        Atoms = new List<Atom>();
        if (blocks == null)
        {
            Walk(root);
        }
        else
        {
            foreach (OpenXmlElement block in blocks)
            {
                Walk(block);
            }
        }
        Length = Atoms.Count == 0 ? 0 : Atoms[Atoms.Count - 1].End;
    }

    private readonly IReadOnlyList<OpenXmlElement> _blocks;

    /// <summary>
    /// A map of just some of the blocks (consecutive paragraphs, say) of an element, so that work on a few paragraphs does not need a map of
    /// everything around them. Positions are counted from the start of the first block.
    /// </summary>
    public static StoryMap BuildWindow(OpenXmlElement parent, IReadOnlyList<OpenXmlElement> blocks)
    {
        return new StoryMap(parent, blocks);
    }

    /// <summary>
    /// True when the map is of the whole of a story, so that its end is the end of the story.
    /// </summary>
    public bool IsStory
    {
        get { return _blocks == null && ReferenceEquals(FindRoot(Root), Root); }
    }

    /// <summary>
    /// The runs in the part of the story that the map covers.
    /// </summary>
    public IEnumerable<Run> Runs()
    {
        return _blocks == null ? Root.Descendants<Run>() : _blocks.SelectMany(b => b.Descendants<Run>());
    }

    public OpenXmlElement Root { get; }

    public List<Atom> Atoms { get; }

    public int Length { get; }

    private Dictionary<OpenXmlElement, Atom> _byElement;

    /// <summary>
    /// The atom for a text, tab or break element, a paragraph (its mark) or a table row (its end), or false if the element is not in this map.
    /// </summary>
    public bool TryGetAtom(OpenXmlElement element, out Atom atom)
    {
        if (_byElement == null)
        {
            _byElement = new Dictionary<OpenXmlElement, Atom>(Atoms.Count);
            foreach (Atom each in Atoms)
            {
                _byElement[each.Element] = each;
            }
        }
        return _byElement.TryGetValue(element, out atom);
    }

    /// <summary>
    /// The element that holds the story an element is in: the body, header, footer or comment (or the top of a detached tree).
    /// </summary>
    public static OpenXmlElement FindRoot(OpenXmlElement element)
    {
        if (element is Document document)
        {
            return document.Body ?? (OpenXmlElement)document;
        }

        OpenXmlElement current = element;
        while (current != null)
        {
            if (current is Body || current is Header || current is Footer || current is Comment || current is FootnoteEndnoteType)
            {
                return current;
            }
            if (current.Parent == null)
            {
                return current;
            }
            current = current.Parent;
        }
        return element;
    }

    /// <summary>
    /// The paragraph after the given one in the document (including paragraphs in table cells), or null after the last one in the story.
    /// Paragraphs inside runs (text boxes) are not part of the story and are skipped.
    /// </summary>
    public static Paragraph NextParagraph(Paragraph paragraph, OpenXmlElement storyRoot)
    {
        OpenXmlElement current = paragraph;
        while (current != null && !ReferenceEquals(current, storyRoot))
        {
            for (OpenXmlElement sibling = current.NextSibling(); sibling != null; sibling = sibling.NextSibling())
            {
                Paragraph found = sibling as Paragraph ?? (sibling is Run ? null : sibling.Descendants<Paragraph>().FirstOrDefault());
                if (found != null)
                {
                    return found;
                }
            }
            current = current.Parent;
        }
        return null;
    }

    #region Walking without a map

    /// <summary>
    /// The paragraphs in an element, in order, produced one at a time without looking inside paragraphs: those of a body, table cell,
    /// table, header and so on (or the paragraph itself). Paragraphs in text boxes are not part of the text of a story and are not included.
    /// </summary>
    public static IEnumerable<Paragraph> ParagraphsIn(OpenXmlElement container)
    {
        if (container is Paragraph paragraph)
        {
            yield return paragraph;
            yield break;
        }
        if (container is Run || container is SectionProperties || !(container is OpenXmlCompositeElement))
        {
            yield break;
        }

        foreach (OpenXmlElement child in container.ChildElements)
        {
            if (child is Paragraph direct)
            {
                yield return direct;
                continue;
            }
            foreach (Paragraph inner in ParagraphsIn(child))
            {
                yield return inner;
            }
        }
    }

    /// <summary>
    /// One piece of the text of a story, as <see cref="Atom"/> but without the work of building a map.
    /// </summary>
    internal readonly struct AtomInfo
    {
        public AtomInfo(OpenXmlElement element, int length)
        {
            Element = element;
            Length = length;
        }

        public OpenXmlElement Element { get; }

        public int Length { get; }
    }

    /// <summary>
    /// The pieces of the text of the story, in order, produced one at a time. This sees the same text as <see cref="Build"/>, and
    /// walking part of the way costs only that part.
    /// </summary>
    public static IEnumerable<AtomInfo> Enumerate(OpenXmlElement root)
    {
        return EnumerateBlock(root);
    }

    private static IEnumerable<AtomInfo> EnumerateBlock(OpenXmlElement element)
    {
        if (element is Paragraph paragraph)
        {
            foreach (AtomInfo info in EnumerateInline(paragraph))
            {
                yield return info;
            }
            yield return new AtomInfo(paragraph, 1);
        }
        else if (element is Table table)
        {
            foreach (TableRow row in table.Elements<TableRow>())
            {
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    foreach (OpenXmlElement block in cell.ChildElements)
                    {
                        foreach (AtomInfo info in EnumerateBlock(block))
                        {
                            yield return info;
                        }
                    }
                }
                yield return new AtomInfo(row, 1);
            }
        }
        else if (element is OpenXmlCompositeElement && !(element is Run) && !(element is SectionProperties))
        {
            foreach (OpenXmlElement child in element.ChildElements)
            {
                foreach (AtomInfo info in EnumerateBlock(child))
                {
                    yield return info;
                }
            }
        }
    }

    private static IEnumerable<AtomInfo> EnumerateInline(OpenXmlElement container)
    {
        foreach (OpenXmlElement child in container.ChildElements)
        {
            if (child is Run run)
            {
                foreach (OpenXmlElement part in run.ChildElements)
                {
                    if (part is Text text)
                    {
                        int length = text.Text?.Length ?? 0;
                        if (length > 0)
                        {
                            yield return new AtomInfo(part, length);
                        }
                    }
                    else if (part is TabChar || part is Break)
                    {
                        yield return new AtomInfo(part, 1);
                    }
                }
            }
            else if (child is OpenXmlCompositeElement && !(child is Paragraph) && !(child is ParagraphProperties))
            {
                foreach (AtomInfo info in EnumerateInline(child))
                {
                    yield return info;
                }
            }
        }
    }

    /// <summary>
    /// The number of positions in the story.
    /// </summary>
    public static int CountLength(OpenXmlElement root)
    {
        int length = 0;
        foreach (AtomInfo info in Enumerate(root))
        {
            length += info.Length;
        }
        return length;
    }

    /// <summary>
    /// The position of an anchor in the story, or false if the text it is on is not in the story.
    /// </summary>
    public static bool TryPositionOf(OpenXmlElement root, Anchor anchor, out int position)
    {
        position = 0;
        if (anchor.Element == null)
        {
            position = CountLength(root);
            return true;
        }
        foreach (AtomInfo info in Enumerate(root))
        {
            if (ReferenceEquals(info.Element, anchor.Element))
            {
                position += Math.Min(anchor.Offset, info.Length);
                return true;
            }
            position += info.Length;
        }
        return false;
    }

    /// <summary>
    /// The number of positions in the story before the first piece of text in <paramref name="scope"/>, an element inside the story (a
    /// paragraph, table cell or table). It is the amount to add to positions in a map of just the scope to get positions in the story.
    /// </summary>
    public static int OffsetOf(OpenXmlElement scope)
    {
        OpenXmlElement story = FindRoot(scope);
        if (ReferenceEquals(story, scope))
        {
            return 0;
        }

        OpenXmlElement firstInScope = null;
        foreach (AtomInfo info in Enumerate(scope))
        {
            firstInScope = info.Element;
            break;
        }
        if (firstInScope == null)
        {
            return 0;
        }

        int position = 0;
        foreach (AtomInfo info in Enumerate(story))
        {
            if (ReferenceEquals(info.Element, firstInScope))
            {
                return position;
            }
            position += info.Length;
        }
        return 0;
    }

    /// <summary>
    /// The anchors for two positions in the story, found by walking to them. False if a position is beyond the end of the story.
    /// </summary>
    public static bool TryAnchorsFor(OpenXmlElement root, int start, int end, out Anchor startAnchor, out Anchor endAnchor)
    {
        startAnchor = null;
        endAnchor = null;
        int position = 0;
        List<AtomInfo> before = new List<AtomInfo>();
        foreach (AtomInfo info in Enumerate(root))
        {
            int atomEnd = position + info.Length;
            if (startAnchor == null && start < atomEnd)
            {
                startAnchor = new Anchor(info.Element, start - position);
                startAnchor.SetNeighbours(before);
            }
            if (endAnchor == null && end < atomEnd)
            {
                endAnchor = new Anchor(info.Element, end - position);
                endAnchor.SetNeighbours(before);
                break;
            }
            before.Add(info);
            if (before.Count > 3)
            {
                before.RemoveAt(0);
            }
            position = atomEnd;
        }

        //A position at the very end of the story is the end of the story.
        if (startAnchor == null && start == position)
        {
            startAnchor = new Anchor(null, 0);
        }
        if (endAnchor == null && end == position)
        {
            endAnchor = new Anchor(null, 0);
        }
        return startAnchor != null && endAnchor != null;
    }

    #endregion

    public static StoryMap Build(OpenXmlElement root)
    {
        return new StoryMap(root);
    }

    /// <summary>
    /// The index of the atom that holds the character at <paramref name="position"/>, or -1 when the position is at or beyond the end.
    /// </summary>
    public int IndexAt(int position)
    {
        int low = 0;
        int high = Atoms.Count - 1;
        while (low <= high)
        {
            int middle = (low + high) / 2;
            Atom atom = Atoms[middle];
            if (position < atom.Start)
            {
                high = middle - 1;
            }
            else if (position >= atom.End)
            {
                low = middle + 1;
            }
            else
            {
                return middle;
            }
        }
        return -1;
    }

    /// <summary>
    /// The atoms that lie wholly inside the positions from <paramref name="start"/> up to <paramref name="end"/>, in order.
    /// </summary>
    public IEnumerable<Atom> AtomsIn(int start, int end)
    {
        int index = start >= Length ? Atoms.Count : IndexAt(start);
        for (int i = Math.Max(index, 0); i < Atoms.Count && Atoms[i].Start < end; i++)
        {
            if (Atoms[i].Start >= start && Atoms[i].End <= end)
            {
                yield return Atoms[i];
            }
        }
    }

    /// <summary>
    /// The anchor for the end of this map: the end of the story if this is the map of the whole story, otherwise the first thing after the
    /// part of the story that this map covers (the map of one paragraph, for example).
    /// </summary>
    public Anchor EndAnchor()
    {
        if (IsStory || Atoms.Count == 0)
        {
            return new Anchor(null, 0);
        }

        return AnchorAfter(Atoms[Atoms.Count - 1].Element, FindRoot(Root));
    }

    /// <summary>
    /// The anchor for the first piece of text after an element (the last paragraph or table row of a part of the story), found by looking
    /// at what follows it rather than walking the story.
    /// </summary>
    public static Anchor AnchorAfter(OpenXmlElement element, OpenXmlElement story)
    {
        OpenXmlElement current = element;
        while (current != null && !ReferenceEquals(current, story))
        {
            for (OpenXmlElement sibling = current.NextSibling(); sibling != null; sibling = sibling.NextSibling())
            {
                foreach (AtomInfo info in Enumerate(sibling))
                {
                    return new Anchor(info.Element, 0);
                }
            }

            //The end of a table row follows its last cell.
            if (current is TableCell && current.Parent is TableRow row)
            {
                return new Anchor(row, 0);
            }
            current = current.Parent;
        }
        return new Anchor(null, 0);
    }

    /// <summary>
    /// The characters of the story, one for each position: the text itself, a tab, a vertical tab or form feed for a break, a carriage return
    /// for the end of a paragraph and a bell for the end of a table row.
    /// </summary>
    public char[] ToCharacters()
    {
        char[] characters = new char[Length];
        foreach (Atom atom in Atoms)
        {
            switch (atom.Kind)
            {
                case AtomKind.Text:
                    atom.Text.CopyTo(0, characters, atom.Start, atom.Length);
                    break;
                case AtomKind.ParagraphMark:
                    characters[atom.Start] = '\r';
                    break;
                case AtomKind.RowMark:
                    characters[atom.Start] = '\a';
                    break;
                default:
                    characters[atom.Start] = atom.Text[0];
                    break;
            }
        }
        return characters;
    }

    /// <summary>
    /// The text of <c>Range.Text</c> for the positions from <paramref name="start"/> up to <paramref name="end"/>.
    /// </summary>
    public string GetText(int start, int end)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        foreach (Atom atom in Atoms)
        {
            if (atom.End <= start)
            {
                continue;
            }
            if (atom.Start >= end)
            {
                break;
            }
            if (atom.Kind == AtomKind.Text)
            {
                int from = Math.Max(start, atom.Start) - atom.Start;
                int to = Math.Min(end, atom.End) - atom.Start;
                sb.Append(atom.Text, from, to - from);
            }
            else
            {
                sb.Append(atom.Text);
            }
        }
        return sb.ToString();
    }

    #region Splitting

    /// <summary>
    /// Makes sure that no run holds characters on both sides of <paramref name="position"/>, splitting a run in two if needed. Splitting a
    /// run does not change the text or the formatting of the document. The map is rebuilt when a split happens, so the atoms change but the
    /// positions do not.
    /// </summary>
    /// <returns>This map, or the new map when the document was changed.</returns>
    public StoryMap EnsureBoundary(int position)
    {
        int index = IndexAt(position);
        if (index < 0)
        {
            return this;
        }
        Atom atom = Atoms[index];
        if (atom.IsMark)
        {
            return this;
        }

        bool changed = false;
        OpenXmlElement element = atom.Element;
        if (atom.Kind == AtomKind.Text && position > atom.Start)
        {
            element = SplitText((Text)element, position - atom.Start);
            changed = true;
        }

        //The first character of a run: the element must not follow another text, tab or break in the same run.
        Run run = (Run)element.Parent;
        if (element.ElementsBefore().Any(IsAtomElement))
        {
            SplitRunBefore(run, element);
            changed = true;
        }

        return changed ? new StoryMap(Root, _blocks) : this;
    }

    private static bool IsAtomElement(OpenXmlElement element)
    {
        return element is Text || element is TabChar || element is Break;
    }

    private static Text SplitText(Text text, int offset)
    {
        string value = text.Text;
        text.Text = value.Substring(0, offset);
        text.Space = SpaceProcessingModeValues.Preserve;
        Text second = new Text(value.Substring(offset)) { Space = SpaceProcessingModeValues.Preserve };
        text.InsertAfterSelf(second);
        SpanRegistry.TextSplit(text, second, offset);
        return second;
    }

    private static Run SplitRunBefore(Run run, OpenXmlElement firstMoved)
    {
        Run second = new Run();
        if (run.RunProperties != null)
        {
            second.AppendChild(run.RunProperties.CloneNode(true));
        }
        List<OpenXmlElement> moved = new List<OpenXmlElement>() { firstMoved };
        moved.AddRange(firstMoved.ElementsAfter());
        foreach (OpenXmlElement element in moved)
        {
            element.Remove();
            second.AppendChild(element);
        }
        run.InsertAfterSelf(second);
        return second;
    }

    #endregion

    #region Building

    private void Walk(OpenXmlElement element)
    {
        if (element is Paragraph paragraph)
        {
            AddParagraph(paragraph);
        }
        else if (element is Table table)
        {
            foreach (TableRow row in table.Elements<TableRow>())
            {
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    foreach (OpenXmlElement block in cell.ChildElements)
                    {
                        Walk(block);
                    }
                }
                Add(AtomKind.RowMark, row, 1, "\r\a");
            }
        }
        else if (element is OpenXmlCompositeElement && !(element is Run) && !(element is SectionProperties))
        {
            foreach (OpenXmlElement child in element.ChildElements)
            {
                Walk(child);
            }
        }
    }

    private void AddParagraph(Paragraph paragraph)
    {
        //Runs can be nested inside other elements (hyperlinks, content controls, smart tags etc.). Runs are not searched for paragraphs, so
        //the paragraphs nested in a run (for example in a text box) are not part of this paragraph.
        AddInline(paragraph);

        //Every paragraph, including an empty one, ends with an end of paragraph marker, and the last one in a table cell also ends the cell.
        bool endsCell = paragraph.Parent is TableCell && paragraph.Parent.LastChild == paragraph;
        Add(AtomKind.ParagraphMark, paragraph, 1, endsCell ? "\r\a" : "\r");
    }

    private void AddInline(OpenXmlElement container)
    {
        foreach (OpenXmlElement child in container.ChildElements)
        {
            if (child is Run run)
            {
                AddRun(run);
            }
            else if (child is OpenXmlCompositeElement && !(child is Paragraph) && !(child is ParagraphProperties))
            {
                AddInline(child);
            }
        }
    }

    private void AddRun(Run run)
    {
        foreach (OpenXmlElement child in run.ChildElements)
        {
            if (child is Text text)
            {
                string value = text.Text ?? String.Empty;
                if (value.Length > 0)
                {
                    Add(AtomKind.Text, child, value.Length, value);
                }
            }
            else if (child is TabChar)
            {
                Add(AtomKind.Tab, child, 1, "\t");
            }
            else if (child is Break lineBreak)
            {
                //Word represents a page break as a form feed and any other break as a vertical tab.
                bool page = lineBreak.Type != null && lineBreak.Type.HasValue && lineBreak.Type.Value == BreakValues.Page;
                Add(AtomKind.Break, child, 1, page ? "\f" : "\v");
            }
        }
    }
    private void Add(AtomKind kind, OpenXmlElement element, int length, string text)
    {
        int start = Atoms.Count == 0 ? 0 : Atoms[Atoms.Count - 1].End;
        Atoms.Add(new Atom() { Kind = kind, Element = element, Index = Atoms.Count, Start = start, Length = length, Text = text });
    }

    #endregion
}
