using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

internal enum AtomKind
{
    Text,
    Tab,
    Break,
    ParagraphMark,
    RowMark,

    /// <summary>The begin, separator or end of a field. One position, no text.</summary>
    FieldMark,

    /// <summary>The characters of the code of a field. A position for each character, no text.</summary>
    FieldCode,

    /// <summary>A picture, drawing, embedded object, symbol or footnote, endnote or comment reference. One position, no text.</summary>
    Object
}

/// <summary>
/// A piece of the text of a story: the text of a <c>w:t</c>, one tab or break, an end of paragraph or end of row marker, or something that
/// takes up positions without showing as text. Every position of the story belongs to exactly one atom and an atom is <see cref="Length"/>
/// positions long.
/// </summary>
internal sealed class Atom
{
    public AtomKind Kind;

    /// <summary>The text, tab, break or other element, the paragraph for a paragraph mark and the table row for a row mark.</summary>
    public OpenXmlElement Element;

    public int Index;
    public int Start;
    public int Length;

    /// <summary>What <c>Range.Text</c> shows for the atom. This is longer than <see cref="Length"/> for the marker that ends a table cell or row,
    /// and empty for the atoms that take up positions without showing.</summary>
    public string Text;

    public int End
    {
        get { return Start + Length; }
    }

    public bool IsMark
    {
        get { return Kind == AtomKind.ParagraphMark || Kind == AtomKind.RowMark; }
    }

    /// <summary>The child of the paragraph (or of an element round runs) that holds the atom: its run, or the field for a field written as one
    /// element. New runs go next to it.</summary>
    public OpenXmlElement Top
    {
        get { return Element is SimpleField || Element is Hyperlink || Element is SdtRun ? Element : Element.Parent; }
    }

    /// <summary>True for the begin, code and separator, or the end, of a field written as one element (<c>w:fldSimple</c>), which are not in a
    /// run and cannot be split.</summary>
    public bool IsSimpleFieldPart
    {
        get { return (Kind == AtomKind.FieldCode && (Element is SimpleField || Element is Hyperlink)) || (Kind == AtomKind.FieldMark && !(Element is FieldChar)); }
    }

    /// <summary>True for the atoms that take up positions but show no text: fields' begin, separator, end and code, and objects.</summary>
    public bool IsHidden
    {
        get { return Kind == AtomKind.FieldMark || Kind == AtomKind.FieldCode || Kind == AtomKind.Object; }
    }
}

/// <summary>
/// The text of a story (the body, a header or footer or a comment) as a sequence of character positions, as Word numbers them: the characters
/// of the text, a position for each tab and break, one for each end of paragraph (which is also the end of a table cell for the last
/// paragraph of a cell) and for each end of table row, one for each begin, separator and end of a field and for each character of its code,
/// and one for each picture, drawing, embedded object, symbol and footnote, endnote or comment reference.
/// </summary>
/// <remarks>
/// A map is a snapshot. Build a new one after the document changes. A field written as a single element (<c>w:fldSimple</c>) keeps its code
/// in an attribute; its begin, code and separator are one piece (with the field as its element) and its end another (with the last child of the field).
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
        foreach (OpenXmlElement block in blocks ?? new[] { root })
        {
            foreach (AtomInfo info in Enumerate(block))
            {
                int start = Atoms.Count == 0 ? 0 : Atoms[Atoms.Count - 1].End;
                Atoms.Add(new Atom() { Kind = info.Kind, Element = info.Element, Index = Atoms.Count, Start = start, Length = info.Length, Text = info.Text });
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
    /// The atom for a text, tab, break or other element, a paragraph (its mark) or a table row (its end), or false if the element is not in this map.
    /// </summary>
    public bool TryGetAtom(OpenXmlElement element, out Atom atom)
    {
        if (_byElement == null)
        {
            _byElement = new Dictionary<OpenXmlElement, Atom>(Atoms.Count, ReferenceEqualityComparer.Instance);
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

    /// <summary>
    /// A map of the blocks (paragraphs and tables) from the one that holds <paramref name="first"/> to the one that holds
    /// <paramref name="last"/>, at the level where both are in the same body or table cell. Null if there is no such run of blocks.
    /// </summary>
    public static StoryMap BuildWindowFor(Paragraph first, Paragraph last, OpenXmlElement story)
    {
        if (!ReferenceEquals(FindRoot(first), story) || !ReferenceEquals(FindRoot(last), story))
        {
            return null;
        }

        OpenXmlElement container = CommonAncestor(first, last);
        while (container is TableRow || container is Table)
        {
            container = container.Parent;
        }
        if (container == null)
        {
            return null;
        }

        OpenXmlElement firstBlock = ChildOf(container, first);
        OpenXmlElement lastBlock = ChildOf(container, last);
        if (firstBlock == null || lastBlock == null)
        {
            return null;
        }

        List<OpenXmlElement> blocks = new List<OpenXmlElement>();
        for (OpenXmlElement block = firstBlock; block != null; block = block.NextSibling())
        {
            blocks.Add(block);
            if (ReferenceEquals(block, lastBlock))
            {
                break;
            }
        }
        return ReferenceEquals(blocks[blocks.Count - 1], lastBlock) ? BuildWindow(container, blocks) : null;
    }

    //The ancestor of an element (or the element itself) that is a child of the container.
    private static OpenXmlElement ChildOf(OpenXmlElement container, OpenXmlElement element)
    {
        OpenXmlElement current = element;
        while (current != null && !ReferenceEquals(current.Parent, container))
        {
            current = current.Parent;
        }
        return current;
    }

    /// <summary>
    /// The lowest element that holds both elements (or one of them, if it holds the other).
    /// </summary>
    public static OpenXmlElement CommonAncestor(OpenXmlElement first, OpenXmlElement second)
    {
        HashSet<OpenXmlElement> ancestors = new HashSet<OpenXmlElement>(ReferenceEqualityComparer.Instance);
        for (OpenXmlElement current = first; current != null; current = current.Parent)
        {
            ancestors.Add(current);
        }
        for (OpenXmlElement current = second; current != null; current = current.Parent)
        {
            if (ancestors.Contains(current))
            {
                return current;
            }
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
        public AtomInfo(AtomKind kind, OpenXmlElement element, int length, string text)
        {
            Kind = kind;
            Element = element;
            Length = length;
            Text = text;
        }

        public AtomKind Kind { get; }

        public OpenXmlElement Element { get; }

        public int Length { get; }

        public string Text { get; }
    }

    /// <summary>
    /// The pieces of the text of the story, in order, produced one at a time. This is the one description of what takes up positions: the map
    /// is built from it, and walking part of the way costs only that part.
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

            //Every paragraph, including an empty one, ends with an end of paragraph marker, and the last one in a table cell also ends the cell.
            bool endsCell = paragraph.Parent is TableCell && paragraph.Parent.LastChild == paragraph;
            yield return new AtomInfo(AtomKind.ParagraphMark, paragraph, 1, endsCell ? "\r\a" : "\r");
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
                yield return new AtomInfo(AtomKind.RowMark, row, 1, "\r\a");
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

    //Runs can be nested inside other elements (hyperlinks, content controls, smart tags etc.). Runs are not searched for paragraphs, so the
    //paragraphs nested in a run (for example in a text box) are not part of this paragraph.
    private static IEnumerable<AtomInfo> EnumerateInline(OpenXmlElement container)
    {
        foreach (OpenXmlElement child in container.ChildElements)
        {
            if (child is Run run)
            {
                foreach (OpenXmlElement part in run.ChildElements)
                {
                    if (TryClassify(part, out AtomInfo info))
                    {
                        yield return info;
                    }
                }
            }
            else if (child is SimpleField || child is Hyperlink)
            {
                //As in Word a field has a begin, its code, a separator, its result and an end. This one keeps its code in an attribute (a
                //hyperlink is a field whose code Word makes up), so the first three are one piece with the field as its element and the end is
                //one piece with the last child of the field as its element (a field without a result has them all in the first).
                int codeLength = child is SimpleField simple ? simple.Instruction?.Value?.Length ?? 0 : HyperlinkCodeLength((Hyperlink)child);
                OpenXmlElement last = child.LastChild;
                yield return new AtomInfo(AtomKind.FieldCode, child, last == null ? codeLength + 3 : codeLength + 2, string.Empty);
                foreach (AtomInfo info in EnumerateInline(child))
                {
                    yield return info;
                }
                if (last != null)
                {
                    yield return new AtomInfo(AtomKind.FieldMark, last, 1, string.Empty);
                }
            }
            else if (child is SdtRun)
            {
                //A content control has a position at each end, as in Word.
                OpenXmlElement content = child.GetFirstChild<SdtContentRun>();
                yield return new AtomInfo(AtomKind.FieldMark, child, content == null ? 2 : 1, string.Empty);
                foreach (AtomInfo info in EnumerateInline(child))
                {
                    yield return info;
                }
                if (content != null)
                {
                    yield return new AtomInfo(AtomKind.FieldMark, content, 1, string.Empty);
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

    //The text Word makes up for the code of a hyperlink, which it counts: the address or place, the frame, the tip and the switch for a hyperlink.
    private static int HyperlinkCodeLength(Hyperlink link)
    {
        System.Text.StringBuilder code = new System.Text.StringBuilder(" HYPERLINK ");
        string id = link.Id?.Value;
        if (!string.IsNullOrEmpty(id))
        {
            OpenXmlPart part = (link.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;
            HyperlinkRelationship relationship = part?.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
            if (relationship != null)
            {
                code.Append('"').Append(relationship.Uri.OriginalString).Append("\" ");
            }
        }
        if (!string.IsNullOrEmpty(link.Anchor?.Value))
        {
            code.Append("\\l \"").Append(link.Anchor.Value).Append("\" ");
        }
        if (!string.IsNullOrEmpty(link.DocLocation?.Value))
        {
            code.Append("\\s \"").Append(link.DocLocation.Value).Append("\" ");
        }
        if (!string.IsNullOrEmpty(link.TargetFrame?.Value))
        {
            code.Append("\\t \"").Append(link.TargetFrame.Value).Append("\" ");
        }
        if (!string.IsNullOrEmpty(link.Tooltip?.Value))
        {
            code.Append("\\o \"").Append(link.Tooltip.Value).Append("\" ");
        }
        code.Append("\\h");
        return code.Length;
    }

    /// <summary>
    /// What a child of a run is in terms of positions: its text, or a tab, a break, or a field character, field code or object that takes up
    /// positions but shows no text. False for what takes up none.
    /// </summary>
    private static bool TryClassify(OpenXmlElement part, out AtomInfo info)
    {
        info = default;
        switch (part)
        {
            case Text text:
                string value = text.Text ?? string.Empty;
                if (value.Length == 0)
                {
                    return false;
                }
                info = new AtomInfo(AtomKind.Text, part, value.Length, value);
                return true;
            case DeletedText deleted:
                //Text deleted while changes were tracked still takes up positions, though it shows no text.
                int deletedLength = deleted.Text?.Length ?? 0;
                if (deletedLength == 0)
                {
                    return false;
                }
                info = new AtomInfo(AtomKind.FieldCode, part, deletedLength, string.Empty);
                return true;
            case TabChar:
                info = new AtomInfo(AtomKind.Tab, part, 1, "\t");
                return true;
            case Break lineBreak:
                //Word represents a page break as a form feed, a column break as a shift out character and any other break as a vertical tab.
                bool page = lineBreak.Type != null && lineBreak.Type.HasValue && lineBreak.Type.Value == BreakValues.Page;
                bool column = lineBreak.Type != null && lineBreak.Type.HasValue && lineBreak.Type.Value == BreakValues.Column;
                info = new AtomInfo(AtomKind.Break, part, 1, page ? "\f" : column ? "\u000e" : "\v");
                return true;
            case CarriageReturn:
                info = new AtomInfo(AtomKind.Break, part, 1, "\v");
                return true;
            case FieldChar:
                info = new AtomInfo(AtomKind.FieldMark, part, 1, string.Empty);
                return true;
            case FieldCode code:
                int codeLength = code.Text?.Length ?? 0;
                if (codeLength == 0)
                {
                    return false;
                }
                info = new AtomInfo(AtomKind.FieldCode, part, codeLength, string.Empty);
                return true;
            case Drawing:
            case Picture:
            case EmbeddedObject:
            case FootnoteReference:
            case EndnoteReference:
            case FootnoteReferenceMark:
            case EndnoteReferenceMark:
            case AnnotationReferenceMark:
            case CommentReference:
            case SymbolChar:
                info = new AtomInfo(AtomKind.Object, part, 1, string.Empty);
                return true;
            default:
                //Drawings are often wrapped in alternate content, which is one object whatever it offers.
                if (part.LocalName == "AlternateContent")
                {
                    info = new AtomInfo(AtomKind.Object, part, 1, string.Empty);
                    return true;
                }
                return false;
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

        //The neighbours of an anchor need the pieces of text around it, so the walk goes on a little past the places it is looking for.
        int reach = Anchor.Reach;
        int position = 0;
        List<AtomInfo> before = new List<AtomInfo>();
        Pending first = null;
        Pending last = null;
        foreach (AtomInfo info in Enumerate(root))
        {
            if (first != null && first.After.Count < reach)
            {
                first.After.Add(info);
            }
            if (last != null && last.After.Count < reach)
            {
                last.After.Add(info);
            }

            int atomEnd = position + info.Length;
            if (first == null && start < atomEnd)
            {
                first = new Pending(info.Element, start - position, before);
            }
            if (last == null && first != null && end < atomEnd)
            {
                last = new Pending(info.Element, end - position, before);
            }

            before.Add(info);
            if (before.Count > reach)
            {
                before.RemoveAt(0);
            }
            position = atomEnd;

            if (first != null && last != null && first.After.Count >= reach && last.After.Count >= reach)
            {
                break;
            }
        }

        //A position at the very end of the story is the end of the story, which has no neighbours after it.
        if (first == null && start == position)
        {
            first = new Pending(null, 0, before);
        }
        if (last == null && end == position)
        {
            last = new Pending(null, 0, before);
        }
        if (first == null || last == null)
        {
            return false;
        }

        startAnchor = first.ToAnchor();
        endAnchor = last.ToAnchor();
        return true;
    }

    //An anchor that is being found while the story is walked: the text it is on, and the text before and after it as far as the walk has got.
    private sealed class Pending
    {
        public Pending(OpenXmlElement element, int offset, List<AtomInfo> before)
        {
            Element = element;
            Offset = offset;
            Before = new List<AtomInfo>(before);
            After = new List<AtomInfo>();
        }

        public OpenXmlElement Element { get; }

        public int Offset { get; }

        public List<AtomInfo> Before { get; }

        public List<AtomInfo> After { get; }

        public Anchor ToAnchor()
        {
            Anchor anchor = new Anchor(Element, Offset);
            return Element == null ? anchor : anchor.WithNeighbours(Before, After);
        }
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

    //How far beyond the part of the story a map covers the neighbours of an anchor are looked for: the nearest few, which are what matter
    //for small removals.
    private const int MinimumNeighbours = 8;

    /// <summary>
    /// The pieces of text before the atom at <paramref name="index"/>, nearest first, up to <paramref name="wanted"/> of them. A map of part of
    /// the story (a paragraph, say) is extended with the paragraphs before it until there are at least the nearest few.
    /// </summary>
    public List<OpenXmlElement> ElementsBefore(int index, int wanted, bool extend = true)
    {
        List<OpenXmlElement> list = new List<OpenXmlElement>();
        for (int i = index - 1; i >= 0 && list.Count < wanted; i--)
        {
            list.Add(Atoms[i].Element);
        }
        if (extend && list.Count < MinimumNeighbours && index - list.Count <= 0 && !IsStory && Atoms.Count > 0)
        {
            OpenXmlElement story = FindRoot(Root);
            Paragraph paragraph = ParagraphOfFirst();
            for (int hops = 0; paragraph != null && list.Count < MinimumNeighbours && hops < MinimumNeighbours; hops++)
            {
                paragraph = PreviousParagraph(paragraph, story);
                if (paragraph == null)
                {
                    break;
                }
                List<AtomInfo> earlier = Enumerate(paragraph).ToList();
                for (int i = earlier.Count - 1; i >= 0 && list.Count < wanted; i--)
                {
                    list.Add(earlier[i].Element);
                }
            }
        }
        return list;
    }

    /// <summary>
    /// The pieces of text after the atom at <paramref name="index"/>, nearest first, up to <paramref name="wanted"/> of them, extended beyond
    /// a map of part of the story in the same way as <see cref="ElementsBefore"/>.
    /// </summary>
    public List<OpenXmlElement> ElementsAfter(int index, int wanted, bool extend = true)
    {
        List<OpenXmlElement> list = new List<OpenXmlElement>();
        for (int i = index + 1; i < Atoms.Count && list.Count < wanted; i++)
        {
            list.Add(Atoms[i].Element);
        }
        if (extend && list.Count < MinimumNeighbours && index + list.Count >= Atoms.Count - 1 && !IsStory && Atoms.Count > 0)
        {
            OpenXmlElement story = FindRoot(Root);
            Paragraph paragraph = ParagraphOfLast();
            for (int hops = 0; paragraph != null && list.Count < MinimumNeighbours && hops < MinimumNeighbours; hops++)
            {
                paragraph = NextParagraph(paragraph, story);
                if (paragraph == null)
                {
                    break;
                }
                foreach (AtomInfo info in Enumerate(paragraph))
                {
                    if (list.Count >= wanted)
                    {
                        break;
                    }
                    list.Add(info.Element);
                }
            }
        }
        return list;
    }

    private Paragraph ParagraphOfFirst()
    {
        OpenXmlElement element = Atoms[0].Element;
        return element as Paragraph ?? element.Ancestors<Paragraph>().FirstOrDefault();
    }

    private Paragraph ParagraphOfLast()
    {
        OpenXmlElement element = Atoms[Atoms.Count - 1].Element;
        return element as Paragraph ?? element.Ancestors<Paragraph>().FirstOrDefault();
    }

    /// <summary>
    /// The paragraph before the given one in the document, or null before the first one in the story.
    /// </summary>
    public static Paragraph PreviousParagraph(Paragraph paragraph, OpenXmlElement storyRoot)
    {
        OpenXmlElement current = paragraph;
        while (current != null && !ReferenceEquals(current, storyRoot))
        {
            for (OpenXmlElement sibling = current.PreviousSibling(); sibling != null; sibling = sibling.PreviousSibling())
            {
                Paragraph found = LastParagraphIn(sibling);
                if (found != null)
                {
                    return found;
                }
            }
            current = current.Parent;
        }
        return null;
    }

    //The last paragraph in an element, found from the end without looking at the rest.
    private static Paragraph LastParagraphIn(OpenXmlElement element)
    {
        if (element is Paragraph paragraph)
        {
            return paragraph;
        }
        if (element is Run || element is SectionProperties || !(element is OpenXmlCompositeElement))
        {
            return null;
        }
        for (OpenXmlElement child = element.LastChild; child != null; child = child.PreviousSibling())
        {
            Paragraph found = LastParagraphIn(child);
            if (found != null)
            {
                return found;
            }
        }
        return null;
    }

    /// <summary>
    /// The atoms that lie wholly inside the positions from <paramref name="start"/> up to <paramref name="end"/>, in order.
    /// </summary>
    public IEnumerable<Atom> AtomsIn(int start, int end)
    {
        int index = start >= Length ? Atoms.Count : IndexAt(start);
        for (int i = Math.Max(index, 0); i < Atoms.Count && Atoms[i].Start < end; i++)
        {
            //The begin, code and separator of a field written as one element cannot be split, so a range that takes in part of them has all.
            if ((Atoms[i].Start >= start && Atoms[i].End <= end) || (Atoms[i].IsSimpleFieldPart && Atoms[i].End > start))
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
    /// for the end of a paragraph, a bell for the end of a table row and a null for a position that shows no text (see
    /// <see cref="Atom.IsHidden"/>).
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
                case AtomKind.Tab:
                case AtomKind.Break:
                    characters[atom.Start] = atom.Text[0];
                    break;
                default:
                    //Positions that show no text are already null characters.
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
        if (atom.IsMark || atom.IsSimpleFieldPart)
        {
            return this;
        }

        bool changed = false;
        OpenXmlElement element = atom.Element;
        if ((element is Text || element is FieldCode) && position > atom.Start)
        {
            element = SplitText(element, position - atom.Start);
            changed = true;
        }

        //The first character of a run: the element must not follow another text, tab, break or object in the same run.
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
        return TryClassify(element, out _);
    }

    //Splits a text or field code element in two at a character, the second part taking the rest. Both parts keep their spaces.
    private static OpenXmlElement SplitText(OpenXmlElement element, int offset)
    {
        if (element is FieldCode code)
        {
            string codeValue = code.Text;
            code.Text = codeValue.Substring(0, offset);
            code.Space = SpaceProcessingModeValues.Preserve;
            FieldCode secondCode = new FieldCode(codeValue.Substring(offset)) { Space = SpaceProcessingModeValues.Preserve };
            code.InsertAfterSelf(secondCode);
            SpanRegistry.TextSplit(code, secondCode, offset);
            return secondCode;
        }

        Text text = (Text)element;
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
}
