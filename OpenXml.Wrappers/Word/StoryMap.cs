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
    {
        Root = root;
        Atoms = new List<Atom>();
        Walk(root);
        Length = Atoms.Count == 0 ? 0 : Atoms[Atoms.Count - 1].End;
    }

    public OpenXmlElement Root { get; }

    public List<Atom> Atoms { get; }

    public int Length { get; }

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

        return changed ? new StoryMap(Root) : this;
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
        //Runs can be nested inside other elements (hyperlinks, content controls, smart tags etc.) but runs that belong to a nested paragraph
        //(for example in a text box) are not part of this paragraph.
        foreach (Run run in paragraph.Descendants<Run>())
        {
            if (run.Ancestors<Paragraph>().First() != paragraph)
            {
                continue;
            }
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

        //Every paragraph, including an empty one, ends with an end of paragraph marker, and the last one in a table cell also ends the cell.
        bool endsCell = paragraph.Parent is TableCell && paragraph.Parent.LastChild == paragraph;
        Add(AtomKind.ParagraphMark, paragraph, 1, endsCell ? "\r\a" : "\r");
    }

    private void Add(AtomKind kind, OpenXmlElement element, int length, string text)
    {
        int start = Atoms.Count == 0 ? 0 : Atoms[Atoms.Count - 1].End;
        Atoms.Add(new Atom() { Kind = kind, Element = element, Start = start, Length = length, Text = text });
    }

    #endregion
}
