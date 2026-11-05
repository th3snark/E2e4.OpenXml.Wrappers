using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// A piece of a story taken out of a document, to be put into another place: paragraphs and tables that are copies (they are not in any tree),
/// with the parts of the first and last paragraph that were outside the range trimmed away. It is what <c>Range.FormattedText</c>,
/// <c>InsertFile</c> and the other ways of copying content use.
/// </summary>
internal sealed class Fragment
{
    private Fragment()
    {
    }

    /// <summary>The paragraphs and tables, in order.</summary>
    public List<OpenXmlElement> Blocks { get; } = new List<OpenXmlElement>();

    /// <summary>
    /// True if the last paragraph comes with its paragraph mark (so that putting the fragment in the middle of a paragraph splits it there),
    /// false if the fragment ends in the middle of a paragraph, whose text then joins what follows the place it goes to.
    /// </summary>
    public bool HasFinalMark { get; set; }

    /// <summary>The part the content was taken from (a header, say), or null for content that was not in a document.</summary>
    public OpenXmlPart SourcePart { get; private set; }

    public bool IsEmpty
    {
        get { return Blocks.Count == 0; }
    }

    /// <summary>
    /// Copies the content of a range. Paragraphs the range only touches are trimmed to the part inside it; a table the range touches is taken
    /// whole. The bookmarks and comments stay with the original, because they can only be in one place.
    /// </summary>
    public static Fragment From(RangeWrapper range)
    {
        range.GetPositions(out StoryMap map, out int start, out int end);
        Fragment fragment = new Fragment();
        OpenXmlElement story = StoryMap.FindRoot(range.AnchorElement);
        fragment.SourcePart = PartOf(story);
        if (end <= start)
        {
            return fragment;
        }

        //The blocks of the story that the range touches, and where the first one starts.
        List<OpenXmlElement> blocks = new List<OpenXmlElement>();
        int firstStart = -1;
        foreach (Atom atom in map.Atoms)
        {
            if (atom.End <= start || atom.Start >= end)
            {
                continue;
            }
            OpenXmlElement block = atom.Element;
            while (block != null && !ReferenceEquals(block.Parent, story))
            {
                block = block.Parent;
            }
            if (block != null && (block is Paragraph || block is Table) && (blocks.Count == 0 || !ReferenceEquals(blocks[blocks.Count - 1], block)))
            {
                if (blocks.Count == 0)
                {
                    firstStart = FirstPositionOfBlock(map, block);
                }
                blocks.Add(block);
            }
        }
        if (blocks.Count == 0)
        {
            return fragment;
        }

        //Trim copies of them in a story of their own, with the same positions less where the first one starts.
        Body copy = new Body();
        foreach (OpenXmlElement block in blocks)
        {
            copy.AppendChild(block.CloneNode(true));
        }
        int total = StoryMap.CountLength(copy);
        int localStart = start - firstStart;
        int localEnd = end - firstStart;
        bool endsInParagraph = blocks[blocks.Count - 1] is Paragraph;
        bool startsInParagraph = blocks[0] is Paragraph;
        bool endsMid = endsInParagraph && localEnd < total;
        if (endsMid && localEnd < total - 1)
        {
            RangeWrapper tail = new RangeWrapper(copy);
            tail.SetRange(localEnd, total - 1);
            tail.Delete();
        }
        if (startsInParagraph && localStart > 0)
        {
            RangeWrapper head = new RangeWrapper(copy);
            head.SetRange(0, localStart);
            head.Delete();
        }

        foreach (OpenXmlElement block in copy.ChildElements.ToList())
        {
            block.Remove();
            fragment.Blocks.Add(block);
        }
        fragment.HasFinalMark = !endsMid;
        fragment.Clean();
        return fragment;
    }

    /// <summary>
    /// Copies the content of the main text of a document: all its paragraphs and tables. The last paragraph comes with its mark.
    /// </summary>
    public static Fragment From(WordprocessingDocument document)
    {
        Fragment fragment = new Fragment();
        Body body = document.MainDocumentPart?.Document?.Body;
        fragment.SourcePart = document.MainDocumentPart;
        if (body == null)
        {
            return fragment;
        }
        foreach (OpenXmlElement block in body.ChildElements.Where(e => e is Paragraph || e is Table))
        {
            fragment.Blocks.Add(block.CloneNode(true));
        }
        fragment.HasFinalMark = true;
        fragment.Clean();
        return fragment;
    }

    //Takes out what cannot be copied to another place: section breaks, bookmarks, comments and what makes an element unique.
    private void Clean()
    {
        foreach (OpenXmlElement block in Blocks)
        {
            foreach (SectionProperties section in block.Descendants<SectionProperties>().ToList())
            {
                section.Remove();
            }
            foreach (ParagraphProperties properties in block.Descendants<ParagraphProperties>().Where(p => !p.HasChildren).ToList())
            {
                properties.Remove();
            }
            block.RemoveAnchors();
        }
    }

    private static int FirstPositionOfBlock(StoryMap map, OpenXmlElement block)
    {
        foreach (Atom atom in map.Atoms)
        {
            OpenXmlElement element = atom.Element;
            while (element != null && !ReferenceEquals(element, block))
            {
                element = element.Parent;
            }
            if (element != null)
            {
                return atom.Start;
            }
        }
        return 0;
    }

    private static OpenXmlPart PartOf(OpenXmlElement story)
    {
        return (story.GetRoot() as OpenXmlPartRootElement)?.OpenXmlPart;
    }

    /// <summary>
    /// Makes the content fit the story it is going to: what it refers to (links, pictures) is added to the part of the story, the styles it
    /// uses are added to the document if it is not the one it came from, and pictures get ids of their own.
    /// </summary>
    public void PrepareFor(OpenXmlElement targetStory)
    {
        OpenXmlPart targetPart = PartOf(targetStory);
        if (SourcePart != null && targetPart != null && !ReferenceEquals(SourcePart, targetPart))
        {
            CopyRelationships(SourcePart, targetPart);
        }

        MainDocumentPart sourceMain = SourcePart as MainDocumentPart ?? SourcePart?.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
        MainDocumentPart targetMain = StyleHelper.GetMainDocumentPart(targetStory);
        if (sourceMain != null && targetMain != null && !ReferenceEquals(sourceMain, targetMain))
        {
            CopyStyles(sourceMain, targetMain);
            foreach (NumberingProperties numbering in Blocks.SelectMany(b => b.Descendants<NumberingProperties>()).ToList())
            {
                numbering.Remove();
            }
        }

        foreach (OpenXmlElement block in Blocks)
        {
            block.GiveDrawingsNewIds();
        }
    }

    //The relationships that the content refers to by id (links, pictures), added to the target part with ids of their own.
    private void CopyRelationships(OpenXmlPart source, OpenXmlPart target)
    {
        Dictionary<string, string> ids = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (OpenXmlElement element in Blocks.SelectMany(b => b.Descendants().Concat(new[] { b })).ToList())
        {
            foreach (OpenXmlAttribute attribute in element.GetAttributes().Where(a => a.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/relationships" && !string.IsNullOrEmpty(a.Value)).ToList())
            {
                string newId = ids.TryGetValue(attribute.Value, out string known) ? known : CopyRelationship(source, target, attribute.Value);
                ids[attribute.Value] = newId;
                if (newId != attribute.Value)
                {
                    element.SetAttribute(new OpenXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, newId));
                }
            }
        }
    }

    private static string CopyRelationship(OpenXmlPart source, OpenXmlPart target, string id)
    {
        foreach (IdPartPair part in source.Parts)
        {
            if (part.RelationshipId == id)
            {
                return target.GetIdOfPart(target.AddPart(part.OpenXmlPart));
            }
        }
        HyperlinkRelationship hyperlink = source.HyperlinkRelationships.FirstOrDefault(r => r.Id == id);
        if (hyperlink != null)
        {
            return target.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal).Id;
        }
        ExternalRelationship external = source.ExternalRelationships.FirstOrDefault(r => r.Id == id);
        if (external != null)
        {
            return target.AddExternalRelationship(external.RelationshipType, external.Uri).Id;
        }
        return id;
    }

    //The styles the content uses (and the styles they are based on, linked to or followed by) that the target document does not have.
    private void CopyStyles(MainDocumentPart source, MainDocumentPart target)
    {
        Styles sourceStyles = source.StyleDefinitionsPart?.Styles;
        if (sourceStyles == null)
        {
            return;
        }
        if (target.StyleDefinitionsPart == null)
        {
            target.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
        }
        Styles targetStyles = target.StyleDefinitionsPart.Styles ?? (target.StyleDefinitionsPart.Styles = new Styles());

        Queue<string> pending = new Queue<string>(Blocks.SelectMany(b => b.Descendants())
            .Select(e => e is ParagraphStyleId p ? p.Val?.Value : e is RunStyle r ? r.Val?.Value : e is TableStyle t ? t.Val?.Value : null)
            .Where(id => id != null)
            .Distinct());
        HashSet<string> done = new HashSet<string>(StringComparer.Ordinal);
        while (pending.Count > 0)
        {
            string id = pending.Dequeue();
            if (!done.Add(id))
            {
                continue;
            }
            Style style = sourceStyles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == id);
            if (style == null)
            {
                continue;
            }
            if (!targetStyles.Elements<Style>().Any(s => s.StyleId?.Value == id))
            {
                targetStyles.AppendChild((Style)style.CloneNode(true));
            }
            foreach (string related in new[] { style.BasedOn?.Val?.Value, style.NextParagraphStyle?.Val?.Value, style.LinkedStyle?.Val?.Value })
            {
                if (related != null)
                {
                    pending.Enqueue(related);
                }
            }
        }
    }
}
