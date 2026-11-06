using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The parts of a document that hold text, as <c>WdStoryType</c> in Word (with the same values).
/// </summary>
public enum StoryType
{
    /// <summary>The main text.</summary>
    MainText = 1,

    /// <summary>The footnotes.</summary>
    Footnotes = 2,

    /// <summary>The endnotes.</summary>
    Endnotes = 3,

    /// <summary>The comments.</summary>
    Comments = 4,

    /// <summary>The text of a text box or shape.</summary>
    TextFrame = 5,

    /// <summary>The header of even pages.</summary>
    EvenPagesHeader = 6,

    /// <summary>The header of the pages that have no other.</summary>
    PrimaryHeader = 7,

    /// <summary>The footer of even pages.</summary>
    EvenPagesFooter = 8,

    /// <summary>The footer of the pages that have no other.</summary>
    PrimaryFooter = 9,

    /// <summary>The header of the first page.</summary>
    FirstPageHeader = 10,

    /// <summary>The footer of the first page.</summary>
    FirstPageFooter = 11,

    /// <summary>The line above the footnotes.</summary>
    FootnoteSeparator = 12,

    /// <summary>The line above the footnotes that carry on from the page before.</summary>
    FootnoteContinuationSeparator = 13,

    /// <summary>The notice that footnotes carry on to the next page.</summary>
    FootnoteContinuationNotice = 14,

    /// <summary>The line above the endnotes.</summary>
    EndnoteSeparator = 15,

    /// <summary>The line above the endnotes that carry on from the page before.</summary>
    EndnoteContinuationSeparator = 16,

    /// <summary>The notice that endnotes carry on to the next page.</summary>
    EndnoteContinuationNotice = 17
}

public partial class RangeWrapper
{
    /// <summary>
    /// The kind of story the range is in, as <c>Range.StoryType</c> in Word. A range in a text box is in <see cref="Word.StoryType.TextFrame"/>,
    /// as it is in Word.
    /// </summary>
    public StoryType StoryType
    {
        get
        {
            OpenXmlElement anchor = AnchorElement;
            if (anchor.Ancestors<TextBoxContent>().Any())
            {
                return StoryType.TextFrame;
            }

            OpenXmlElement root = StoryMap.FindRoot(anchor);
            if (root is Header || root is Footer)
            {
                return HeaderFooterStoryType(root);
            }
            if (root is Comment)
            {
                return StoryType.Comments;
            }
            if (root is Footnote footnote)
            {
                FootnoteEndnoteValues? type = footnote.Type?.Value;
                return type == FootnoteEndnoteValues.Separator ? StoryType.FootnoteSeparator
                    : type == FootnoteEndnoteValues.ContinuationSeparator ? StoryType.FootnoteContinuationSeparator
                    : type == FootnoteEndnoteValues.ContinuationNotice ? StoryType.FootnoteContinuationNotice : StoryType.Footnotes;
            }
            if (root is Endnote endnote)
            {
                FootnoteEndnoteValues? type = endnote.Type?.Value;
                return type == FootnoteEndnoteValues.Separator ? StoryType.EndnoteSeparator
                    : type == FootnoteEndnoteValues.ContinuationSeparator ? StoryType.EndnoteContinuationSeparator
                    : type == FootnoteEndnoteValues.ContinuationNotice ? StoryType.EndnoteContinuationNotice : StoryType.Endnotes;
            }
            return StoryType.MainText;
        }
    }

    //The kind of a header or footer is the one its first reference in the document gives it.
    private static StoryType HeaderFooterStoryType(OpenXmlElement root)
    {
        bool isHeader = root is Header;
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(root);
        OpenXmlPart part = (root as OpenXmlPartRootElement)?.OpenXmlPart;
        string id = part == null || main == null ? null : main.GetIdOfPart(part);
        HeaderFooterValues? type = null;
        if (id != null && main.Document?.Body != null)
        {
            if (isHeader)
            {
                type = main.Document.Body.Descendants<HeaderReference>().FirstOrDefault(r => r.Id?.Value == id)?.Type?.Value;
            }
            else
            {
                type = main.Document.Body.Descendants<FooterReference>().FirstOrDefault(r => r.Id?.Value == id)?.Type?.Value;
            }
        }
        if (type == HeaderFooterValues.First)
        {
            return isHeader ? StoryType.FirstPageHeader : StoryType.FirstPageFooter;
        }
        if (type == HeaderFooterValues.Even)
        {
            return isHeader ? StoryType.EvenPagesHeader : StoryType.EvenPagesFooter;
        }
        return isHeader ? StoryType.PrimaryHeader : StoryType.PrimaryFooter;
    }

    /// <summary>
    /// The same story in the next section that has one of its own, as <c>Range.NextStoryRange</c> in Word: the header or footer of the next section
    /// that does not show the one before it. Null for every other kind of story, and for the last such header or footer.
    /// </summary>
    public RangeWrapper NextStoryRange
    {
        get
        {
            StoryType type = StoryType;
            bool isHeader = type == StoryType.PrimaryHeader || type == StoryType.FirstPageHeader || type == StoryType.EvenPagesHeader;
            bool isFooter = type == StoryType.PrimaryFooter || type == StoryType.FirstPageFooter || type == StoryType.EvenPagesFooter;
            WordprocessingDocumentWrapper document = Document;
            if ((!isHeader && !isFooter) || document == null)
            {
                return null;
            }

            HeaderFooterIndex index = type == StoryType.FirstPageHeader || type == StoryType.FirstPageFooter ? HeaderFooterIndex.FirstPage
                : type == StoryType.EvenPagesHeader || type == StoryType.EvenPagesFooter ? HeaderFooterIndex.EvenPages : HeaderFooterIndex.Primary;
            OpenXmlElement root = StoryMap.FindRoot(AnchorElement);
            List<SectionWrapper> sections = document.Sections.ToList();
            int current = -1;
            for (int i = 0; i < sections.Count && current < 0; i++)
            {
                HeaderFooterWrapper own = isHeader ? sections[i].Headers[index] : sections[i].Footers[index];
                if (own.Exists && ReferenceEquals(StoryMap.FindRoot(own.Range.AnchorElement), root))
                {
                    current = i;
                }
            }
            for (int i = current + 1; current >= 0 && i < sections.Count; i++)
            {
                HeaderFooterWrapper next = isHeader ? sections[i].Headers[index] : sections[i].Footers[index];
                if (next.Exists)
                {
                    return next.Range;
                }
            }
            return null;
        }
    }
}

/// <summary>
/// The first range of each kind of story a document has, as <c>StoryRanges</c> in Word. Only the stories the library can read are listed: the
/// main text, the first comment (Word joins all the comments into one story, where each is a story of its own here), and the headers and\n/// footers of the first section. The footnotes, endnotes and text boxes follow in the steps that
/// add them.
/// </summary>
public class StoryRangesWrapper : IEnumerable<RangeWrapper>
{
    private readonly WordprocessingDocumentWrapper _document;

    internal StoryRangesWrapper(WordprocessingDocumentWrapper document)
    {
        _document = document;
    }

    /// <summary>The document the stories are in.</summary>
    public WordprocessingDocumentWrapper Parent
    {
        get { return _document; }
    }

    private RangeWrapper Find(StoryType type)
    {
        if (type == StoryType.MainText)
        {
            return _document.Content;
        }
        if (type == StoryType.Comments)
        {
            //Each comment is a story of its own here, where Word joins them into one: this is the first.
            Comment firstComment = _document.WordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>().FirstOrDefault();
            Paragraph firstParagraph = firstComment?.Elements<Paragraph>().FirstOrDefault();
            return firstComment == null ? null : firstParagraph == null ? new RangeWrapper(firstComment) : new RangeWrapper(firstParagraph, firstComment.Elements<Paragraph>().Last());
        }

        bool isHeader = type == StoryType.PrimaryHeader || type == StoryType.FirstPageHeader || type == StoryType.EvenPagesHeader;
        bool isFooter = type == StoryType.PrimaryFooter || type == StoryType.FirstPageFooter || type == StoryType.EvenPagesFooter;
        if (!isHeader && !isFooter)
        {
            return null;
        }
        HeaderFooterIndex index = type == StoryType.FirstPageHeader || type == StoryType.FirstPageFooter ? HeaderFooterIndex.FirstPage
            : type == StoryType.EvenPagesHeader || type == StoryType.EvenPagesFooter ? HeaderFooterIndex.EvenPages : HeaderFooterIndex.Primary;
        SectionWrapper first = _document.Sections.First;
        HeaderFooterWrapper headerFooter = isHeader ? first.Headers[index] : first.Footers[index];
        return headerFooter.Exists ? headerFooter.Range : null;
    }

    private IEnumerable<StoryType> Types()
    {
        foreach (StoryType type in new[] { StoryType.MainText, StoryType.Comments, StoryType.EvenPagesHeader, StoryType.PrimaryHeader, StoryType.EvenPagesFooter,
            StoryType.PrimaryFooter, StoryType.FirstPageHeader, StoryType.FirstPageFooter })
        {
            if (Find(type) != null)
            {
                yield return type;
            }
        }
    }

    /// <summary>The range of the first story of the kind, as <c>StoryRanges.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The document has no story of the kind.</exception>
    public RangeWrapper this[StoryType type]
    {
        get
        {
            RangeWrapper range = Find(type);
            if (range == null)
            {
                throw new ArgumentOutOfRangeException(nameof(type), "The document has no story of this kind.");
            }
            return range;
        }
    }

    /// <summary>The number of kinds of story the document has.</summary>
    public int Count
    {
        get { return Types().Count(); }
    }

    /// <inheritdoc/>
    public IEnumerator<RangeWrapper> GetEnumerator()
    {
        foreach (StoryType type in Types().ToList())
        {
            yield return this[type];
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}
