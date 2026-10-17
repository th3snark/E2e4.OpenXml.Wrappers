using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// Inserts text at the start of the range. Tabs and line breaks (<c>\t</c>, <c>\n</c> and <c>\v</c>) in the text become tab and break
    /// characters. A <see cref="RangeWrapper"/> is made of whole elements, so unlike in Word this range object does not need to grow:
    /// the text is inserted inside the range for bookmark and comment ranges and into the first paragraph for the other ranges.
    /// </summary>
    /// <exception cref="NotSupportedException">The range is of a type that does not support inserting text, such as a table or table row.</exception>
    public void InsertBefore(string text)
    {
        Insert(text, false);
    }

    /// <summary>
    /// Inserts text at the end of the range. For a paragraph (and for a table cell or the document) the text is added at the end of the
    /// last paragraph, before its paragraph mark. In Word the text of a whole-paragraph range goes after the paragraph mark, which is at
    /// the start of the next paragraph. See <see cref="InsertBefore"/> for the other rules.
    /// </summary>
    /// <exception cref="NotSupportedException">The range is of a type that does not support inserting text, such as a table or table row.</exception>
    public void InsertAfter(string text)
    {
        Insert(text, true);
    }

    private void Insert(string text, bool after)
    {
        if (String.IsNullOrEmpty(text))
        {
            return;
        }

        if (_isSpan)
        {
            InsertIntoSpan(text, after);
            return;
        }

        InsertRuns(paragraph => BuildNewRunListFromText(text, paragraph), after);
    }

    /// <summary>
    /// Inserts the runs made by <paramref name="build"/> at the start or end of the range, by the rules of <see cref="InsertBefore"/> and
    /// <see cref="InsertAfter"/>.
    /// </summary>
    private void InsertRuns(Func<Paragraph, List<Run>> build, bool after)
    {
        if (_isSpan)
        {
            GetScoped(false, out StoryMap map, out int start, out int end);
            int position = after ? end : start;
            Paragraph paragraph = ParagraphForPosition(map, position, out int local);
            int lengthBefore = StoryMap.Build(paragraph).Length;
            InsertRunsAt(paragraph, local, build, null);

            //The range grows to include what was added.
            StoryMap changed = StoryMap.Build(paragraph);
            if (after)
            {
                MakeSpanAnchors(_startAnchor, Anchor.At(changed, local + changed.Length - lengthBefore));
            }
            else
            {
                MakeSpanAnchors(Anchor.At(changed, local), _endAnchor);
            }
            return;
        }

        if (_startOpenXmlElement is Paragraph)
        {
            InsertRunsIntoParagraph((Paragraph)_startOpenXmlElement, build, after);
        }
        else if (_startOpenXmlElement is TableCell)
        {
            IEnumerable<Paragraph> paragraphs = _startOpenXmlElement.Elements<Paragraph>();
            Paragraph paragraph = after ? paragraphs.LastOrDefault() : paragraphs.FirstOrDefault();
            if (paragraph == null)
            {
                paragraph = new Paragraph();
                _startOpenXmlElement.AppendChild(paragraph);
            }
            InsertRunsIntoParagraph(paragraph, build, after);
        }
        else if (_startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart)
        {
            InsertRunsAtMarker(after ? _endOpenXmlElement : _startOpenXmlElement, build, after);
        }
        else if (_startOpenXmlElement is Document || _startOpenXmlElement is Body)
        {
            InsertRunsIntoBody(GetBody(), build, after);
        }
        else if (_startOpenXmlElement is Header || _startOpenXmlElement is Footer)
        {
            InsertRunsIntoBody((OpenXmlCompositeElement)_startOpenXmlElement, build, after);
        }
        else
        {
            throw new NotSupportedException(String.Format(CultureInfo.InvariantCulture, "Inserting text is not currently supported for a range starting with the type {0}", _startOpenXmlElement.GetType().Name));
        }
    }

    private Body GetBody()
    {
        Body body = _startOpenXmlElement as Body ?? ((Document)_startOpenXmlElement).Body;
        if (body == null)
        {
            throw new InvalidOperationException("The document does not have a body.");
        }
        return body;
    }

    private void InsertRunsIntoParagraph(Paragraph paragraph, Func<Paragraph, List<Run>> build, bool after)
    {
        List<Run> runs = build(paragraph);
        if (after)
        {
            foreach (Run run in runs)
            {
                paragraph.AppendChild(run);
            }
        }
        else
        {
            //Prepend in reverse order so that the runs end up in their original order (and after any paragraph properties).
            for (int i = runs.Count - 1; i >= 0; i--)
            {
                paragraph.PrependRun(runs[i]);
            }
        }
    }

    private void InsertRunsAtMarker(OpenXmlElement marker, Func<Paragraph, List<Run>> build, bool after)
    {
        //New runs can only be placed next to the marker when the marker is a direct child of the paragraph.
        Paragraph paragraph = marker.Parent as Paragraph;
        if (paragraph == null)
        {
            throw new NotSupportedException("Inserting text is only supported when the markers of the range are directly inside a paragraph.");
        }

        List<Run> runs = build(paragraph);
        if (after)
        {
            //The end marker closes the range so the new runs go immediately before it, which keeps them inside the range.
            foreach (Run run in runs)
            {
                marker.InsertBeforeSelf(run);
            }
        }
        else
        {
            for (int i = runs.Count - 1; i >= 0; i--)
            {
                marker.InsertAfterSelf(runs[i]);
            }
        }
    }

    private void InsertRunsIntoBody(OpenXmlCompositeElement body, Func<Paragraph, List<Run>> build, bool after)
    {
        //The section properties are not content, so they are never the first or last block.
        OpenXmlElement block = after
            ? body.ChildElements.LastOrDefault(e => !(e is SectionProperties))
            : body.ChildElements.FirstOrDefault(e => !(e is SectionProperties));

        if (block is Paragraph)
        {
            InsertRunsIntoParagraph((Paragraph)block, build, after);
            return;
        }

        //The body starts or ends with something other than a paragraph (such as a table) so the text needs a paragraph of its own.
        Paragraph paragraph = new Paragraph();
        foreach (Run run in build(paragraph))
        {
            paragraph.AppendChild(run);
        }
        if (block == null)
        {
            SectionProperties sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
            if (sectionProperties != null)
            {
                sectionProperties.InsertBeforeSelf(paragraph);
            }
            else
            {
                body.AppendChild(paragraph);
            }
        }
        else if (after)
        {
            block.InsertAfterSelf(paragraph);
        }
        else
        {
            block.InsertBeforeSelf(paragraph);
        }
    }
}
