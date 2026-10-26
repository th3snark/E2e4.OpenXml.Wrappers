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
    /// The name of the style of the range, as <c>Range.Style</c> in Word.
    /// </summary>
    /// <remarks>
    /// <para>For a range of whole paragraphs this is the paragraph style, or null if the paragraphs have different styles. For a bookmark or
    /// comment range it is the character style if every run has the same one, otherwise the paragraph style.</para>
    /// <para>The name is the one stored in the document, in which the names of built-in styles are lower case ("heading 1") although
    /// Word shows them as "Heading 1". Names are matched without regard to case when setting, so compare them without regard to case too.</para>
    /// <para>Setting a paragraph style applies it to every paragraph the range touches, and a character style to every run of the range. Setting
    /// null removes the direct style so that the default applies. A table style cannot be set here; use <see cref="TableWrapper.Style"/>.</para>
    /// </remarks>
    /// <exception cref="ArgumentException">There is no style with the name.</exception>
    /// <exception cref="InvalidOperationException">The document does not have any styles.</exception>
    public string Style
    {
        get
        {
            bool partial = IsPartialParagraphRange();

            if (partial)
            {
                //A character style applied to all the runs is the style of the range.
                List<string> characterStyleIds = GetRuns().Select(r => r.RunProperties?.RunStyle?.Val?.Value).Distinct().ToList();
                if (characterStyleIds.Count == 1 && characterStyleIds[0] != null)
                {
                    return StyleHelper.GetStyleName(AnchorElement, characterStyleIds[0]);
                }
            }

            List<string> paragraphStyleIds = GetStyledParagraphs().Select(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value).Distinct().ToList();
            if (paragraphStyleIds.Count != 1)
            {
                return null;
            }
            return paragraphStyleIds[0] == null
                ? StyleHelper.GetDefaultStyleName(AnchorElement, StyleValues.Paragraph) ?? "Normal"
                : StyleHelper.GetStyleName(AnchorElement, paragraphStyleIds[0]);
        }
        set
        {
            //This throws for the types of range that do not contain text, before anything is looked up.
            GetRunsByParagraph();

            bool partial = IsPartialParagraphRange();
            if (value == null)
            {
                if (partial)
                {
                    foreach (Run run in GetRuns().Where(r => r.RunProperties != null))
                    {
                        run.RunProperties.RemoveRunProperty("rStyle");
                        if (!run.RunProperties.HasChildren)
                        {
                            run.RunProperties.Remove();
                        }
                    }
                }
                else
                {
                    foreach (Paragraph paragraph in GetStyledParagraphs().Where(p => p.ParagraphProperties != null))
                    {
                        paragraph.ParagraphProperties.RemoveAllChildren<ParagraphStyleId>();
                    }
                }
                return;
            }

            string styleId = StyleHelper.FindStyleId(AnchorElement, value, out StyleValues type);
            if (type == StyleValues.Paragraph)
            {
                foreach (Paragraph paragraph in GetStyledParagraphs())
                {
                    paragraph.GetOrAddParagraphProperties().ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
                }
            }
            else if (type == StyleValues.Character)
            {
                foreach (Run run in GetRuns())
                {
                    run.GetOrAddRunProperties().SetRunProperty(new RunStyle() { Val = styleId });
                }
            }
            else if (type == StyleValues.Table)
            {
                throw new NotSupportedException("A table style cannot be applied to a range. Use the Style property of the table.");
            }
            else
            {
                throw new NotSupportedException(string.Format(CultureInfo.InvariantCulture, "Applying a style of the type {0} is not currently supported.", type));
            }
        }
    }

    /// <summary>
    /// The style of the range as an object, as <c>Range.Style</c> returns it in Word (null where the range has different styles, or the
    /// document does not define the style). Setting it applies the style as <see cref="Style"/> does, and null removes the direct style, so
    /// <c>range.StyleObject = document.Styles["Heading 1"]</c> and <c>= document.Styles[BuiltInStyle.Heading1]</c> work.
    /// </summary>
    public StyleWrapper StyleObject
    {
        get { return StyleHelper.GetStyleObject(AnchorElement, Style); }
        set { Style = value?.Name; }
    }

    /// <summary>
    /// True for the ranges that cover part of a paragraph, whose runs can have a character style.
    /// </summary>
    private bool IsPartialParagraphRange()
    {
        return _isSpan || _startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart;
    }

    /// <summary>
    /// The paragraphs that have the style of the range: all those it touches.
    /// </summary>
    private IList<Paragraph> GetStyledParagraphs()
    {
        //This throws for the types of range that do not contain text.
        GetRunsByParagraph();
        return Paragraphs.Select(p => p.Paragraph).ToList();
    }
}
