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
    /// <para>As in Word, applying a paragraph style removes the direct paragraph formatting of the paragraph. Of each kind of direct character
    /// formatting it removes what the style gives the text anyway, and the commonest other value if that covers at least as much of the
    /// paragraph, counting its mark, as the text with the style's value. Highlighting, superscript and subscript, shading, borders, character
    /// scale, language and character styles stay.</para>
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
                        paragraph.RemoveEmptyProperties();
                    }
                }
                return;
            }

            string styleId = StyleHelper.FindStyleId(AnchorElement, value, out StyleValues type);
            if (type == StyleValues.Paragraph)
            {
                //The default style is what a paragraph has when it names none, so Word does not name it.
                bool isDefault = StyleHelper.IsDefaultParagraphStyle(AnchorElement, styleId);
                foreach (Paragraph paragraph in GetStyledParagraphs())
                {
                    if (isDefault)
                    {
                        paragraph.ParagraphProperties?.RemoveAllChildren<ParagraphStyleId>();
                    }
                    else
                    {
                        paragraph.GetOrAddParagraphProperties().ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
                    }
                    ClearFormattingUnderParagraphStyle(paragraph);
                    paragraph.RemoveEmptyProperties();
                    if (!(_startOpenXmlElement is Run))
                    {
                        paragraph.CoalesceRuns();
                    }
                }
            }
            else if (type == StyleValues.Character)
            {
                List<Run> runs = GetRuns().ToList();
                foreach (Run run in runs)
                {
                    run.GetOrAddRunProperties().SetRunProperty(new RunStyle() { Val = styleId });
                }
                if (!(_startOpenXmlElement is Run))
                {
                    foreach (OpenXmlElement parent in runs.Select(r => r.Parent).Where(e => e != null).Distinct().ToList())
                    {
                        parent.CoalesceRuns();
                    }
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

    //What a paragraph style leaves of the direct paragraph formatting: these are not formatting, or are looked after separately.
    private static readonly string[] KeptUnderParagraphStyle = { "pStyle", "cnfStyle", "rPr", "sectPr", "pPrChange" };

    //The character formatting a paragraph style can take away. Word leaves the rest, such as highlighting and superscript, whatever it covers.
    private static readonly string[] ClearedByParagraphStyle = { "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "vanish", "color", "spacing", "kern", "position", "sz", "szCs", "u", "em" };

    //Takes off a paragraph what Word takes off when a paragraph style is applied. All of its direct paragraph formatting goes. Of each kind of
    //direct character formatting, what the style gives the text anyway goes, and so does the commonest other value (the first, of those that
    //are equal) if it covers at least as much of the paragraph, counting its mark, as the text with the style's value.
    private static void ClearFormattingUnderParagraphStyle(Paragraph paragraph)
    {
        ParagraphProperties properties = paragraph.ParagraphProperties;
        foreach (OpenXmlElement child in properties?.ChildElements.Where(e => Array.IndexOf(KeptUnderParagraphStyle, e.LocalName) < 0).ToList() ?? new List<OpenXmlElement>())
        {
            child.Remove();
        }

        //The values in the order of the text, each with how much of the paragraph it covers.
        int total = 0;
        List<(OpenXmlElement Value, int Length)> values = new List<(OpenXmlElement Value, int Length)>();
        foreach (Atom atom in StoryMap.Build(paragraph).Atoms)
        {
            total += atom.Length;
            OpenXmlElement holder = atom.IsMark ? properties?.ParagraphMarkRunProperties : (atom.Element.Parent as Run)?.RunProperties;
            foreach (OpenXmlElement value in holder?.ChildElements.Where(e => Array.IndexOf(ClearedByParagraphStyle, e.LocalName) >= 0) ?? Enumerable.Empty<OpenXmlElement>())
            {
                values.Add((value, atom.Length));
            }
        }

        RunProperties style = new FormattingResolver(paragraph).ResolveParagraphStyleRun(paragraph);
        foreach (IGrouping<string, (OpenXmlElement Value, int Length)> kind in values.GroupBy(v => v.Value.LocalName).ToList())
        {
            List<(OpenXmlElement Value, int Length)> differing = kind.Where(v => !FormattingResolver.GivesTheSame(v.Value, style)).ToList();
            int withStyleValue = total - differing.Sum(v => v.Length);
            string commonest = differing.GroupBy(v => v.Value.OuterXml).OrderByDescending(g => g.Sum(v => v.Length)).FirstOrDefault(g => g.Sum(v => v.Length) >= withStyleValue)?.Key;
            foreach (OpenXmlElement value in kind.Select(v => v.Value).Distinct().Where(v => FormattingResolver.GivesTheSame(v, style) || v.OuterXml == commonest).ToList())
            {
                OpenXmlElement holder = value.Parent;

                //The colour of an underline stays, without the underline.
                if (value is Underline underline && underline.Color != null)
                {
                    underline.Val = null;
                }
                else
                {
                    value.Remove();
                }
                if (!holder.HasChildren)
                {
                    holder.Remove();
                }
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
