using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// Inserts a page, column or line break, or a section break, as <c>Range.InsertBreak</c> in Word.
    /// </summary>
    /// <remarks>
    /// For a page, column or line break in a span of characters the span is replaced by the break (collapse it first to keep the text, as in
    /// Word). A range made of whole elements (a paragraph, say) is not replaced: the break is inserted at its start. A section break is inserted
    /// immediately before the range: the paragraph is split there and the first part, which may be empty, ends the section.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">The type is not a break type.</exception>
    /// <exception cref="NotSupportedException">The range is of a type that does not support inserting breaks.</exception>
    public void InsertBreak(BreakType type = BreakType.PageBreak)
    {
        if (!EnumCompat.IsDefined(type))
        {
            throw new ArgumentOutOfRangeException(nameof(type));
        }

        string breakText = type == BreakType.PageBreak ? "\f"
            : type == BreakType.ColumnBreak ? "\u000e"
            : type == BreakType.LineBreak || type == BreakType.TextWrappingBreak ? "\v"
            : null;
        if (breakText != null)
        {
            if (_isSpan)
            {
                Text = breakText;
            }
            else
            {
                InsertBefore(breakText);
            }
            return;
        }

        string sectionType = type == BreakType.SectionBreakContinuous ? "continuous"
            : type == BreakType.SectionBreakEvenPage ? "evenPage"
            : type == BreakType.SectionBreakOddPage ? "oddPage"
            : "nextPage";
        ParagraphWrapper inserted = InsertParagraphBefore();
        //For a span the new paragraph is the second half of the split one; the first half ends the section.
        Paragraph ending = _isSpan && inserted.Paragraph.PreviousSibling<Paragraph>() != null ? inserted.Paragraph.PreviousSibling<Paragraph>() : inserted.Paragraph;
        new ParagraphWrapper(ending).InsertSectionBreak(sectionType);
    }
    /// <summary>
    /// Makes room for a block (a table, say) at the start of a span: the text of a span that is not collapsed is deleted, and the paragraph is
    /// split at that point unless it is already at the start of one. Returns the paragraph that the block goes in front of.
    /// </summary>
    internal Paragraph PrepareInsertBlock()
    {
        if (!_isSpan)
        {
            throw new InvalidOperationException("The range is not a span of characters.");
        }

        GetScoped(false, out StoryMap map, out int start, out int end);
        if (start != end)
        {
            Delete();
            GetScoped(false, out map, out start, out end);
        }

        Paragraph holder = ParagraphForPosition(map, start, out int local);
        return local == 0 ? holder : SplitParagraphAt(holder, local);
    }
}
