using System;
using System.Globalization;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// Promoting and demoting a paragraph in the outline.
/// </summary>
public partial class ParagraphWrapper
{
    private const string HeadingPrefix = "Heading ";

    /// <summary>
    /// Moves the paragraph one level down in the outline, as <c>Paragraph.OutlineDemote</c> in Word: Heading n becomes Heading n + 1
    /// (Heading 9 stays), a paragraph with a level of its own gets the next level, and body text becomes a heading one level below the
    /// heading before it (Heading 2 if there is none).
    /// </summary>
    public void OutlineDemote()
    {
        MoveInOutline(1);
    }

    /// <summary>
    /// Moves the paragraph one level up in the outline, as <c>Paragraph.OutlinePromote</c> in Word: Heading n becomes Heading n - 1
    /// (Heading 1 stays), a paragraph with a level of its own gets the level before, and body text becomes a heading of the level of the
    /// heading before it (Heading 1 if there is none).
    /// </summary>
    public void OutlinePromote()
    {
        MoveInOutline(-1);
    }

    /// <summary>
    /// Makes the paragraph body text, as <c>Paragraph.OutlineDemoteToBody</c> in Word: it gets the Normal style and loses a level it had of
    /// its own.
    /// </summary>
    public void OutlineDemoteToBody()
    {
        SetStyleKeepingFormatting("Normal");
        Format.OutlineLevel = null;
    }

    private void MoveInOutline(int delta)
    {
        int? heading = HeadingNumber(Style);
        if (heading.HasValue)
        {
            int target = Math.Max(1, Math.Min(9, heading.Value + delta));
            if (target != heading.Value)
            {
                SetStyleKeepingFormatting(HeadingPrefix + target.ToString(CultureInfo.InvariantCulture));
            }
            return;
        }

        int level = (int)(Format.OutlineLevel ?? ParagraphOutlineLevel.BodyText);
        if (level <= 9)
        {
            Format.OutlineLevel = (ParagraphOutlineLevel)Math.Max(1, Math.Min(9, level + delta));
            return;
        }

        //Body text becomes a heading level with the heading before it: below it when demoted, beside it when promoted.
        int basis = LevelBefore() ?? 1;
        SetStyleKeepingFormatting(HeadingPrefix + Math.Max(1, Math.Min(9, delta > 0 ? basis + 1 : basis)).ToString(CultureInfo.InvariantCulture));
    }

    //The outline commands change the style and nothing else: unlike applying a style, they keep the direct formatting.
    private void SetStyleKeepingFormatting(string name)
    {
        string styleId = StyleHelper.FindStyleId(_paragraph, name, out _);
        if (StyleHelper.IsDefaultParagraphStyle(_paragraph, styleId))
        {
            _paragraph.ParagraphProperties?.RemoveAllChildren<ParagraphStyleId>();
        }
        else
        {
            _paragraph.GetOrAddParagraphProperties().ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
        }
        _paragraph.RemoveEmptyProperties();
    }

    //The outline level of the nearest paragraph before this one that has one (1 to 9).
    private int? LevelBefore()
    {
        for (ParagraphWrapper before = Previous(); before != null; before = before.Previous())
        {
            int level = (int)(before.Format.OutlineLevel ?? ParagraphOutlineLevel.BodyText);
            if (level <= 9)
            {
                return level;
            }
        }
        return null;
    }

    private static int? HeadingNumber(string style)
    {
        if (style != null && style.Length == HeadingPrefix.Length + 1 && style.StartsWith(HeadingPrefix, StringComparison.Ordinal)
            && style[HeadingPrefix.Length] >= '1' && style[HeadingPrefix.Length] <= '9')
        {
            return style[HeadingPrefix.Length] - '0';
        }
        return null;
    }
}
