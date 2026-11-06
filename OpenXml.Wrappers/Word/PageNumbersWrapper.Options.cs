using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The character that comes between a chapter number and a page number, as <c>WdSeparatorType</c> in Word (with the same values).
/// </summary>
public enum SeparatorType
{
    /// <summary>A hyphen (the default).</summary>
    Hyphen = 0,

    /// <summary>A period.</summary>
    Period = 1,

    /// <summary>A colon.</summary>
    Colon = 2,

    /// <summary>An em dash.</summary>
    EmDash = 3,

    /// <summary>An en dash.</summary>
    EnDash = 4
}

public partial class PageNumbersWrapper : IEnumerable<PageNumberWrapper>
{
    private IEnumerable<FieldWrapper> PageFields()
    {
        return _headerFooter.Range.Fields.Where(f => f.Type == FieldType.Page).ToList();
    }

    /// <summary>The page number at the index (0-based), as <c>PageNumbers.Item</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">There is no page number at the index.</exception>
    public PageNumberWrapper this[int index]
    {
        get
        {
            Guard.ThrowIfNegative(index);
            FieldWrapper field = PageFields().ElementAtOrDefault(index);
            if (field == null)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new PageNumberWrapper(field);
        }
    }

    /// <inheritdoc/>
    public IEnumerator<PageNumberWrapper> GetEnumerator()
    {
        foreach (FieldWrapper field in PageFields())
        {
            yield return new PageNumberWrapper(field);
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    private void RemoveNumberingIfEmpty()
    {
        PageNumberType numbering = Existing;
        if (numbering != null && !numbering.HasAttributes)
        {
            Setup.RemoveProperty<PageNumberType>();
        }
    }

    /// <summary>True to put the number of the chapter before the page number, as <c>PageNumbers.IncludeChapterNumber</c> in Word.</summary>
    public bool IncludeChapterNumber
    {
        get { return Existing?.ChapterStyle != null && Existing.ChapterStyle.HasValue; }
        set
        {
            if (value)
            {
                PageNumberType numbering = Setup.GetOrAddProperty<PageNumberType>();
                if (numbering.ChapterStyle == null || !numbering.ChapterStyle.HasValue)
                {
                    numbering.ChapterStyle = 1;
                }
            }
            else if (Existing != null)
            {
                Existing.ChapterStyle = null;
                RemoveNumberingIfEmpty();
            }
        }
    }

    /// <summary>
    /// The level of the heading style that starts a chapter, 0 for Heading 1 up to 8 for Heading 9, as <c>PageNumbers.HeadingLevelForChapter</c>
    /// in Word. It is 0 until chapter numbers are included.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The level is not between 0 and 8.</exception>
    /// <exception cref="InvalidOperationException">The chapter number is not included.</exception>
    public int HeadingLevelForChapter
    {
        get { return IncludeChapterNumber ? Existing.ChapterStyle.Value - 1 : 0; }
        set
        {
            if (value < 0 || value > 8)
            {
                throw new ArgumentOutOfRangeException(nameof(value), "The level must be between 0 and 8.");
            }
            if (!IncludeChapterNumber)
            {
                throw new InvalidOperationException("The chapter number must be included before its heading level is set.");
            }
            Existing.ChapterStyle = (byte)(value + 1);
        }
    }

    /// <summary>The character between the chapter number and the page number, as <c>PageNumbers.ChapterPageSeparator</c> in Word.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not defined.</exception>
    public SeparatorType ChapterPageSeparator
    {
        get
        {
            ChapterSeparatorValues? separator = Existing?.ChapterSeparator?.Value;
            return separator == ChapterSeparatorValues.Period ? SeparatorType.Period
                : separator == ChapterSeparatorValues.Colon ? SeparatorType.Colon
                : separator == ChapterSeparatorValues.EmDash ? SeparatorType.EmDash
                : separator == ChapterSeparatorValues.EnDash ? SeparatorType.EnDash : SeparatorType.Hyphen;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            if (value == SeparatorType.Hyphen)
            {
                if (Existing != null)
                {
                    Existing.ChapterSeparator = null;
                    RemoveNumberingIfEmpty();
                }
                return;
            }
            Setup.GetOrAddProperty<PageNumberType>().ChapterSeparator = value == SeparatorType.Period ? ChapterSeparatorValues.Period
                : value == SeparatorType.Colon ? ChapterSeparatorValues.Colon
                : value == SeparatorType.EmDash ? ChapterSeparatorValues.EmDash : ChapterSeparatorValues.EnDash;
        }
    }

    /// <summary>
    /// True (the default) to number the first page of the section, as <c>PageNumbers.ShowFirstPageNumber</c> in Word. Without it the section has a
    /// different first page header and footer, so the first page has none of the page numbers in the others.
    /// </summary>
    public bool ShowFirstPageNumber
    {
        get { return !Setup.DifferentFirstPageHeaderFooter; }
        set { Setup.DifferentFirstPageHeaderFooter = !value; }
    }
}

/// <summary>
/// A page number of a header or footer, as <c>PageNumber</c> in Word.
/// </summary>
public class PageNumberWrapper
{
    private readonly FieldWrapper _field;

    internal PageNumberWrapper(FieldWrapper field)
    {
        _field = field;
    }

    /// <summary>
    /// Where the page number is, as <c>PageNumber.Alignment</c> in Word, read from the frame it is in. Word adds another page number when it
    /// is set, so here it can only be read.
    /// </summary>
    public PageNumberAlignment Alignment
    {
        get
        {
            Paragraph paragraph = _field.Key.Ancestors<Paragraph>().FirstOrDefault();
            HorizontalAlignmentValues? align = paragraph?.ParagraphProperties?.FrameProperties?.XAlign?.Value;
            return align == HorizontalAlignmentValues.Center ? PageNumberAlignment.Center
                : align == HorizontalAlignmentValues.Right ? PageNumberAlignment.Right
                : align == HorizontalAlignmentValues.Inside ? PageNumberAlignment.Inside
                : align == HorizontalAlignmentValues.Outside ? PageNumberAlignment.Outside : PageNumberAlignment.Left;
        }
    }
}
