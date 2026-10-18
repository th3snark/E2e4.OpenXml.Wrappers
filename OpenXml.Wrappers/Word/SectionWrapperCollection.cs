using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The sections of a document, as <c>Sections</c> in Word.
/// </summary>
public class SectionWrapperCollection : IEnumerable<SectionWrapper>
{
    private readonly Body _body;

    public SectionWrapperCollection(Body body)
    {
        Guard.ThrowIfNull(body);
        _body = body;
    }

    /// <summary>
    /// The range of the whole document, as the parent of the other collections of the document.
    /// </summary>
    public RangeWrapper Parent
    {
        get
        {
            Document document = _body.Ancestors<Document>().FirstOrDefault();
            return new RangeWrapper(document == null ? _body : document);
        }
    }

    public SectionWrapper this[int index]
    {
        get
        {
            IList<SectionWrapper> sections = GetList();
            if (index < 0 || index >= sections.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(index), String.Format(CultureInfo.InvariantCulture, "The requested member of the collection '{0}' does not exist.", index));
            }
            return sections[index];
        }
    }

    /// <summary>The first section.</summary>
    public SectionWrapper First
    {
        get { return GetList()[0]; }
    }

    /// <summary>The last section.</summary>
    public SectionWrapper Last
    {
        get { return GetList().Last(); }
    }

    /// <summary>
    /// Ends the section at the range with a section break and returns the section that starts after it, as <c>Sections.Add</c> in Word. The
    /// range is the place of the break (the end of the document when omitted); see <see cref="RangeWrapper.InsertBreak"/>.
    /// </summary>
    /// <param name="range">Where the break goes, or null for the end of the document.</param>
    /// <param name="start">Where the new section starts: on a new page, continuously, or on an even or odd page.</param>
    /// <exception cref="ArgumentOutOfRangeException">The start is not one that a section can have (a new column is not supported here).</exception>
    public SectionWrapper Add(RangeWrapper range = null, SectionStart start = SectionStart.NewPage)
    {
        if (!EnumCompat.IsDefined(start) || start == SectionStart.NewColumn)
        {
            throw new ArgumentOutOfRangeException(nameof(start));
        }

        if (range == null)
        {
            range = Parent.Duplicate();
            range.SetRange(range.StoryLength, range.StoryLength);
        }

        HashSet<SectionProperties> before = new HashSet<SectionProperties>(_body.Descendants<ParagraphProperties>().Select(p => p.GetFirstChild<SectionProperties>()).Where(s => s != null));
        range.InsertBreak(start == SectionStart.Continuous ? BreakType.SectionBreakContinuous
            : start == SectionStart.EvenPage ? BreakType.SectionBreakEvenPage
            : start == SectionStart.OddPage ? BreakType.SectionBreakOddPage : BreakType.SectionBreakNextPage);

        //The break ends a section, so the new section is the one after the section properties that were added.
        List<SectionProperties> after = _body.Descendants<ParagraphProperties>().Select(p => p.GetFirstChild<SectionProperties>()).Where(s => s != null).ToList();
        int added = after.FindIndex(s => !before.Contains(s));
        return GetList()[added + 1];
    }

    public int Count
    {
        get { return GetList().Count; }
    }

    private IList<SectionWrapper> GetList()
    {
        List<SectionWrapper> sections = _body.Descendants<ParagraphProperties>()
            .Select(p => p.GetFirstChild<SectionProperties>())
            .Where(s => s != null)
            .Select(s => new SectionWrapper(_body, s))
            .ToList();

        //The last section always exists, whether or not its properties have been written.
        sections.Add(new SectionWrapper(_body, null));
        return sections;
    }

    #region IEnumerable Members

    public IEnumerator<SectionWrapper> GetEnumerator()
    {
        foreach (SectionWrapper section in GetList())
        {
            yield return section;
        }
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
