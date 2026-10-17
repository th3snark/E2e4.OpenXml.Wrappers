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
        ArgumentNullException.ThrowIfNull(body);
        _body = body;
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
