using System;
using System.Collections;
using System.Collections.Generic;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The headers or the footers of a section, as <c>HeadersFooters</c> in Word: one each for the primary, first page and even pages.
/// </summary>
/// <remarks>
/// Unlike the other collections this is indexed by a <see cref="HeaderFooterIndex"/>, as in Word, where the members are 1, 2 and 3 and not
/// positions.
/// </remarks>
public class HeaderFooterWrapperCollection : IEnumerable<HeaderFooterWrapper>
{
    private readonly SectionWrapper _section;
    private readonly bool _isHeader;

    internal HeaderFooterWrapperCollection(SectionWrapper section, bool isHeader)
    {
        _section = section;
        _isHeader = isHeader;
    }

    public SectionWrapper Parent
    {
        get { return _section; }
    }

    public HeaderFooterWrapper this[HeaderFooterIndex index]
    {
        get
        {
            if (!Enum.IsDefined(index))
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            return new HeaderFooterWrapper(_section, _isHeader, index);
        }
    }

    public int Count
    {
        get { return 3; }
    }

    #region IEnumerable Members

    public IEnumerator<HeaderFooterWrapper> GetEnumerator()
    {
        yield return this[HeaderFooterIndex.Primary];
        yield return this[HeaderFooterIndex.FirstPage];
        yield return this[HeaderFooterIndex.EvenPages];
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion
}
