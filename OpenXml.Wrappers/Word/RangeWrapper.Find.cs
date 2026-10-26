using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    private FindWrapper _find;

    /// <summary>
    /// Finds, and optionally replaces, text in the range. As in Word the same object is returned each time so that its properties can be set
    /// in one statement and used in the next.
    /// </summary>
    public FindWrapper Find
    {
        get
        {
            if (_find == null)
            {
                _find = new FindWrapper(this);
            }
            return _find;
        }
    }

    /// <summary>
    /// Returns the runs of the range grouped by the paragraph that holds them, in document order. Runs that belong to a paragraph nested in
    /// another one (for example in a text box) are grouped with that nested paragraph.
    /// </summary>
    internal IList<List<Run>> GetRunsByParagraph()
    {
        if (_isSpan)
        {
            return GetSpanRunsByParagraph();
        }

        if (_startOpenXmlElement is BookmarkStart || _startOpenXmlElement is CommentRangeStart)
        {
            return _startOpenXmlElement.GetElements<Run>(_endOpenXmlElement)
                .GroupBy(r => r.Ancestors<Paragraph>().First())
                .Select(g => g.ToList())
                .ToList();
        }

        if (_startOpenXmlElement is Paragraph || _startOpenXmlElement is TableCell || _startOpenXmlElement is TableRow || _startOpenXmlElement is Table
            || _startOpenXmlElement is Document || _startOpenXmlElement is Body || _startOpenXmlElement is Comment
            || _startOpenXmlElement is Header || _startOpenXmlElement is Footer)
        {
            return _startOpenXmlElement.GetElements<Paragraph>()
                .Select(p => p.Descendants<Run>().Where(r => r.Ancestors<Paragraph>().First() == p).ToList())
                .Where(runs => runs.Count > 0)
                .ToList();
        }

        throw new NotSupportedException(string.Format(CultureInfo.InvariantCulture, "Finding text is not currently supported for a range starting with the type {0}", _startOpenXmlElement.GetType().Name));
    }
}
