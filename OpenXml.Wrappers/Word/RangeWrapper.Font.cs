using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// The character formatting of the range.
    /// </summary>
    public FontWrapper Font
    {
        get { return new FontWrapper(this); }
    }

    /// <summary>
    /// Returns the runs of the range in document order.
    /// </summary>
    internal IList<Run> GetRuns()
    {
        return GetRunsByParagraph().SelectMany(runs => runs).ToList();
    }

    /// <summary>
    /// Returns the paragraphs whose paragraph marks are in the range: those of a range of whole paragraphs, and none for a bookmark or comment
    /// range, which covers part of a paragraph.
    /// </summary>
    internal IList<Paragraph> GetParagraphs()
    {
        if (_startOpenXmlElement is Paragraph || _startOpenXmlElement is TableCell || _startOpenXmlElement is TableRow || _startOpenXmlElement is Table
            || _startOpenXmlElement is Document || _startOpenXmlElement is Body || _startOpenXmlElement is Comment)
        {
            return _startOpenXmlElement.GetElements<Paragraph>().ToList();
        }

        return new List<Paragraph>();
    }
}
