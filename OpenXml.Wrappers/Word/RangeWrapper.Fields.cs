using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// The fields in the range.
    /// </summary>
    public FieldWrapperCollection Fields
    {
        get { return new FieldWrapperCollection(this); }
    }

    //Inserts the runs of a field at the end of the range. In a span of characters they have the formatting of the text before them.
    internal void InsertFieldRuns(Func<Paragraph, List<Run>> build)
    {
        if (_isSpan)
        {
            GetScoped(false, out StoryMap map, out int start, out int end);
            RunProperties formatting = FindFormatting(map, end, end);
            InsertRuns(build, true, formatting);
            return;
        }
        InsertRuns(build, true, null);
    }
}
