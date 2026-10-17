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

    internal void InsertFieldRuns(Func<Paragraph, List<Run>> build)
    {
        InsertRuns(build, true);
    }
}
