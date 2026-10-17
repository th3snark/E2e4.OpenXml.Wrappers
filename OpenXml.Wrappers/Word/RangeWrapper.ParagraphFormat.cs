namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// The paragraph formatting of the paragraphs that the range touches.
    /// </summary>
    public ParagraphFormatWrapper ParagraphFormat
    {
        get { return new ParagraphFormatWrapper(this); }
    }
}
