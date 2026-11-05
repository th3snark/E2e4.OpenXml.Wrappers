namespace OpenXml.Wrappers.Word;

public partial class RangeWrapper
{
    /// <summary>
    /// The paragraph formatting of the paragraphs that the range touches. Assigning a format (usually a <c>Duplicate</c> of one) gives them the
    /// formatting it has; properties it has no single value for are left as they are.
    /// </summary>
    public ParagraphFormatWrapper ParagraphFormat
    {
        get { return new ParagraphFormatWrapper(this); }
        set
        {
            Guard.ThrowIfNull(value, nameof(value));
            value.CopyTo(new ParagraphFormatWrapper(this));
        }
    }
}
