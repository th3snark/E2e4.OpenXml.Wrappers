namespace OpenXml.Wrappers.Word;

/// <summary>
/// The kind of a field, as <c>WdFieldType</c> in Word. Fields of other kinds report <see cref="Other"/>.
/// </summary>
public enum FieldType
{
    Other = -1,
    Ref = 3,
    If = 7,
    StyleRef = 10,
    Sequence = 12,
    TableOfContents = 13,
    Author = 17,
    NumPages = 26,
    FileName = 29,
    Date = 31,
    Time = 32,
    Page = 33,
    PageRef = 37,
    MergeField = 59,
    Title = 72,
    Hyperlink = 88
}
