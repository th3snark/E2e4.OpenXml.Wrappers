namespace OpenXml.Wrappers.Word;

/// <summary>
/// The orientation of a page, as <c>WdOrientation</c> in Word.
/// </summary>
public enum PageOrientation
{
    Portrait = 0,
    Landscape = 1
}

/// <summary>
/// Where a section starts, as <c>WdSectionStart</c> in Word.
/// </summary>
public enum SectionStart
{
    Continuous = 0,
    NewColumn = 1,
    NewPage = 2,
    EvenPage = 3,
    OddPage = 4
}

/// <summary>
/// The type of a style, as <c>WdStyleType</c> in Word.
/// </summary>
public enum StyleType
{
    Paragraph = 1,
    Character = 2,
    Table = 3,
    List = 4
}

/// <summary>
/// Which header or footer of a section, as <c>WdHeaderFooterIndex</c> in Word.
/// </summary>
public enum HeaderFooterIndex
{
    Primary = 1,
    FirstPage = 2,
    EvenPages = 3
}
