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

/// <summary>
/// The styles that come with Word, as <c>WdBuiltinStyle</c> in Word (with the same values). Only the styles that the library can define are
/// listed; they are added to a document when they are asked for with <see cref="StyleWrapperCollection"/> or applied by name.
/// </summary>
public enum BuiltInStyle
{
    Normal = -1,
    Heading1 = -2,
    Heading2 = -3,
    Heading3 = -4,
    Heading4 = -5,
    Heading5 = -6,
    Heading6 = -7,
    Heading7 = -8,
    Heading8 = -9,
    Heading9 = -10,
    Toc1 = -20,
    Toc2 = -21,
    Toc3 = -22,
    Toc4 = -23,
    Toc5 = -24,
    Toc6 = -25,
    Toc7 = -26,
    Toc8 = -27,
    Toc9 = -28,
    FootnoteText = -30,
    CommentText = -31,
    Header = -32,
    Footer = -33,
    Caption = -35,
    FootnoteReference = -39,
    CommentReference = -40,
    LineNumber = -41,
    PageNumber = -42,
    EndnoteReference = -43,
    EndnoteText = -44,
    Title = -63,
    Closing = -64,
    Signature = -65,
    DefaultParagraphFont = -66,
    BodyText = -67,
    BodyTextIndent = -68,
    Subtitle = -75,
    Salutation = -76,
    Date = -77,
    BodyText2 = -81,
    BodyText3 = -82,
    Hyperlink = -86,
    HyperlinkFollowed = -87,
    Strong = -88,
    Emphasis = -89,
    PlainText = -91,
    NormalTable = -106,
    ListParagraph = -180,
    Quote = -181,
    IntenseQuote = -182,
    SubtleEmphasis = -261,
    IntenseEmphasis = -262,
    SubtleReference = -263,
    IntenseReference = -264,
    BookTitle = -265
}
