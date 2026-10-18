namespace OpenXml.Wrappers.Word;

/// <summary>
/// How the height of a row is applied, as <c>WdRowHeightRule</c> in Word.
/// </summary>
public enum RowHeightRule
{
    Auto = 0,
    AtLeast = 1,
    Exactly = 2
}

/// <summary>
/// The vertical alignment of the text in a cell, as <c>WdCellVerticalAlignment</c> in Word.
/// </summary>
public enum CellVerticalAlignment
{
    Top = 0,
    Center = 1,
    Bottom = 3
}

/// <summary>
/// The border of a table or cell, as <c>WdBorderType</c> in Word.
/// </summary>
public enum BorderPosition
{
    Top = -1,
    Left = -2,
    Bottom = -3,
    Right = -4,
    Horizontal = -5,
    Vertical = -6,
    DiagonalDown = -7,
    DiagonalUp = -8
}

/// <summary>
/// The line style of a border, as <c>WdLineStyle</c> in Word. <see cref="Other"/> is what is read for styles that have no
/// member here; it cannot be set.
/// </summary>
public enum BorderLineStyle
{
    None = 0,
    SingleLine = 1,
    Dotted = 2,
    DashSmallGap = 3,
    DashLargeGap = 4,
    DashDot = 5,
    DashDotDot = 6,
    DoubleLine = 7,
    Triple = 8,
    ThinThickSmallGap = 9,
    ThickThinSmallGap = 10,
    ThinThickThinSmallGap = 11,
    ThinThickMediumGap = 12,
    ThickThinMediumGap = 13,
    ThinThickThinMediumGap = 14,
    ThinThickLargeGap = 15,
    ThickThinLargeGap = 16,
    ThinThickThinLargeGap = 17,
    SingleWavy = 18,
    DoubleWavy = 19,
    DashDotStroked = 20,
    Emboss3D = 21,
    Engrave3D = 22,
    Outset = 23,
    Inset = 24,
    Other = -1
}

/// <summary>
/// The horizontal alignment of a table, as <c>WdRowAlignment</c> in Word.
/// </summary>
public enum RowAlignment
{
    Left = 0,
    Center = 1,
    Right = 2
}

/// <summary>
/// How a table is made, as <c>WdDefaultTableBehavior</c> in Word (with the same values).
/// </summary>
public enum DefaultTableBehavior
{
    /// <summary>Fixed column widths.</summary>
    Word8TableBehavior = 0,

    /// <summary>The columns can resize to fit their content.</summary>
    Word9TableBehavior = 1
}

/// <summary>
/// How a table is sized, as <c>WdAutoFitBehavior</c> in Word (with the same values).
/// </summary>
public enum AutoFitBehavior
{
    /// <summary>Fixed column widths.</summary>
    Fixed = 0,

    /// <summary>The table is as wide as its content needs.</summary>
    Content = 1,

    /// <summary>The table is as wide as the window (the page between the margins).</summary>
    Window = 2
}

/// <summary>
/// What the preferred width of a table, column or cell means, as <c>WdPreferredWidthType</c> in Word (with the same values).
/// </summary>
public enum PreferredWidthType
{
    /// <summary>The width is worked out from the content.</summary>
    Auto = 1,

    /// <summary>The width is a percentage of the width available.</summary>
    Percent = 2,

    /// <summary>The width is in points.</summary>
    Points = 3
}
