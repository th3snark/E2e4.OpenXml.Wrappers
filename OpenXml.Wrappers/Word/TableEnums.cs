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
    Dashed = 3,
    DotDash = 4,
    DoubleLine = 7,
    Triple = 8,
    Wave = 18,
    Inset = 21,
    Outset = 22,
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
