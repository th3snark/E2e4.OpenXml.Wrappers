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
/// The horizontal alignment of a table, as <c>WdRowAlignment</c> in Word.
/// </summary>
public enum RowAlignment
{
    Left = 0,
    Center = 1,
    Right = 2
}
