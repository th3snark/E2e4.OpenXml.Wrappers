namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a list template is applied to, as <c>WdListApplyTo</c> in Word (with the same values).
/// </summary>
public enum ListApplyTo
{
    /// <summary>The whole list the range is in; for a range that is not in a list, the range.</summary>
    WholeList = 0,

    /// <summary>From the start of the range to the end of the list. Not supported.</summary>
    ThisPointForward = 1,

    /// <summary>Only the range.</summary>
    Selection = 2
}

/// <summary>
/// How the indents of a list are worked out, as <c>WdDefaultListBehavior</c> in Word (with the same values). Only the behaviour of Word 2002
/// and later is implemented, whichever value is given.
/// </summary>
public enum DefaultListBehavior
{
    /// <summary>Word 97.</summary>
    Word97 = 0,

    /// <summary>Word 2000.</summary>
    Word2000 = 1,

    /// <summary>Word 2002 and later.</summary>
    Word10 = 2
}

/// <summary>
/// Whether a list template can carry on the list before a range, as <c>WdContinue</c> in Word (with the same values).
/// </summary>
public enum ListContinue
{
    /// <summary>There is no list before the range.</summary>
    Disabled = 0,

    /// <summary>The numbering can be started again.</summary>
    ResetList = 1,

    /// <summary>The list before the range can be carried on.</summary>
    ContinueList = 2
}
