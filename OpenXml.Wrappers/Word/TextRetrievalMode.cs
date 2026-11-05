using System;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The view of a document, as <c>WdViewType</c> in Word.
/// </summary>
public enum ViewType
{
    /// <summary>Normal (draft) view.</summary>
    Normal = 1,

    /// <summary>Outline view.</summary>
    Outline = 2,

    /// <summary>Print layout view. It is the default.</summary>
    Print = 3,

    /// <summary>Print preview.</summary>
    PrintPreview = 4,

    /// <summary>Master document view.</summary>
    Master = 5,

    /// <summary>Web layout view.</summary>
    Web = 6,

    /// <summary>Reading layout view.</summary>
    Reading = 7,

    /// <summary>Conflict view.</summary>
    Conflict = 8
}

/// <summary>
/// What <see cref="RangeWrapper.Text"/> of one range includes, as <c>TextRetrievalMode</c> in Word. A new range has the defaults: field results
/// and hidden text, no field codes. A range made from another with <see cref="RangeWrapper.Duplicate"/> keeps the settings, and a range
/// that is moved or collapsed is the same range, so it does too.
/// </summary>
public sealed class TextRetrievalMode
{
    private ViewType _viewType = ViewType.Print;

    /// <summary>
    /// True to include the codes of fields instead of their results: <c>U+0013</c>, the code and <c>U+0015</c> for each field, without its
    /// separator and result. False (the default) includes the results.
    /// </summary>
    public bool IncludeFieldCodes { get; set; }

    /// <summary>True (the default) to include text that is hidden; false to leave it out.</summary>
    public bool IncludeHiddenText { get; set; } = true;

    /// <summary>
    /// The view the text is read in, as <c>TextRetrievalMode.ViewType</c> in Word. It is kept but does not change the text, which does not
    /// depend on a view here.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not a view.</exception>
    public ViewType ViewType
    {
        get { return _viewType; }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            _viewType = value;
        }
    }

    /// <summary>A copy of the settings.</summary>
    public TextRetrievalMode Duplicate()
    {
        return new TextRetrievalMode() { IncludeFieldCodes = IncludeFieldCodes, IncludeHiddenText = IncludeHiddenText, _viewType = _viewType };
    }

    internal bool ChangesText
    {
        get { return IncludeFieldCodes || !IncludeHiddenText; }
    }
}
