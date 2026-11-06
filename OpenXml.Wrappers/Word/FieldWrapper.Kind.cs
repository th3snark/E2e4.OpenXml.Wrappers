namespace OpenXml.Wrappers.Word;

/// <summary>
/// What a field does, as <c>WdFieldKind</c> in Word (with the same values).
/// </summary>
public enum FieldKind
{
    /// <summary>Not a field with a kind.</summary>
    None = 0,

    /// <summary>A field that does something when it is clicked or shows a result without being updated, such as MACROBUTTON, SYMBOL or AUTONUM.</summary>
    Hot = 1,

    /// <summary>A field with a result that is worked out when it is updated, such as PAGE, DATE or REF.</summary>
    Warm = 2,

    /// <summary>A field with no result that only marks a place, such as XE or TC.</summary>
    Cold = 3
}

public partial class FieldWrapper
{
    /// <summary>
    /// What the field does, as <c>Field.Kind</c> in Word: <see cref="FieldKind.Hot"/> for MACROBUTTON, GOTOBUTTON, SYMBOL, advance, the
    /// automatic numbers, BARCODE and formulas; <see cref="FieldKind.Cold"/> for the entries of an index, a table of contents or authorities, and
    /// PRIVATE; <see cref="FieldKind.Warm"/> for the rest (checked in Word 16 for the types that Word lists).
    /// </summary>
    public FieldKind Kind
    {
        get
        {
            FieldType type = Type;
            if (type == FieldType.Empty)
            {
                return FieldKind.None;
            }
            if (type == FieldType.IndexEntry || type == FieldType.TocEntry || type == FieldType.TableOfAuthoritiesEntry || type == FieldType.Private)
            {
                return FieldKind.Cold;
            }
            if (type == FieldType.Formula || type == FieldType.GoToButton || type == FieldType.MacroButton || type == FieldType.AutoNumOutline
                || type == FieldType.AutoNumLegal || type == FieldType.AutoNum || type == FieldType.Symbol || type == FieldType.BarCode
                || type == FieldType.Advance || type == FieldType.ListNumber || type == FieldType.BidiOutline)
            {
                return FieldKind.Hot;
            }
            return FieldKind.Warm;
        }
    }
}
