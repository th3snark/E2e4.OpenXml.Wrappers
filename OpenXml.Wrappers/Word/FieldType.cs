using System;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The kind of a field, as <c>WdFieldType</c> in Word (with the same values). Fields of kinds that are not listed report <see cref="Other"/>.
/// </summary>
public enum FieldType
{
    /// <summary>A field of a kind that is not listed.</summary>
    Other = -2,

    /// <summary>With <c>Fields.Add</c>: the text is the whole field code.</summary>
    Empty = -1,
    Ref = 3,
    IndexEntry = 4,
    FootnoteRef = 5,
    Set = 6,
    If = 7,
    Index = 8,
    TocEntry = 9,
    StyleRef = 10,
    Sequence = 12,
    TableOfContents = 13,
    Info = 14,
    Title = 15,
    Subject = 16,
    Author = 17,
    Keywords = 18,
    Comments = 19,
    LastSavedBy = 20,
    CreateDate = 21,
    SaveDate = 22,
    PrintDate = 23,
    RevisionNumber = 24,
    EditTime = 25,
    NumPages = 26,
    NumWords = 27,
    NumChars = 28,
    FileName = 29,
    Template = 30,
    Date = 31,
    Time = 32,
    Page = 33,
    Expression = 34,
    Quote = 35,
    Include = 36,
    PageRef = 37,
    Ask = 38,
    FillIn = 39,
    Data = 40,
    Next = 41,
    NextIf = 42,
    SkipIf = 43,
    MergeRecord = 44,
    Dde = 45,
    DdeAuto = 46,
    Glossary = 47,
    Print = 48,
    Formula = 49,
    GoToButton = 50,
    MacroButton = 51,
    AutoNumOutline = 52,
    AutoNumLegal = 53,
    AutoNum = 54,
    Import = 55,
    Link = 56,
    Symbol = 57,
    Embed = 58,
    MergeField = 59,
    UserName = 60,
    UserInitials = 61,
    UserAddress = 62,
    BarCode = 63,
    DocVariable = 64,
    Section = 65,
    SectionPages = 66,
    IncludePicture = 67,
    IncludeText = 68,
    FileSize = 69,
    FormTextInput = 70,
    FormCheckBox = 71,
    NoteRef = 72,
    TableOfAuthorities = 73,
    TableOfAuthoritiesEntry = 74,
    MergeSequence = 75,
    Private = 77,
    Database = 78,
    AutoText = 79,
    Compare = 80,
    Addin = 81,
    Subscriber = 82,
    FormDropDown = 83,
    Advance = 84,
    DocProperty = 85,
    Hyperlink = 88,
    AutoTextList = 89,
    ListNumber = 90,
    BidiOutline = 92,
    AddressBlock = 93,
    GreetingLine = 94
}

/// <summary>
/// The names that fields have in their codes.
/// </summary>
internal static class FieldCodes
{
    private static readonly (string Name, FieldType Type)[] Names =
    {
        ("REF", FieldType.Ref), ("XE", FieldType.IndexEntry), ("FTNREF", FieldType.FootnoteRef), ("SET", FieldType.Set), ("IF", FieldType.If),
        ("INDEX", FieldType.Index), ("TC", FieldType.TocEntry), ("STYLEREF", FieldType.StyleRef), ("SEQ", FieldType.Sequence),
        ("TOC", FieldType.TableOfContents), ("INFO", FieldType.Info), ("TITLE", FieldType.Title), ("SUBJECT", FieldType.Subject),
        ("AUTHOR", FieldType.Author), ("KEYWORDS", FieldType.Keywords), ("COMMENTS", FieldType.Comments), ("LASTSAVEDBY", FieldType.LastSavedBy),
        ("CREATEDATE", FieldType.CreateDate), ("SAVEDATE", FieldType.SaveDate), ("PRINTDATE", FieldType.PrintDate), ("REVNUM", FieldType.RevisionNumber),
        ("EDITTIME", FieldType.EditTime), ("NUMPAGES", FieldType.NumPages), ("NUMWORDS", FieldType.NumWords), ("NUMCHARS", FieldType.NumChars),
        ("FILENAME", FieldType.FileName), ("TEMPLATE", FieldType.Template), ("DATE", FieldType.Date), ("TIME", FieldType.Time), ("PAGE", FieldType.Page),
        ("QUOTE", FieldType.Quote), ("INCLUDE", FieldType.Include), ("PAGEREF", FieldType.PageRef), ("ASK", FieldType.Ask), ("FILLIN", FieldType.FillIn),
        ("DATA", FieldType.Data), ("NEXT", FieldType.Next), ("NEXTIF", FieldType.NextIf), ("SKIPIF", FieldType.SkipIf), ("MERGEREC", FieldType.MergeRecord),
        ("DDE", FieldType.Dde), ("DDEAUTO", FieldType.DdeAuto), ("GLOSSARY", FieldType.Glossary), ("PRINT", FieldType.Print), ("=", FieldType.Formula),
        ("GOTOBUTTON", FieldType.GoToButton), ("MACROBUTTON", FieldType.MacroButton), ("AUTONUMOUT", FieldType.AutoNumOutline),
        ("AUTONUMLGL", FieldType.AutoNumLegal), ("AUTONUM", FieldType.AutoNum), ("IMPORT", FieldType.Import), ("LINK", FieldType.Link),
        ("SYMBOL", FieldType.Symbol), ("EMBED", FieldType.Embed), ("MERGEFIELD", FieldType.MergeField), ("USERNAME", FieldType.UserName),
        ("USERINITIALS", FieldType.UserInitials), ("USERADDRESS", FieldType.UserAddress), ("BARCODE", FieldType.BarCode), ("DOCVARIABLE", FieldType.DocVariable),
        ("SECTION", FieldType.Section), ("SECTIONPAGES", FieldType.SectionPages), ("INCLUDEPICTURE", FieldType.IncludePicture),
        ("INCLUDETEXT", FieldType.IncludeText), ("FILESIZE", FieldType.FileSize), ("FORMTEXT", FieldType.FormTextInput),
        ("FORMCHECKBOX", FieldType.FormCheckBox), ("NOTEREF", FieldType.NoteRef), ("TOA", FieldType.TableOfAuthorities),
        ("TA", FieldType.TableOfAuthoritiesEntry), ("MERGESEQ", FieldType.MergeSequence), ("PRIVATE", FieldType.Private), ("DATABASE", FieldType.Database),
        ("AUTOTEXT", FieldType.AutoText), ("COMPARE", FieldType.Compare), ("ADDIN", FieldType.Addin), ("SUBSCRIBER", FieldType.Subscriber),
        ("FORMDROPDOWN", FieldType.FormDropDown), ("ADVANCE", FieldType.Advance), ("DOCPROPERTY", FieldType.DocProperty),
        ("HYPERLINK", FieldType.Hyperlink), ("AUTOTEXTLIST", FieldType.AutoTextList), ("LISTNUM", FieldType.ListNumber),
        ("BIDIOUTLINE", FieldType.BidiOutline), ("ADDRESSBLOCK", FieldType.AddressBlock), ("GREETINGLINE", FieldType.GreetingLine)
    };

    /// <summary>The type of field with the name in its code (ignoring case), or <see cref="FieldType.Other"/>.</summary>
    public static FieldType FromName(string name)
    {
        foreach ((string known, FieldType type) in Names)
        {
            if (string.Equals(known, name, StringComparison.OrdinalIgnoreCase))
            {
                return type;
            }
        }
        return FieldType.Other;
    }

    /// <summary>The name of the type of field in a code.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The type has no name (<see cref="FieldType.Other"/> and <see cref="FieldType.Empty"/>).</exception>
    public static string ToName(FieldType type)
    {
        foreach ((string name, FieldType known) in Names)
        {
            if (known == type)
            {
                return name;
            }
        }
        throw new ArgumentOutOfRangeException(nameof(type));
    }
}
