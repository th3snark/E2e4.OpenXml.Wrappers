using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The code or result of a field, as the <c>Range</c> that <c>Field.Code</c> and <c>Field.Result</c> are in Word: it has the <see cref="Text"/>
/// and a <see cref="Range"/> (a span of the characters of the code or result), and <see cref="Font"/> is the font of that range.
/// </summary>
public class FieldPartWrapper
{
    private readonly Func<string> _get;
    private readonly Action<string> _set;
    private readonly Func<RangeWrapper> _range;

    internal FieldPartWrapper(Func<string> get, Action<string> set, Func<RangeWrapper> range)
    {
        _get = get;
        _set = set;
        _range = range;
    }

    /// <summary>The text. Setting it replaces the code (the result), keeping the formatting of its first run.</summary>
    public string Text
    {
        get { return _get(); }
        set { _set(value); }
    }

    /// <summary>
    /// The characters of the code or result as a range, which is empty when there are none. It is a snapshot of the characters there are now:
    /// after setting <see cref="Text"/> get the range again.
    /// </summary>
    public RangeWrapper Range
    {
        get { return _range(); }
    }

    /// <summary>The font of the characters of the code or result (<c>Field.Result.Font</c> in Word).</summary>
    public FontWrapper Font
    {
        get { return Range.Font; }
    }
}
public partial class FieldWrapper
{
    /// <summary>
    /// The field code, such as <c>PAGE \* MERGEFORMAT</c>, without its surrounding spaces (<c>Field.Code</c> in Word).
    /// </summary>
    public FieldPartWrapper Code
    {
        get { return new FieldPartWrapper(() => CodeText, value => CodeText = value, () => GetPartRange(true)); }
    }

    /// <summary>
    /// The result of the field: the text stored in the document, which Word shows until the field is updated (<c>Field.Result</c> in Word).
    /// </summary>
    /// <exception cref="NotSupportedException">Setting the text when the result is not within a single paragraph.</exception>
    public FieldPartWrapper Result
    {
        get { return new FieldPartWrapper(() => ResultText, value => ResultText = value, () => GetPartRange(false)); }
    }

    /// <summary>
    /// Whether the field is locked, so that it is not updated (<c>Field.Locked</c> in Word).
    /// </summary>
    public bool Locked
    {
        get
        {
            OnOffValue locked = _simpleField != null ? _simpleField.FieldLock : _begin.FieldLock;
            return locked != null && locked.HasValue && locked.Value;
        }
        set
        {
            if (_simpleField != null)
            {
                _simpleField.FieldLock = value ? (OnOffValue)true : null;
            }
            else
            {
                _begin.FieldLock = value ? (OnOffValue)true : null;
            }
        }
    }

    /// <summary>
    /// The 0-based position of the field among the fields of its story (<c>Field.Index</c> in Word, which counts from 1).
    /// </summary>
    public int Index
    {
        get { return GetSiblings().IndexOf(this); }
    }

    private List<FieldWrapper> GetSiblings()
    {
        return new FieldWrapperCollection(new RangeWrapper(Key.GetRoot())).ToList();
    }

    /// <summary>The next field in the story, or null for the last one (<c>Field.Next</c> in Word).</summary>
    public FieldWrapper Next()
    {
        List<FieldWrapper> fields = GetSiblings();
        int index = fields.IndexOf(this);
        return index >= 0 && index + 1 < fields.Count ? fields[index + 1] : null;
    }

    /// <summary>The previous field in the story, or null for the first one (<c>Field.Previous</c> in Word).</summary>
    public FieldWrapper Previous()
    {
        List<FieldWrapper> fields = GetSiblings();
        int index = fields.IndexOf(this);
        return index > 0 ? fields[index - 1] : null;
    }

    /// <summary>
    /// Works out the result of the field again, as <c>Field.Update</c> in Word, and returns whether it did. Fields that do not depend on how
    /// the pages are laid out are worked out: the date and time, document properties (author, title, DOCPROPERTY and so on), file name, user
    /// name, number of words and characters, SEQ, REF, STYLEREF (in the main text), IF, QUOTE and formulas (<c>=</c>, in a table or not). The switches for the format of the result (<c>\@</c> for dates and
    /// <c>\*</c> for case and numbers) are applied. A locked field, and fields such as PAGE, NUMPAGES, PAGEREF, TOC and MERGEFIELD, are left as
    /// they are (the latter are marked for Word to update when the document is opened) and report false.
    /// </summary>
    public bool Update()
    {
        if (Locked)
        {
            return false;
        }

        string result = FieldUpdater.Compute(this, CodeText, out bool ok);
        if (result == null)
        {
            return false;
        }

        ResultText = result;
        if (_simpleField != null)
        {
            _simpleField.Dirty = null;
        }
        else
        {
            _begin.Dirty = null;
        }
        return ok;
    }

    /// <summary>
    /// Replaces the field by its result as ordinary text, as <c>Field.Unlink</c> in Word.
    /// </summary>
    public void Unlink()
    {
        if (_simpleField != null)
        {
            foreach (OpenXmlElement child in _simpleField.ChildElements.ToList())
            {
                child.Remove();
                _simpleField.InsertBeforeSelf(child);
            }
            _simpleField.Remove();
            return;
        }

        Parts parts = Parse();
        Run beginRun = (Run)_begin.Parent;
        Run separateRun = parts.Separate == null ? null : (Run)parts.Separate.Parent;
        Run endRun = parts.End == null ? null : (Run)parts.End.Parent;

        //Everything up to the separator is the code, and the separator and the end go too. Without a separator there is no result.
        bool inCode = true;
        foreach (Run run in parts.AllRuns.ToList())
        {
            bool isSeparate = run == separateRun;
            bool isEnd = run == endRun;
            if (inCode && !isSeparate && !(separateRun == null && isEnd))
            {
                run.Remove();
                continue;
            }

            if (isSeparate)
            {
                parts.Separate.Remove();
                inCode = false;
            }
            if (isEnd)
            {
                parts.End.Remove();
            }
            if (!run.Elements().Any(e => !(e is RunProperties)))
            {
                run.Remove();
            }
        }

        GC.KeepAlive(beginRun);
    }
}

/// <summary>
/// Works out the results of fields that do not depend on the layout of the pages.
/// </summary>
internal static class FieldUpdater
{
    internal static void RememberPath(WordprocessingDocument document, string path)
    {
        DocumentFiles.Remember(document, path);
    }

    private sealed class Instruction
    {
        public string Name;
        public readonly List<string> Arguments = new List<string>();
        public readonly List<KeyValuePair<string, string>> Switches = new List<KeyValuePair<string, string>>();

        public string GetSwitch(string key)
        {
            foreach (KeyValuePair<string, string> pair in Switches)
            {
                if (pair.Key == key)
                {
                    return pair.Value ?? string.Empty;
                }
            }
            return null;
        }
    }

    //Splits the code into the name, the arguments and the switches. A switch can have an argument: \@ "format", \* Upper, \r 5.
    private static Instruction Parse(string code)
    {
        List<string> tokens = new List<string>();
        StringBuilder current = new StringBuilder();
        bool quoted = false;
        bool hasToken = false;
        foreach (char c in code ?? string.Empty)
        {
            if (c == '"' || c == '“' || c == '”')
            {
                quoted = !quoted;
                hasToken = true;
            }
            else if (char.IsWhiteSpace(c) && !quoted)
            {
                if (hasToken)
                {
                    tokens.Add(current.ToString());
                    current.Length = 0;
                    hasToken = false;
                }
            }
            else
            {
                current.Append(c);
                hasToken = true;
            }
        }
        if (hasToken)
        {
            tokens.Add(current.ToString());
        }

        Instruction instruction = new Instruction() { Name = tokens.Count > 0 ? tokens[0].ToUpperInvariant() : string.Empty };
        for (int i = 1; i < tokens.Count; i++)
        {
            string token = tokens[i];
            if (token.Length >= 2 && token[0] == '\\')
            {
                string key = token.Substring(1);
                string value = null;
                bool takesArgument = key == "@" || key == "*" || key == "#" || ((key == "r" || key == "s") && i + 1 < tokens.Count && int.TryParse(tokens[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out _));
                if (takesArgument && i + 1 < tokens.Count)
                {
                    value = tokens[++i];
                }
                instruction.Switches.Add(new KeyValuePair<string, string>(key, value));
            }
            else
            {
                instruction.Arguments.Add(token);
            }
        }
        return instruction;
    }

    /// <summary>
    /// Returns the new result of the field, or null when it cannot be worked out. <paramref name="ok"/> is false when a result is returned
    /// that says the field is in error, such as a reference to a bookmark that is not there.
    /// </summary>
    public static string Compute(FieldWrapper field, string code, out bool ok)
    {
        ok = true;
        Instruction instruction = Parse(code);
        MainDocumentPart main = StyleHelper.GetMainDocumentPart(field.Key);
        WordprocessingDocument document = main?.OpenXmlPackage as WordprocessingDocument;
        string text = ComputeText(field, instruction, main, document, ref ok);
        if (text == null)
        {
            return null;
        }
        //An error is shown as it is, without the format switches, as in Word.
        return ok ? ApplyFormat(text, instruction) : text;
    }

    private static string ComputeText(FieldWrapper field, Instruction instruction, MainDocumentPart main, WordprocessingDocument document, ref bool ok)
    {
        CultureInfo culture = CultureInfo.CurrentCulture;
        DocumentPropertiesWrapper properties = document == null ? null : new DocumentPropertiesWrapper(document);
        string picture = instruction.GetSwitch("@");

        if (instruction.Name.Length > 0 && instruction.Name[0] == '=')
        {
            string expression = instruction.Name + string.Concat(instruction.Arguments);
            TableCell cell = field.Key.Ancestors<TableCell>().FirstOrDefault();
            string result = TableFormulas.Compute(cell, expression, instruction.GetSwitch("#"))
                ?? FieldFormulas.Compute(expression.Substring(1), field.Key.GetRoot(), cell != null, instruction.GetSwitch("#"));
            //A number format of \* is not applied to an error of the formula here (how Word shows that is not measured), so such a field is
            //left as it is.
            if (result != null && result.Length > 0 && result[0] == '!'
                && instruction.Switches.Any(s => s.Key == "*" && FieldNumberFormats.IsNumberFormat(s.Value)))
            {
                return null;
            }
            return result;
        }

        switch (instruction.Name)
        {
            case "DATE":
                return FormatDate(DateTime.Now, picture ?? culture.DateTimeFormat.ShortDatePattern, culture);
            case "TIME":
                return FormatDate(DateTime.Now, picture ?? culture.DateTimeFormat.ShortTimePattern, culture);
            case "CREATEDATE":
                return properties?.CreationDate == null ? null : FormatDate(properties.CreationDate.Value.ToLocalTime(), picture ?? DefaultDateTime(culture), culture);
            case "SAVEDATE":
                return properties?.LastSaveTime == null ? null : FormatDate(properties.LastSaveTime.Value.ToLocalTime(), picture ?? DefaultDateTime(culture), culture);
            case "PRINTDATE":
                DateTime? printed = (document == null ? null : DocumentPropertiesWrapper.GetCoreProperties(document).LastPrinted);
                return printed == null ? "0/0/0000 0:00" : FormatDate(printed.Value.ToLocalTime(), picture ?? DefaultDateTime(culture), culture);
            case "AUTHOR":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Author ?? string.Empty;
            case "TITLE":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Title ?? string.Empty;
            case "SUBJECT":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Subject ?? string.Empty;
            case "KEYWORDS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Keywords ?? string.Empty;
            case "COMMENTS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Comments ?? string.Empty;
            case "LASTSAVEDBY":
                return properties?.LastAuthor ?? string.Empty;
            case "REVNUM":
                return (document == null ? null : DocumentPropertiesWrapper.GetCoreProperties(document).Revision) ?? "1";
            case "USERNAME":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : Environment.UserName;
            case "USERINITIALS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : Environment.UserName.Length > 0 ? Environment.UserName.Substring(0, 1).ToUpperInvariant() : string.Empty;
            case "USERADDRESS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : string.Empty;
            case "DOCPROPERTY":
                return instruction.Arguments.Count == 0 || document == null ? null : DocumentProperty(document, properties, instruction.Arguments[0], ref ok);
            case "FILENAME":
                if (document == null || !DocumentFiles.TryGetPath(document, out string path))
                {
                    return null;
                }
                return instruction.GetSwitch("p") != null ? System.IO.Path.GetFullPath(path) : System.IO.Path.GetFileName(path);
            case "NUMWORDS":
                return main == null ? null : CountWords(main).ToString(CultureInfo.InvariantCulture);
            case "NUMCHARS":
                return main == null ? null : CountCharacters(main).ToString(CultureInfo.InvariantCulture);
            case "DOCVARIABLE":
                if (instruction.Arguments.Count == 0 || main == null)
                {
                    return null;
                }
                DocumentVariable variable = main.DocumentSettingsPart?.Settings?.GetFirstChild<DocumentVariables>()?.Elements<DocumentVariable>()
                    .FirstOrDefault(v => string.Equals(v.Name?.Value, instruction.Arguments[0], StringComparison.OrdinalIgnoreCase));
                if (variable == null)
                {
                    ok = false;
                    return "Error! No document variable supplied.";
                }
                return variable.Val?.Value ?? string.Empty;
            case "QUOTE":
                return string.Join(" ", instruction.Arguments);
            case "SEQ":
                return Sequence(field, instruction);
            case "REF":
                return Reference(field, instruction, ref ok);
            case "STYLEREF":
                return StyleReference(field, instruction, ref ok);
            case "IF":
                return Condition(instruction);
            case "HYPERLINK":
                //There is no result to work out: the text shown is the result.
                return field.ResultText;
            default:
                return null;
        }
    }

    private static string DefaultDateTime(CultureInfo culture)
    {
        return culture.DateTimeFormat.ShortDatePattern + " " + culture.DateTimeFormat.ShortTimePattern;
    }

    #region Document properties

    private static string DocumentProperty(WordprocessingDocument document, DocumentPropertiesWrapper properties, string name, ref bool ok)
    {
        switch (name.Replace(" ", string.Empty, StringComparison.Ordinal).ToUpperInvariant())
        {
            case "TITLE":
                return properties.Title ?? string.Empty;
            case "SUBJECT":
                return properties.Subject ?? string.Empty;
            case "AUTHOR":
                return properties.Author ?? string.Empty;
            case "KEYWORDS":
                return properties.Keywords ?? string.Empty;
            case "COMMENTS":
                return properties.Comments ?? string.Empty;
            case "CATEGORY":
                return properties.Category ?? string.Empty;
            case "LASTAUTHOR":
            case "LASTSAVEDBY":
                return properties.LastAuthor ?? string.Empty;
            case "COMPANY":
                return properties.Company ?? string.Empty;
            case "MANAGER":
                return properties.Manager ?? string.Empty;
            case "CREATIONDATE":
            case "CREATEDATE":
                return properties.CreationDate?.ToLocalTime().ToString(CultureInfo.CurrentCulture) ?? string.Empty;
            case "LASTSAVETIME":
            case "SAVEDATE":
                return properties.LastSaveTime?.ToLocalTime().ToString(CultureInfo.CurrentCulture) ?? string.Empty;
        }

        CustomDocumentProperty custom = new CustomDocumentPropertyCollection(document)[name];
        if (custom == null)
        {
            ok = false;
            return "Error! Unknown document property name.";
        }
        object value = custom.Value;
        if (value is DateTime date)
        {
            return date.ToLocalTime().ToString(CultureInfo.CurrentCulture);
        }
        return Convert.ToString(value, CultureInfo.CurrentCulture) ?? string.Empty;
    }

    private static string GetStoryText(MainDocumentPart main)
    {
        return new RangeWrapper(main.Document).Text ?? string.Empty;
    }

    private static int CountWords(MainDocumentPart main)
    {
        return GetStoryText(main).Split(new[] { ' ', '\t', '\r', '\n', '\a', '\v', '\f', ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private static int CountCharacters(MainDocumentPart main)
    {
        return GetStoryText(main).Count(c => !char.IsWhiteSpace(c) && !char.IsControl(c));
    }

    #endregion

    #region SEQ, REF and IF

    private static string Sequence(FieldWrapper field, Instruction instruction)
    {
        if (instruction.Arguments.Count == 0)
        {
            return null;
        }

        string identifier = instruction.Arguments[0];
        OpenXmlElement root = field.Key.GetRoot();
        long value = 0;
        int previousParagraph = -1;
        Dictionary<Paragraph, int> paragraphIndexes = null;
        List<KeyValuePair<int, int>> headings = null;
        foreach (FieldWrapper other in new FieldWrapperCollection(new RangeWrapper(root)))
        {
            Instruction otherInstruction = ReferenceEquals(other, field) || other.Equals(field) ? instruction : Parse(other.CodeText);
            if (otherInstruction.Name == "SEQ" && otherInstruction.Arguments.Count > 0
                && string.Equals(otherInstruction.Arguments[0], identifier, StringComparison.OrdinalIgnoreCase))
            {
                if (paragraphIndexes == null)
                {
                    paragraphIndexes = new Dictionary<Paragraph, int>();
                    foreach (Paragraph paragraph in root.Descendants<Paragraph>())
                    {
                        paragraphIndexes[paragraph] = paragraphIndexes.Count;
                    }
                }
                Paragraph own = other.Key.Ancestors<Paragraph>().FirstOrDefault();
                int paragraphIndex = own != null && paragraphIndexes.TryGetValue(own, out int found) ? found : previousParagraph;
                value = NextSequenceValue(value, otherInstruction, previousParagraph, paragraphIndex, () => headings ?? (headings = Headings(paragraphIndexes)));
                previousParagraph = paragraphIndex;
            }

            if (other.Equals(field))
            {
                break;
            }
        }
        //The switch that hides the number leaves the result empty, unless a number format of \* shows it after all, as in Word. The \*
        //formats are applied to the number afterwards, as to any field's result.
        bool shown = instruction.GetSwitch("h") == null
            || instruction.Switches.Any(s => s.Key == "*" && FieldNumberFormats.IsNumberFormat(s.Value));
        return shown ? value.ToString(CultureInfo.InvariantCulture) : string.Empty;
    }

    //The number a SEQ field gives after the previous one of its sequence gave value, as Word works it out. \s n starts the count again after a
    //heading of level n or higher that begins after the previous field of the sequence (the field's own heading paragraph counts, and so does
    //an empty heading). Of \c (repeat) and \r (start at), the first one written wins, and Word ignores either when it comes after \s.
    private static long NextSequenceValue(long value, Instruction instruction, int previousParagraph, int paragraph, Func<List<KeyValuePair<int, int>>> headings)
    {
        int restartAt = instruction.Switches.FindIndex(s => s.Key == "s");
        if (restartAt >= 0 && int.TryParse(instruction.Switches[restartAt].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int level)
            && headings().Any(h => h.Value <= level && h.Key > previousParagraph && h.Key <= paragraph))
        {
            value = 0;
        }

        int last = restartAt >= 0 ? restartAt : instruction.Switches.Count;
        KeyValuePair<string, string> mode = instruction.Switches.Take(last).FirstOrDefault(s => s.Key == "c" || s.Key == "r");
        if (mode.Key == "c")
        {
            return value;
        }
        if (mode.Key == "r" && long.TryParse(mode.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number))
        {
            return number;
        }
        return value + 1;
    }

    //The paragraphs in the built-in heading styles, as their index in document order and their level (1 to 9).
    private static List<KeyValuePair<int, int>> Headings(Dictionary<Paragraph, int> paragraphIndexes)
    {
        List<KeyValuePair<int, int>> headings = new List<KeyValuePair<int, int>>();
        foreach (KeyValuePair<Paragraph, int> paragraph in paragraphIndexes)
        {
            string styleId = paragraph.Key.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            string name = styleId == null ? null : StyleHelper.GetStyleName(paragraph.Key, styleId);
            if (name != null && name.Length == 9 && name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase) && name[8] >= '1' && name[8] <= '9')
            {
                headings.Add(new KeyValuePair<int, int>(paragraph.Value, name[8] - '0'));
            }
        }
        return headings;
    }

    private static string Reference(FieldWrapper field, Instruction instruction, ref bool ok)
    {
        if (instruction.Arguments.Count == 0)
        {
            return null;
        }
        //\f (note numbers) and \d (the separator of the number) are left for Word.
        if (instruction.GetSwitch("f") != null || instruction.GetSwitch("d") != null)
        {
            return null;
        }
        if (!TryGetNumberSwitches(instruction, out ListNumbering.NumberContext? context, out bool withoutWords, out bool position))
        {
            return null;
        }

        OpenXmlElement root = field.Key.GetRoot();
        BookmarkStart bookmark = root.Descendants<BookmarkStart>().FirstOrDefault(b => string.Equals(b.Name?.Value, instruction.Arguments[0], StringComparison.OrdinalIgnoreCase));
        if (bookmark == null)
        {
            ok = false;
            return "Error! Reference source not found.";
        }

        string text;
        if (context.HasValue)
        {
            Paragraph target = bookmark.Ancestors<Paragraph>().FirstOrDefault() ?? bookmark.ElementsAfter().OfType<Paragraph>().FirstOrDefault();
            text = target == null ? null : ParagraphNumber(field, target, context.Value, withoutWords);
        }
        else
        {
            //With \p alone the field says only where the bookmark is.
            text = position ? string.Empty : new BookmarkWrapper(bookmark).Range.Text.TrimEnd('\r', '\a');
        }
        if (text == null || !position)
        {
            return text;
        }
        if (!ReferenceEquals(StoryMap.FindRoot(bookmark), StoryMap.FindRoot(field.Key)))
        {
            return null;
        }
        string where = bookmark.IsBefore(field.Key) ? "above" : "below";
        return text.Length == 0 ? where : text + " " + where;
    }

    //The list-number switches of REF and STYLEREF: at most one of \n, \r and \w (the context), \t (without words, only with one of them) and
    //\p (above or below). False for a combination that is left for Word.
    private static bool TryGetNumberSwitches(Instruction instruction, out ListNumbering.NumberContext? context, out bool withoutWords, out bool position)
    {
        List<ListNumbering.NumberContext> contexts = new List<ListNumbering.NumberContext>();
        foreach (KeyValuePair<string, ListNumbering.NumberContext> pair in new Dictionary<string, ListNumbering.NumberContext>
        {
            ["n"] = ListNumbering.NumberContext.None,
            ["r"] = ListNumbering.NumberContext.Relative,
            ["w"] = ListNumbering.NumberContext.Full,
        })
        {
            if (instruction.Switches.Any(s => s.Key == pair.Key))
            {
                contexts.Add(pair.Value);
            }
        }
        context = contexts.Count == 1 ? contexts[0] : null;
        withoutWords = instruction.Switches.Any(s => s.Key == "t");
        position = instruction.Switches.Any(s => s.Key == "p");
        return contexts.Count <= 1 && (!withoutWords || contexts.Count == 1);
    }

    //STYLEREF in the main text, as measured in Word: the nearest paragraph in the style, looking back from the field's own paragraph and then
    //forward, gives its text, or with \n, \r or \w its list number ("0" without one), and with \p "above" or "below". The field's own
    //paragraph gives no text or position (a self-reference error) but does give its number. "STYLEREF 1" is Heading 1, and \l is the same as
    //without it here. A character style, \s, and a field outside the main text (where it depends on the page) are left for Word.
    private static string StyleReference(FieldWrapper field, Instruction instruction, ref bool ok)
    {
        const string SelfReference = "Error! Not a valid bookmark self-reference.";
        if (instruction.Arguments.Count == 0 || instruction.GetSwitch("s") != null)
        {
            return null;
        }
        if (!TryGetNumberSwitches(instruction, out ListNumbering.NumberContext? context, out bool withoutWords, out bool position))
        {
            return null;
        }
        Paragraph own = field.Key.Ancestors<Paragraph>().FirstOrDefault();
        if (own == null || !(StoryMap.FindRoot(own) is Body body))
        {
            return null;
        }

        string name = instruction.Arguments[0];
        if (name.Length == 1 && name[0] >= '1' && name[0] <= '9')
        {
            name = "Heading " + name;
        }
        Styles styles = StyleHelper.GetStyles(own);
        //By name only, as in Word: "Heading1" (the id of Heading 1) is not a style to STYLEREF.
        Style style = styles?.Elements<Style>().FirstOrDefault(s => string.Equals(StyleHelper.GetName(s), name, StringComparison.OrdinalIgnoreCase));
        if (style == null)
        {
            ok = false;
            return BuiltInStyles.IsBuiltInName(name) && string.Equals(BuiltInStyles.GetDisplayName(name), name, StringComparison.OrdinalIgnoreCase)
                ? "Error! No text of specified style in document."
                : string.Format(CultureInfo.InvariantCulture, "Error! Use the Home tab to apply {0} to the text that you want to appear here.", name);
        }
        if (style.Type != null && style.Type.HasValue && style.Type.Value != StyleValues.Paragraph)
        {
            return null;
        }

        string styleId = style.StyleId?.Value;
        string defaultId = styles.Elements<Style>().FirstOrDefault(s => s.Type != null && s.Type.HasValue && s.Type.Value == StyleValues.Paragraph
            && s.Default != null && s.Default.HasValue && s.Default.Value)?.StyleId?.Value;
        List<Paragraph> paragraphs = body.Descendants<Paragraph>().Where(p => !p.Ancestors<TextBoxContent>().Any()).ToList();
        int at = paragraphs.IndexOf(own);
        bool InStyle(Paragraph p) => string.Equals(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? defaultId, styleId, StringComparison.Ordinal);
        Paragraph found = paragraphs.Take(at + 1).LastOrDefault(InStyle) ?? paragraphs.Skip(at + 1).FirstOrDefault(InStyle);
        if (found == null)
        {
            ok = false;
            return "Error! No text of specified style in document.";
        }

        bool self = ReferenceEquals(found, own);
        string text;
        if (context.HasValue)
        {
            text = ParagraphNumber(field, found, context.Value, withoutWords);
        }
        else if (self)
        {
            ok = false;
            return SelfReference;
        }
        else
        {
            text = position ? string.Empty : new RangeWrapper(found).Text.TrimEnd('\r', '\a');
        }
        if (text == null || !position || self)
        {
            return text;
        }
        string where = found.IsBefore(own) ? "above" : "below";
        return text.Length == 0 ? where : text + " " + where;
    }

    //The list number of a paragraph, as REF and STYLEREF \n, \r and \w show it (and with \t without its words), measured in Word: "0" for a
    //paragraph without a number. Relative to the list item at or before the field. Null (left for Word) for a bullet.
    private static string ParagraphNumber(FieldWrapper field, Paragraph target, ListNumbering.NumberContext context, bool withoutWords)
    {
        OpenXmlElement story = StoryMap.FindRoot(target);
        List<ListNumbering.Item> items = ListNumbering.GetItems(story);
        ListNumbering.Item item = items.FirstOrDefault(i => ReferenceEquals(i.Paragraph, target));
        if (item == null)
        {
            return "0";
        }
        if (ListNumbering.TypeOf(item.Template) == ListType.Bullet)
        {
            return null;
        }

        ListNumbering.Item from = null;
        Paragraph own = field.Key.Ancestors<Paragraph>().FirstOrDefault();
        if (context == ListNumbering.NumberContext.Relative && own != null && ReferenceEquals(StoryMap.FindRoot(own), story))
        {
            Dictionary<Paragraph, int> order = new Dictionary<Paragraph, int>();
            foreach (Paragraph paragraph in story.Descendants<Paragraph>())
            {
                order[paragraph] = order.Count;
            }
            int at = order.TryGetValue(own, out int index) ? index : -1;
            from = items.LastOrDefault(i => order.TryGetValue(i.Paragraph, out int itemIndex) && itemIndex <= at);
        }
        return ListNumbering.NumberText(item, context, from, withoutWords);
    }

    //IF expression1 operator expression2 true-text false-text
    private static string Condition(Instruction instruction)
    {
        List<string> arguments = instruction.Arguments;
        if (arguments.Count < 3)
        {
            return null;
        }

        string left = arguments[0];
        string op = arguments[1];
        string right = arguments[2];
        string whenTrue = arguments.Count > 3 ? arguments[3] : string.Empty;
        string whenFalse = arguments.Count > 4 ? arguments[4] : string.Empty;

        bool numbers = double.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out double a) & double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out double b);
        int comparison = numbers ? a.CompareTo(b) : string.Compare(left, right, StringComparison.OrdinalIgnoreCase);
        bool result;
        switch (op)
        {
            case "=":
                result = numbers ? comparison == 0 : Matches(left, right);
                break;
            case "<>":
                result = numbers ? comparison != 0 : !Matches(left, right);
                break;
            case "<":
                result = comparison < 0;
                break;
            case ">":
                result = comparison > 0;
                break;
            case "<=":
                result = comparison <= 0;
                break;
            case ">=":
                result = comparison >= 0;
                break;
            default:
                return null;
        }
        return result ? whenTrue : whenFalse;
    }

    //Equality as Word does it for text: ignoring case, with * for any string and ? for any character in the second operand.
    private static bool Matches(string value, string pattern)
    {
        string expression = "^" + System.Text.RegularExpressions.Regex.Escape(pattern).Replace("\\*", ".*", StringComparison.Ordinal).Replace("\\?", ".", StringComparison.Ordinal) + "$";
        return System.Text.RegularExpressions.Regex.IsMatch(value, expression, System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant | System.Text.RegularExpressions.RegexOptions.Singleline);
    }

    #endregion

    #region Formats

    //The \* switches of a field, applied in the order they are written, as in Word: Upper, Lower, FirstCap and Caps change the case of the
    //text so far, and a number format (FieldNumberFormats) writes again the number the result started with, so it replaces what the switches
    //before it did. A result that does not start with a number is left as it is by a number format. Gives null (the field is left as it is)
    //for what is not worked out here: a number format before the \# picture switch, which Word reports as an error, or a number Word reads in
    //its own way.
    private static string ApplyFormat(string text, Instruction instruction)
    {
        string original = text;
        bool? parsed = null;
        decimal number = 0;
        bool picture = false;
        foreach (KeyValuePair<string, string> pair in instruction.Switches)
        {
            if (pair.Key == "#")
            {
                picture = true;
            }
            if (pair.Key != "*" || pair.Value == null)
            {
                continue;
            }

            if (FieldNumberFormats.IsNumberFormat(pair.Value))
            {
                if (instruction.Switches.Any(s => s.Key == "#") && !picture)
                {
                    return null;
                }
                if (parsed == null)
                {
                    parsed = FieldNumberFormats.TryParseLeadingNumber(original, CultureInfo.CurrentCulture, out number);
                    if (parsed == null)
                    {
                        return null;
                    }
                }
                if (parsed == true)
                {
                    text = FieldNumberFormats.Format(number, pair.Value);
                }
                continue;
            }

            //The names are not case sensitive, as in Word.
            switch (pair.Value.ToUpperInvariant())
            {
                case "UPPER":
                    text = text.ToUpperInvariant();
                    break;
                case "LOWER":
                    text = text.ToLowerInvariant();
                    break;
                case "FIRSTCAP":
                    text = FirstCap(text);
                    break;
                case "CAPS":
                    text = Caps(text);
                    break;
            }
        }
        return text;
    }

    //FirstCap, as measured in Word: the first character after any spaces is made a capital if it is a letter; the rest is left as it is.
    private static string FirstCap(string text)
    {
        int i = 0;
        while (i < text.Length && char.IsWhiteSpace(text[i]))
        {
            i++;
        }
        if (i == text.Length || !char.IsLetter(text[i]))
        {
            return text;
        }
        char[] characters = text.ToCharArray();
        characters[i] = char.ToUpperInvariant(characters[i]);
        return new string(characters);
    }

    //Caps, as measured in Word: each word (after a space or a hyphen) starts with a capital and goes on in lowercase, except that when the
    //text has lowercase letters a word all in capitals (such as an abbreviation) is left as it is.
    private static string Caps(string text)
    {
        bool keepCapitals = text.Any(char.IsLower);
        char[] characters = text.ToCharArray();
        int start = 0;
        while (start < characters.Length)
        {
            if (char.IsWhiteSpace(characters[start]) || characters[start] == '-')
            {
                start++;
                continue;
            }
            int end = start;
            while (end < characters.Length && !char.IsWhiteSpace(characters[end]) && characters[end] != '-')
            {
                end++;
            }
            bool allCapitals = true;
            for (int k = start; k < end; k++)
            {
                allCapitals &= !char.IsLower(characters[k]);
            }
            if (!(keepCapitals && allCapitals))
            {
                characters[start] = char.ToUpperInvariant(characters[start]);
                for (int k = start + 1; k < end; k++)
                {
                    characters[k] = char.ToLowerInvariant(characters[k]);
                }
            }
            start = end;
        }
        return new string(characters);
    }

    /// <summary>
    /// Formats a date and time with a Word date-time picture: <c>M</c> months, <c>d</c> days, <c>yy</c> and <c>yyyy</c> years, <c>H</c> hours
    /// of 24, <c>h</c> hours of 12, <c>m</c> minutes, <c>s</c> seconds, <c>AM/PM</c>, <c>am/pm</c>, <c>A/P</c> and <c>a/p</c>, with text in
    /// single quotes.
    /// </summary>
    internal static string FormatDate(DateTime date, string picture, CultureInfo culture)
    {

        StringBuilder sb = new StringBuilder();
        int i = 0;
        while (i < picture.Length)
        {
            char c = picture[i];
            int run = 1;
            while (i + run < picture.Length && picture[i + run] == c)
            {
                run++;
            }

            if (c == '\'')
            {
                int close = picture.IndexOf('\'', i + 1);
                if (close < 0)
                {
                    close = picture.Length;
                }
                sb.Append(picture, i + 1, close - i - 1);
                i = close + 1;
                continue;
            }

            string remaining = picture.Substring(i);
            if (remaining.StartsWith("AM/PM", StringComparison.Ordinal))
            {
                sb.Append(date.Hour < 12 ? "AM" : "PM");
                i += 5;
            }
            else if (remaining.StartsWith("am/pm", StringComparison.Ordinal))
            {
                sb.Append(date.Hour < 12 ? "am" : "pm");
                i += 5;
            }
            else if (remaining.StartsWith("A/P", StringComparison.Ordinal))
            {
                sb.Append(date.Hour < 12 ? "A" : "P");
                i += 3;
            }
            else if (remaining.StartsWith("a/p", StringComparison.Ordinal))
            {
                sb.Append(date.Hour < 12 ? "a" : "p");
                i += 3;
            }
            else
            {
                switch (c)
                {
                    case 'M':
                        sb.Append(run == 1 ? date.Month.ToString(CultureInfo.InvariantCulture)
                            : run == 2 ? date.Month.ToString("00", CultureInfo.InvariantCulture)
                            : run == 3 ? culture.DateTimeFormat.GetAbbreviatedMonthName(date.Month)
                            : culture.DateTimeFormat.GetMonthName(date.Month));
                        break;
                    case 'd':
                        sb.Append(run == 1 ? date.Day.ToString(CultureInfo.InvariantCulture)
                            : run == 2 ? date.Day.ToString("00", CultureInfo.InvariantCulture)
                            : run == 3 ? culture.DateTimeFormat.GetAbbreviatedDayName(date.DayOfWeek)
                            : culture.DateTimeFormat.GetDayName(date.DayOfWeek));
                        break;
                    case 'y':
                        sb.Append(run <= 2 ? (date.Year % 100).ToString("00", CultureInfo.InvariantCulture) : date.Year.ToString("0000", CultureInfo.InvariantCulture));
                        break;
                    case 'H':
                        sb.Append(date.Hour.ToString(run == 1 ? "0" : "00", CultureInfo.InvariantCulture));
                        break;
                    case 'h':
                        int hour = date.Hour % 12 == 0 ? 12 : date.Hour % 12;
                        sb.Append(hour.ToString(run == 1 ? "0" : "00", CultureInfo.InvariantCulture));
                        break;
                    case 'm':
                        sb.Append(date.Minute.ToString(run == 1 ? "0" : "00", CultureInfo.InvariantCulture));
                        break;
                    case 's':
                        sb.Append(date.Second.ToString(run == 1 ? "0" : "00", CultureInfo.InvariantCulture));
                        break;
                    default:
                        sb.Append(c, run);
                        break;
                }
                i += run;
            }
        }
        return sb.ToString();
    }

    #endregion
}
