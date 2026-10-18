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
    /// name, number of words and characters, SEQ, REF, IF and QUOTE. The switches for the format of the result (<c>\@</c> for dates and
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
                    return pair.Value ?? String.Empty;
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
        foreach (char c in code ?? String.Empty)
        {
            if (c == '"' || c == '“' || c == '”')
            {
                quoted = !quoted;
                hasToken = true;
            }
            else if (Char.IsWhiteSpace(c) && !quoted)
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

        Instruction instruction = new Instruction() { Name = tokens.Count > 0 ? tokens[0].ToUpperInvariant() : String.Empty };
        for (int i = 1; i < tokens.Count; i++)
        {
            string token = tokens[i];
            if (token.Length >= 2 && token[0] == '\\')
            {
                string key = token.Substring(1);
                string value = null;
                bool takesArgument = key == "@" || key == "*" || key == "#" || ((key == "r" || key == "s") && i + 1 < tokens.Count && Int32.TryParse(tokens[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out _));
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
        return ApplyFormat(text, instruction);
    }

    private static string ComputeText(FieldWrapper field, Instruction instruction, MainDocumentPart main, WordprocessingDocument document, ref bool ok)
    {
        CultureInfo culture = CultureInfo.CurrentCulture;
        DocumentPropertiesWrapper properties = document == null ? null : new DocumentPropertiesWrapper(document);
        string picture = instruction.GetSwitch("@");

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
                DateTime? printed = document?.PackageProperties.LastPrinted;
                return printed == null ? "0/0/0000 0:00" : FormatDate(printed.Value.ToLocalTime(), picture ?? DefaultDateTime(culture), culture);
            case "AUTHOR":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Author ?? String.Empty;
            case "TITLE":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Title ?? String.Empty;
            case "SUBJECT":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Subject ?? String.Empty;
            case "KEYWORDS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Keywords ?? String.Empty;
            case "COMMENTS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : properties?.Comments ?? String.Empty;
            case "LASTSAVEDBY":
                return properties?.LastAuthor ?? String.Empty;
            case "REVNUM":
                return document?.PackageProperties.Revision ?? "1";
            case "USERNAME":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : Environment.UserName;
            case "USERINITIALS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : Environment.UserName.Length > 0 ? Environment.UserName.Substring(0, 1).ToUpperInvariant() : String.Empty;
            case "USERADDRESS":
                return instruction.Arguments.Count > 0 ? instruction.Arguments[0] : String.Empty;
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
                    .FirstOrDefault(v => String.Equals(v.Name?.Value, instruction.Arguments[0], StringComparison.OrdinalIgnoreCase));
                if (variable == null)
                {
                    ok = false;
                    return "Error! No document variable supplied.";
                }
                return variable.Val?.Value ?? String.Empty;
            case "QUOTE":
                return String.Join(" ", instruction.Arguments);
            case "SEQ":
                return Sequence(field, instruction);
            case "REF":
                return Reference(field, instruction, ref ok);
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
        switch (name.Replace(" ", String.Empty, StringComparison.Ordinal).ToUpperInvariant())
        {
            case "TITLE":
                return properties.Title ?? String.Empty;
            case "SUBJECT":
                return properties.Subject ?? String.Empty;
            case "AUTHOR":
                return properties.Author ?? String.Empty;
            case "KEYWORDS":
                return properties.Keywords ?? String.Empty;
            case "COMMENTS":
                return properties.Comments ?? String.Empty;
            case "CATEGORY":
                return properties.Category ?? String.Empty;
            case "LASTAUTHOR":
            case "LASTSAVEDBY":
                return properties.LastAuthor ?? String.Empty;
            case "COMPANY":
                return properties.Company ?? String.Empty;
            case "MANAGER":
                return properties.Manager ?? String.Empty;
            case "CREATIONDATE":
            case "CREATEDATE":
                return properties.CreationDate?.ToLocalTime().ToString(CultureInfo.CurrentCulture) ?? String.Empty;
            case "LASTSAVETIME":
            case "SAVEDATE":
                return properties.LastSaveTime?.ToLocalTime().ToString(CultureInfo.CurrentCulture) ?? String.Empty;
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
        return Convert.ToString(value, CultureInfo.CurrentCulture) ?? String.Empty;
    }

    private static string GetStoryText(MainDocumentPart main)
    {
        return new RangeWrapper(main.Document).Text ?? String.Empty;
    }

    private static int CountWords(MainDocumentPart main)
    {
        return GetStoryText(main).Split(new[] { ' ', '\t', '\r', '\n', '\a', '\v', '\f', ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
    }

    private static int CountCharacters(MainDocumentPart main)
    {
        return GetStoryText(main).Count(c => !Char.IsWhiteSpace(c) && !Char.IsControl(c));
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
        long value = 0;
        foreach (FieldWrapper other in new FieldWrapperCollection(new RangeWrapper(field.Key.GetRoot())))
        {
            Instruction otherInstruction = ReferenceEquals(other, field) || other.Equals(field) ? instruction : Parse(other.CodeText);
            if (otherInstruction.Name == "SEQ" && otherInstruction.Arguments.Count > 0
                && String.Equals(otherInstruction.Arguments[0], identifier, StringComparison.OrdinalIgnoreCase))
            {
                string reset = otherInstruction.GetSwitch("r");
                if (reset != null && Int64.TryParse(reset, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number))
                {
                    value = number;
                }
                else if (otherInstruction.GetSwitch("c") == null)
                {
                    value++;
                }
            }

            if (other.Equals(field))
            {
                break;
            }
        }
        //The switch that hides the number leaves the result empty.
        return instruction.GetSwitch("h") != null ? String.Empty : FormatNumber(value, instruction.GetSwitch("*"));
    }

    private static string Reference(FieldWrapper field, Instruction instruction, ref bool ok)
    {
        if (instruction.Arguments.Count == 0)
        {
            return null;
        }
        foreach (string unsupported in new[] { "n", "r", "w", "p", "f", "t" })
        {
            if (instruction.GetSwitch(unsupported) != null)
            {
                return null;
            }
        }

        OpenXmlElement root = field.Key.GetRoot();
        BookmarkStart bookmark = root.Descendants<BookmarkStart>().FirstOrDefault(b => String.Equals(b.Name?.Value, instruction.Arguments[0], StringComparison.OrdinalIgnoreCase));
        if (bookmark == null)
        {
            ok = false;
            return "Error! Reference source not found.";
        }
        return new BookmarkWrapper(bookmark).Range.Text.TrimEnd('\r', '\a');
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
        string whenTrue = arguments.Count > 3 ? arguments[3] : String.Empty;
        string whenFalse = arguments.Count > 4 ? arguments[4] : String.Empty;

        bool numbers = Double.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out double a) & Double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out double b);
        int comparison = numbers ? a.CompareTo(b) : String.Compare(left, right, StringComparison.OrdinalIgnoreCase);
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

    //The format switches of a field: \* Upper, Lower, FirstCap and Caps for text, and ARABIC, roman, ROMAN, alphabetic and ALPHABETIC for numbers.
    private static string ApplyFormat(string text, Instruction instruction)
    {
        foreach (KeyValuePair<string, string> pair in instruction.Switches)
        {
            if (pair.Key != "*" || pair.Value == null)
            {
                continue;
            }

            switch (pair.Value)
            {
                case "Upper":
                    text = text.ToUpperInvariant();
                    break;
                case "Lower":
                    text = text.ToLowerInvariant();
                    break;
                case "FirstCap":
                    text = text.Length == 0 ? text : Char.ToUpperInvariant(text[0]) + text.Substring(1).ToLowerInvariant();
                    break;
                case "Caps":
                    text = CultureInfo.InvariantCulture.TextInfo.ToTitleCase(text.ToLowerInvariant());
                    break;
            }
        }
        return text;
    }

    private static string FormatNumber(long value, string format)
    {
        if (format == null || String.Equals(format, "ARABIC", StringComparison.OrdinalIgnoreCase))
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }
        if (value <= 0)
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }
        if (String.Equals(format, "ROMAN", StringComparison.OrdinalIgnoreCase))
        {
            string roman = ToRoman(value);
            return format == "ROMAN" ? roman : roman.ToLowerInvariant();
        }
        if (String.Equals(format, "ALPHABETIC", StringComparison.OrdinalIgnoreCase))
        {
            string letters = ToLetters(value);
            return format == "ALPHABETIC" ? letters : letters.ToLowerInvariant();
        }
        return value.ToString(CultureInfo.InvariantCulture);
    }

    private static string ToRoman(long value)
    {
        (int Value, string Symbol)[] table = { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
        StringBuilder sb = new StringBuilder();
        foreach ((int number, string symbol) in table)
        {
            while (value >= number)
            {
                sb.Append(symbol);
                value -= number;
            }
        }
        return sb.ToString();
    }

    //A, B, ... Z, AA, BB, ... as Word numbers with letters.
    private static string ToLetters(long value)
    {
        char letter = (char)('A' + (int)((value - 1) % 26));
        return new string(letter, (int)((value - 1) / 26) + 1);
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
