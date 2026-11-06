using System;
using System.Globalization;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Office2010 = DocumentFormat.OpenXml.Office2010.Word;
using Office2013 = DocumentFormat.OpenXml.Office2013.Word;
using WCalendar = DocumentFormat.OpenXml.Wordprocessing.Calendar;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// How a date content control stores its value, as <c>WdContentControlDateStorageFormat</c> in Word (with the same values).
/// </summary>
public enum ContentControlDateStorageFormat
{
    /// <summary>As text, as it is displayed.</summary>
    Text = 0,

    /// <summary>As a date.</summary>
    Date = 1,

    /// <summary>As a date and time.</summary>
    DateTime = 2
}

/// <summary>
/// The calendar a date is shown in, as <c>WdCalendarType</c> in Word (with the same values).
/// </summary>
public enum CalendarType
{
    /// <summary>The Gregorian calendar.</summary>
    Western = 0,

    /// <summary>The Arabic (Hijri) calendar.</summary>
    Arabic = 1,

    /// <summary>The Hebrew calendar.</summary>
    Hebrew = 2,

    /// <summary>The Taiwanese calendar.</summary>
    Taiwan = 3,

    /// <summary>The Japanese calendar.</summary>
    Japan = 4,

    /// <summary>The Thai calendar.</summary>
    Thai = 5,

    /// <summary>The Korean calendar.</summary>
    Korean = 6,

    /// <summary>The Saka era calendar.</summary>
    SakaEra = 7,

    /// <summary>The Gregorian calendar, with English month names.</summary>
    TranslitEnglish = 8,

    /// <summary>The Gregorian calendar, with French month names.</summary>
    TranslitFrench = 9
}

/// <summary>
/// Where a content control sits, as <c>WdContentControlLevel</c> in Word (with the same values).
/// </summary>
public enum ContentControlLevel
{
    /// <summary>Inside a paragraph.</summary>
    Inline = 0,

    /// <summary>Around paragraphs.</summary>
    Paragraph = 1,

    /// <summary>Around a table row.</summary>
    Row = 2,

    /// <summary>Around a table cell.</summary>
    Cell = 3
}

public partial class ContentControlWrapper
{
    private const string W14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml";

    private T Require<T>(string what) where T : OpenXmlElement
    {
        T found = Find<T>();
        if (found == null)
        {
            throw new InvalidOperationException("The content control is not " + what + ".");
        }
        return found;
    }

    /// <summary>Where the control is, as <c>ContentControl.Level</c> in Word.</summary>
    public ContentControlLevel Level
    {
        get
        {
            return _sdt is SdtBlock ? ContentControlLevel.Paragraph
                : _sdt is SdtRow ? ContentControlLevel.Row
                : _sdt is SdtCell ? ContentControlLevel.Cell : ContentControlLevel.Inline;
        }
    }

    /// <summary>True if the control is taken away once its contents are edited, as <c>ContentControl.Temporary</c> in Word.</summary>
    public bool Temporary
    {
        get { return GetProperties(false)?.ChildElements.Any(e => e.LocalName == "temporary") == true; }
        set { Set("temporary", value ? new TemporarySdt() : null); }
    }

    /// <summary>
    /// The colour of the outline of the control, as <c>ContentControl.Color</c> in Word. Automatic when none is set.
    /// </summary>
    public WordColor Color
    {
        get
        {
            string stored = Find<Office2013.Color>()?.Val?.Value;
            return stored == null ? WordColor.Automatic : WordColor.FromStored(stored);
        }
        set { Set("color", value.IsAutomatic ? null : new Office2013.Color() { Val = value.Hex }); }
    }

    /// <summary>
    /// The character style the contents of the control are given, as <c>ContentControl.DefaultTextStyle</c> in Word; null when there is none.
    /// </summary>
    /// <exception cref="ArgumentException">There is no style with the name, or it is not a character style.</exception>
    public string DefaultTextStyle
    {
        get
        {
            string id = GetProperties(false)?.GetFirstChild<RunProperties>()?.RunStyle?.Val?.Value;
            return id == null ? null : StyleHelper.GetStyleName(_sdt, id);
        }
        set
        {
            SdtProperties properties = GetProperties(true);
            foreach (RunProperties existing in properties.Elements<RunProperties>().ToList())
            {
                existing.Remove();
            }
            if (value == null)
            {
                return;
            }
            string id = StyleHelper.FindStyleId(_sdt, value, out StyleValues type);
            if (type != StyleValues.Character)
            {
                throw new ArgumentException("The style is not a character style.", nameof(value));
            }
            Place(properties, new RunProperties(new RunStyle() { Val = id }));
            if (!ShowingPlaceholderText)
            {
                Range.Style = value;
            }
        }
    }

    //Makes the control hold the text, as when an entry is chosen or a check box is ticked.
    internal void ReplaceContents(string text, Run template = null)
    {
        SdtProperties properties = GetProperties(false);
        properties?.RemoveAllChildren<ShowingPlaceholder>();
        SdtContentRun content = (_sdt as SdtRun)?.SdtContentRun;
        if (content == null)
        {
            throw new NotSupportedException("Only a control inside a paragraph can hold other text.");
        }
        content.RemoveAllChildren();
        Run run = template ?? new Run();
        run.RemoveAllChildren<Text>();
        run.AppendChild(OpenXmlExtensions.NewText(text));
        content.AppendChild(run);
    }

    #region List entries

    /// <summary>The entries of a drop-down list or combo box, as <c>ContentControl.DropdownListEntries</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a drop-down list or combo box.</exception>
    public ContentControlListEntries DropdownListEntries
    {
        get
        {
            OpenXmlCompositeElement list = (OpenXmlCompositeElement)Find<SdtContentDropDownList>() ?? Find<SdtContentComboBox>();
            if (list == null)
            {
                throw new InvalidOperationException("The content control is not a drop-down list or combo box.");
            }
            return new ContentControlListEntries(list);
        }
    }

    #endregion

    #region Check box

    private OpenXmlElement CheckBoxPart(string name)
    {
        Office2010.SdtContentCheckBox box = Require<Office2010.SdtContentCheckBox>("a check box");
        return box.ChildElements.First(e => e.LocalName == name);
    }

    private static string Attribute(OpenXmlElement element, string name)
    {
        return element.GetAttributes().FirstOrDefault(a => a.LocalName == name).Value;
    }

    private static void SetAttribute(OpenXmlElement element, string name, string value)
    {
        element.SetAttribute(new OpenXmlAttribute("w14", name, W14Namespace, value));
    }

    /// <summary>True if the check box is ticked, as <c>ContentControl.Checked</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a check box.</exception>
    public bool Checked
    {
        get
        {
            string value = Attribute(CheckBoxPart("checked"), "val");
            return value == "1" || value == "true";
        }
        set
        {
            SetAttribute(CheckBoxPart("checked"), "val", value ? "1" : "0");
            ShowState();
        }
    }

    private void ShowState()
    {
        OpenXmlElement state = CheckBoxPart(Checked ? "checkedState" : "uncheckedState");
        string font = Attribute(state, "font");
        string character = char.ConvertFromUtf32(int.Parse(Attribute(state, "val"), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
        RunFonts fonts = font == "MS Gothic"
            ? new RunFonts() { Ascii = font, EastAsia = font, HighAnsi = font, Hint = FontTypeHintValues.EastAsia }
            : new RunFonts() { Ascii = font, HighAnsi = font };
        ReplaceContents(character, new Run(new RunProperties(fonts)));
    }

    private void SetSymbol(string name, int character, string font)
    {
        OpenXmlElement state = CheckBoxPart(name);
        SetAttribute(state, "font", font ?? Attribute(state, "font"));
        SetAttribute(state, "val", character.ToString("X4", CultureInfo.InvariantCulture));
        ShowState();
    }

    /// <summary>Sets the symbol a ticked check box shows, as <c>ContentControl.SetCheckedSymbol</c> in Word.</summary>
    /// <param name="characterNumber">The character code of the symbol.</param>
    /// <param name="font">The font of the symbol; the font it has now when null.</param>
    public void SetCheckedSymbol(int characterNumber, string font = null)
    {
        SetSymbol("checkedState", characterNumber, font);
    }

    /// <summary>Sets the symbol an unticked check box shows, as <c>ContentControl.SetUncheckedSymbol</c> in Word.</summary>
    /// <param name="characterNumber">The character code of the symbol.</param>
    /// <param name="font">The font of the symbol; the font it has now when null.</param>
    public void SetUncheckedSymbol(int characterNumber, string font = null)
    {
        SetSymbol("uncheckedState", characterNumber, font);
    }

    #endregion

    #region Text and date

    /// <summary>True if the text of a plain text control can run over several lines, as <c>ContentControl.MultiLine</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a plain text control.</exception>
    public bool MultiLine
    {
        get { return Require<SdtContentText>("a plain text control").MultiLine?.Value == true; }
        set
        {
            SdtContentText text = Require<SdtContentText>("a plain text control");
            //Word writes "1", where the SDK would write "true".
            const string main = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            if (value)
            {
                text.SetAttribute(new OpenXmlAttribute("w", "multiLine", main, "1"));
            }
            else
            {
                text.RemoveAttribute("multiLine", main);
            }
        }
    }

    /// <summary>The format a date is shown in, as <c>ContentControl.DateDisplayFormat</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a date picker.</exception>
    public string DateDisplayFormat
    {
        get { return Require<SdtContentDate>("a date picker").GetFirstChild<DateFormat>()?.Val?.Value ?? string.Empty; }
        set
        {
            Guard.ThrowIfNull(value);
            SdtContentDate date = Require<SdtContentDate>("a date picker");
            DateFormat format = date.GetFirstChild<DateFormat>();
            if (format == null)
            {
                format = new DateFormat();
                date.PrependChild(format);
            }
            format.Val = value;
        }
    }

    /// <summary>The language a date is shown in, as the locale id <c>ContentControl.DateDisplayLocale</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a date picker.</exception>
    public int DateDisplayLocale
    {
        get
        {
            string name = Require<SdtContentDate>("a date picker").GetFirstChild<LanguageId>()?.Val?.Value;
            return name == null ? CultureInfo.CurrentCulture.LCID : new CultureInfo(name).LCID;
        }
        set
        {
            SdtContentDate date = Require<SdtContentDate>("a date picker");
            LanguageId language = date.GetFirstChild<LanguageId>();
            if (language == null)
            {
                language = new LanguageId();
                DateFormat format = date.GetFirstChild<DateFormat>();
                if (format != null)
                {
                    format.InsertAfterSelf(language);
                }
                else
                {
                    date.PrependChild(language);
                }
            }
            language.Val = new CultureInfo(value).Name;
        }
    }

    private static readonly CalendarType[] CalendarOrder =
    {
        CalendarType.Western, CalendarType.Arabic, CalendarType.Hebrew, CalendarType.Taiwan, CalendarType.Japan,
        CalendarType.Thai, CalendarType.Korean, CalendarType.SakaEra, CalendarType.TranslitEnglish, CalendarType.TranslitFrench
    };

    private static CalendarValues ToCalendarValue(CalendarType type)
    {
        switch (type)
        {
            case CalendarType.Arabic:
                return CalendarValues.Hijri;
            case CalendarType.Hebrew:
                return CalendarValues.Hebrew;
            case CalendarType.Taiwan:
                return CalendarValues.Taiwan;
            case CalendarType.Japan:
                return CalendarValues.Japan;
            case CalendarType.Thai:
                return CalendarValues.Thai;
            case CalendarType.Korean:
                return CalendarValues.Korea;
            case CalendarType.SakaEra:
                return CalendarValues.Saka;
            case CalendarType.TranslitEnglish:
                return CalendarValues.GregorianTransliteratedEnglish;
            case CalendarType.TranslitFrench:
                return CalendarValues.GregorianTransliteratedFrench;
            default:
                return CalendarValues.Gregorian;
        }
    }

    /// <summary>The calendar a date is shown in, as <c>ContentControl.DateCalendarType</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a date picker.</exception>
    public CalendarType DateCalendarType
    {
        get
        {
            CalendarValues? stored = Require<SdtContentDate>("a date picker").GetFirstChild<WCalendar>()?.Val?.Value;
            return stored == null ? CalendarType.Western : CalendarOrder.FirstOrDefault(t => ToCalendarValue(t) == stored.Value);
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            SdtContentDate date = Require<SdtContentDate>("a date picker");
            WCalendar calendar = date.GetFirstChild<WCalendar>();
            if (calendar == null)
            {
                calendar = new WCalendar();
                date.AppendChild(calendar);
            }
            calendar.Val = ToCalendarValue(value);
        }
    }

    /// <summary>How a date is stored, as <c>ContentControl.DateStorageFormat</c> in Word.</summary>
    /// <exception cref="InvalidOperationException">The control is not a date picker.</exception>
    public ContentControlDateStorageFormat DateStorageFormat
    {
        get
        {
            DateFormatValues? stored = Require<SdtContentDate>("a date picker").GetFirstChild<SdtDateMappingType>()?.Val?.Value;
            return stored == DateFormatValues.Text ? ContentControlDateStorageFormat.Text
                : stored == DateFormatValues.Date ? ContentControlDateStorageFormat.Date : ContentControlDateStorageFormat.DateTime;
        }
        set
        {
            if (!EnumCompat.IsDefined(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            SdtContentDate date = Require<SdtContentDate>("a date picker");
            SdtDateMappingType mapping = date.GetFirstChild<SdtDateMappingType>();
            if (mapping == null)
            {
                mapping = new SdtDateMappingType();
                OpenXmlElement before = date.GetFirstChild<WCalendar>();
                if (before != null)
                {
                    before.InsertBeforeSelf(mapping);
                }
                else
                {
                    date.AppendChild(mapping);
                }
            }
            mapping.Val = value == ContentControlDateStorageFormat.Text ? DateFormatValues.Text
                : value == ContentControlDateStorageFormat.Date ? DateFormatValues.Date : DateFormatValues.DateTime;
        }
    }

    #endregion
}
