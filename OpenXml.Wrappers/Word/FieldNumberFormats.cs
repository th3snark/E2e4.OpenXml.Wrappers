using System;
using System.Globalization;
using System.Text;

namespace OpenXml.Wrappers.Word;

/// <summary>
/// The number formats of a field's <c>\*</c> switch (ARABIC, ROMAN and roman, ALPHABETIC and alphabetic, Ordinal, CardText, OrdText, Hex,
/// ArabicDash and DollarText), as measured in Word: <c>tools\WordReference\Measure-NumberFormats.ps1</c> writes what Word shows, and
/// <c>FieldNumberFormatTest</c> compares with it.
/// </summary>
internal static class FieldNumberFormats
{
    internal const string CannotBeRepresented = "Error! Number cannot be represented in specified format.";

    private static readonly string[] Names = { "ARABIC", "ROMAN", "ALPHABETIC", "ORDINAL", "CARDTEXT", "ORDTEXT", "HEX", "ARABICDASH", "DOLLARTEXT" };

    private static readonly string[] Units =
    {
        "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen", "fourteen",
        "fifteen", "sixteen", "seventeen", "eighteen", "nineteen",
    };

    private static readonly string[] Tens = { string.Empty, string.Empty, "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

    /// <summary>
    /// Whether the value of a <c>\*</c> switch is a number format rather than a case format (Upper, Lower, FirstCap, Caps) or MERGEFORMAT.
    /// The names are not case sensitive.
    /// </summary>
    internal static bool IsNumberFormat(string format)
    {
        return format != null && Array.IndexOf(Names, format.ToUpperInvariant()) >= 0;
    }

    /// <summary>
    /// The number a field's result starts with, as Word reads it for a number format: spaces before it are skipped, digit grouping is
    /// allowed and anything after it is ignored. Gives false when the text does not start with a number (the format then leaves it as it
    /// is), and null when it is one Word reads in a way not worked out here (a number with an exponent).
    /// </summary>
    internal static bool? TryParseLeadingNumber(string text, CultureInfo culture, out decimal value)
    {
        value = 0;
        NumberFormatInfo numbers = culture.NumberFormat;
        int i = 0;
        while (i < text.Length && char.IsWhiteSpace(text[i]))
        {
            i++;
        }
        int start = i;
        if (i < text.Length && text[i] == '-')
        {
            i++;
        }

        StringBuilder digits = new StringBuilder();
        bool any = false;
        while (i < text.Length)
        {
            if (char.IsDigit(text[i]))
            {
                digits.Append(text[i]);
                any = true;
                i++;
            }
            else if (any && string.CompareOrdinal(text, i, numbers.NumberGroupSeparator, 0, numbers.NumberGroupSeparator.Length) == 0
                && i + numbers.NumberGroupSeparator.Length < text.Length && char.IsDigit(text[i + numbers.NumberGroupSeparator.Length]))
            {
                i += numbers.NumberGroupSeparator.Length;
            }
            else
            {
                break;
            }
        }
        if (string.CompareOrdinal(text, i, numbers.NumberDecimalSeparator, 0, numbers.NumberDecimalSeparator.Length) == 0)
        {
            int after = i + numbers.NumberDecimalSeparator.Length;
            if (after < text.Length && char.IsDigit(text[after]))
            {
                digits.Append('.');
                i = after;
                while (i < text.Length && char.IsDigit(text[i]))
                {
                    digits.Append(text[i]);
                    any = true;
                    i++;
                }
            }
        }
        if (!any)
        {
            return false;
        }
        if (i < text.Length && (text[i] == 'e' || text[i] == 'E'))
        {
            return null;
        }
        if (!decimal.TryParse(digits.ToString(), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value))
        {
            return null;
        }
        if (text[start] == '-')
        {
            value = -value;
        }
        return true;
    }

    /// <summary>
    /// Writes a number in a format of the <c>\*</c> switch, or gives Word's error when the format cannot show it. A value that is not a whole
    /// number is rounded half away from zero (DollarText keeps its cents instead).
    /// </summary>
    internal static string Format(decimal value, string format)
    {
        string name = format.ToUpperInvariant();
        if (name == "DOLLARTEXT")
        {
            return DollarText(value);
        }

        //Word rounds half away from zero, except that a fraction from a half up to (but not) one cannot be shown at all.
        if (value >= 0.5m && value < 1m)
        {
            return CannotBeRepresented;
        }
        decimal rounded = Math.Round(value, MidpointRounding.AwayFromZero);
        if (Math.Abs(rounded) > long.MaxValue / 10)
        {
            return CannotBeRepresented;
        }
        long number = (long)rounded;

        switch (name)
        {
            case "ARABIC":
                return number.ToString(CultureInfo.InvariantCulture);
            case "ARABICDASH":
                return "- " + number.ToString(CultureInfo.InvariantCulture) + " -";
            case "ORDINAL":
                return number.ToString(CultureInfo.InvariantCulture) + OrdinalSuffix(Math.Abs(number));
            case "ROMAN":
                return Roman(number, char.IsLower(format[0]));
            case "ALPHABETIC":
                return Letters(number, char.IsLower(format[0]));
            case "HEX":
                return number < 0 || number > 0xFFFF ? CannotBeRepresented : number.ToString("X", CultureInfo.InvariantCulture);
            case "CARDTEXT":
                return number < 0 || number > 999999 ? CannotBeRepresented : Cardinal(number);
            case "ORDTEXT":
                return number < 0 || number > 999999 ? CannotBeRepresented : OrdinalWords(Cardinal(number));
            default:
                return number.ToString(CultureInfo.InvariantCulture);
        }
    }

    private static string OrdinalSuffix(long number)
    {
        long tens = number % 100;
        if (tens >= 11 && tens <= 13)
        {
            return "th";
        }
        switch (number % 10)
        {
            case 1:
                return "st";
            case 2:
                return "nd";
            case 3:
                return "rd";
            default:
                return "th";
        }
    }

    //Up to 32767, with as many Ms as there are thousands; nought shows as a space, as in Word.
    private static string Roman(long number, bool lower)
    {
        if (number == 0)
        {
            return " ";
        }
        if (number < 0 || number > 32767)
        {
            return CannotBeRepresented;
        }
        (int Value, string Symbol)[] table = { (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
        StringBuilder text = new StringBuilder(new string('M', (int)(number / 1000)));
        long rest = number % 1000;
        foreach ((int part, string symbol) in table)
        {
            while (rest >= part)
            {
                text.Append(symbol);
                rest -= part;
            }
        }
        return lower ? text.ToString().ToLowerInvariant() : text.ToString();
    }

    //A to Z, then AA to ZZ and so on, up to 30 letters (780); nought shows as a space, as in Word.
    private static string Letters(long number, bool lower)
    {
        if (number == 0)
        {
            return " ";
        }
        if (number < 0 || number > 780)
        {
            return CannotBeRepresented;
        }
        char letter = (char)((lower ? 'a' : 'A') + (int)((number - 1) % 26));
        return new string(letter, (int)((number - 1) / 26) + 1);
    }

    private static string DollarText(decimal value)
    {
        if (value < 0)
        {
            return CannotBeRepresented;
        }
        decimal whole = decimal.Truncate(value);
        if (whole > 999999)
        {
            return CannotBeRepresented;
        }
        //The cents are rounded on their own: what rounds up to a whole dollar shows as 00, as in Word.
        int cents = (int)(Math.Round((value - whole) * 100, MidpointRounding.AwayFromZero) % 100);
        return Cardinal((long)whole) + " and " + cents.ToString("00", CultureInfo.InvariantCulture) + "/100";
    }

    //Below a million, in Word's words: "one hundred five", "twenty-one", "one thousand two hundred thirty-four".
    private static string Cardinal(long number)
    {
        if (number >= 1000)
        {
            long rest = number % 1000;
            return Cardinal(number / 1000) + " thousand" + (rest == 0 ? string.Empty : " " + Cardinal(rest));
        }
        if (number >= 100)
        {
            long rest = number % 100;
            return Units[number / 100] + " hundred" + (rest == 0 ? string.Empty : " " + Cardinal(rest));
        }
        if (number >= 20)
        {
            return Tens[number / 10] + (number % 10 == 0 ? string.Empty : "-" + Units[number % 10]);
        }
        return Units[number];
    }

    //The ordinal of the words of a number: the last word becomes its ordinal ("twenty-one" gives "twenty-first").
    private static string OrdinalWords(string cardinal)
    {
        int at = Math.Max(cardinal.LastIndexOf(' '), cardinal.LastIndexOf('-')) + 1;
        string last = cardinal.Substring(at);
        string ordinal;
        switch (last)
        {
            case "one":
                ordinal = "first";
                break;
            case "two":
                ordinal = "second";
                break;
            case "three":
                ordinal = "third";
                break;
            case "five":
                ordinal = "fifth";
                break;
            case "eight":
                ordinal = "eighth";
                break;
            case "nine":
                ordinal = "ninth";
                break;
            case "twelve":
                ordinal = "twelfth";
                break;
            default:
                ordinal = last[last.Length - 1] == 'y' ? last.Remove(last.Length - 1) + "ieth" : last + "th";
                break;
        }
        return cardinal.Remove(at) + ordinal;
    }
}
