using System.Runtime.CompilerServices;

namespace OpenXml.Wrappers;

/// <summary>
/// The argument checks of <c>ArgumentNullException.ThrowIfNull</c> and friends, which <c>netstandard2.0</c> does not have, so that the same code
/// builds for every target.
/// </summary>
internal static class Guard
{
    public static void ThrowIfNull(object argument, [CallerArgumentExpression(nameof(argument))] string paramName = null)
    {
        if (argument == null)
        {
            throw new ArgumentNullException(paramName);
        }
    }

    public static void ThrowIfNegative(int value, [CallerArgumentExpression(nameof(value))] string paramName = null)
    {
        if (value < 0)
        {
            throw new ArgumentOutOfRangeException(paramName, value, "The value must not be negative.");
        }
    }

    public static void ThrowIfLessThan(int value, int other, [CallerArgumentExpression(nameof(value))] string paramName = null)
    {
        if (value < other)
        {
            throw new ArgumentOutOfRangeException(paramName, value, "The value must be at least " + other.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
        }
    }

    public static void ThrowIfLessThanOrEqual(long value, long other, [CallerArgumentExpression(nameof(value))] string paramName = null)
    {
        if (value <= other)
        {
            throw new ArgumentOutOfRangeException(paramName, value, "The value must be greater than " + other.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".");
        }
    }
}

/// <summary>
/// The generic <c>Enum.IsDefined</c> and <c>Enum.GetValues</c>, which <c>netstandard2.0</c> does not have.
/// </summary>
internal static class EnumCompat
{
    public static bool IsDefined<T>(T value)
        where T : struct, Enum
    {
#if NETSTANDARD2_0
        return Enum.IsDefined(typeof(T), value);
#else
        return Enum.IsDefined(value);
#endif
    }

    public static T[] GetValues<T>()
        where T : struct, Enum
    {
#if NETSTANDARD2_0
        return (T[])Enum.GetValues(typeof(T));
#else
        return Enum.GetValues<T>();
#endif
    }
}

#if NETSTANDARD2_0

/// <summary>
/// The members of <c>string</c> that <c>netstandard2.0</c> does not have. On the other targets the instance members are used.
/// </summary>
internal static class StringCompat
{
    public static bool Contains(this string text, char value)
    {
        return text.IndexOf(value) >= 0;
    }

    public static bool Contains(this string text, char value, StringComparison comparison)
    {
        return text.IndexOf(value.ToString(), comparison) >= 0;
    }

    public static bool Contains(this string text, string value, StringComparison comparison)
    {
        return text.IndexOf(value, comparison) >= 0;
    }

    public static int IndexOf(this string text, char value, StringComparison comparison)
    {
        return text.IndexOf(value.ToString(), comparison);
    }

    public static bool EndsWith(this string text, char value)
    {
        return text.Length > 0 && text[text.Length - 1] == value;
    }

    public static bool StartsWith(this string text, char value)
    {
        return text.Length > 0 && text[0] == value;
    }

    public static string Replace(this string text, string oldValue, string newValue, StringComparison comparison)
    {
        Guard.ThrowIfNull(oldValue);
        if (oldValue.Length == 0)
        {
            throw new ArgumentException("The string to replace must not be empty.", nameof(oldValue));
        }
        newValue = newValue ?? String.Empty;
        System.Text.StringBuilder result = null;
        int position = 0;
        int found;
        while ((found = text.IndexOf(oldValue, position, comparison)) >= 0)
        {
            result = result ?? new System.Text.StringBuilder(text.Length);
            result.Append(text, position, found - position).Append(newValue);
            position = found + oldValue.Length;
        }
        return result == null ? text : result.Append(text, position, text.Length - position).ToString();
    }
}

#endif
