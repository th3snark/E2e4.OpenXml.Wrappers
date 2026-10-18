#if NETFRAMEWORK

namespace OpenXml.Wrappers.Samples;

/// <summary>
/// The members of <c>string</c> that the examples use and that .NET Framework does not have.
/// </summary>
internal static class NetFrameworkCompat
{
    public static bool Contains(this string text, string value, StringComparison comparison)
    {
        return text.IndexOf(value, comparison) >= 0;
    }
}

#endif
