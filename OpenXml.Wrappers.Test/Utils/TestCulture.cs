using System.Globalization;
using System.Runtime.CompilerServices;

namespace OpenXml.Wrappers.Test.Utils;

/// <summary>
/// Fixes the culture of the tests. A new document takes its page size, header distance and language from the current culture, as Word takes
/// them from Windows, so the tests would otherwise depend on the machine; en-GB is the culture of the Word that wrote the references.
/// </summary>
internal static class TestCulture
{
    public static readonly CultureInfo Reference = CultureInfo.GetCultureInfo("en-GB");

#pragma warning disable CA2255 // This is a test assembly, not a library: nothing else depends on when this runs.
    [ModuleInitializer]
#pragma warning restore CA2255
    internal static void Initialize()
    {
        CultureInfo.DefaultThreadCurrentCulture = Reference;
        CultureInfo.CurrentCulture = Reference;
    }
}
