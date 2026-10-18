#if !NET5_0_OR_GREATER

namespace System.Runtime.CompilerServices;

//Lets records and init-only properties compile for .NET Framework, which does not have this type.
internal static class IsExternalInit
{
}

#endif
