#if NETFRAMEWORK
namespace System.Runtime.CompilerServices;

//net48 does not have the attribute; the compiler only looks for its name.
[AttributeUsage(AttributeTargets.Method, Inherited = false)]
internal sealed class ModuleInitializerAttribute : Attribute
{
}
#endif
