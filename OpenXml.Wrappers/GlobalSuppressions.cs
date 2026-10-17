using System.Diagnostics.CodeAnalysis;

// The wrapper collections deliberately implement only the non-generic IEnumerable with a hand-written enumerator (see CLAUDE.md).
[assembly: SuppressMessage("Design", "CA1010:Collections should implement generic interface", Justification = "Deliberate collection pattern; changing it would alter the public API.", Scope = "namespaceanddescendants", Target = "~N:OpenXml.Wrappers.Word")]
