# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A C# library (`OpenXml.Wrappers`, namespace `OpenXml.Wrappers.Word`) that wraps the Open XML SDK's WordprocessingML types to give a Word-object-model-like API (`Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments`, etc.). Solution file: `E2e4.OpenXml.Wrappers.slnx`. Hosted in the private GitHub repo `th3snark/E2e4.OpenXml.Wrappers` (default branch `main`). See `README.md` for usage. There is no CI workflow.

## Build & test

Targets **.NET 10** with `DocumentFormat.OpenXml` 3.x; the `dotnet` CLI is all that is needed:

```bash
dotnet build E2e4.OpenXml.Wrappers.slnx
dotnet test E2e4.OpenXml.Wrappers.slnx
dotnet test E2e4.OpenXml.Wrappers.slnx --filter "FullyQualifiedName~ParagraphWrapperTest"
dotnet format E2e4.OpenXml.Wrappers.slnx --verify-no-changes
```

- Tests are **xUnit**. Per-test setup/teardown is the constructor / `IDisposable.Dispose`. Parallelization is disabled in `Properties\xunit.assemblyinfo.cs` because test classes share fixture files on disk (some opened editable).
- Tests open fixture documents in `OpenXml.Wrappers.Test\TestFiles\` (`Document1.docm`, `Document2.docx`, `Document3.docx`), referenced through constants in `Utils\TestFile.cs` and copied to the output directory by a `CopyToOutputDirectory` item in the test csproj. Tests assert against the specific content of these documents. `Document2.docx` is a minimal synthetic fixture (a 2x3 table holding the `confReportType` and `StudyNumber` bookmarks); do not replace it with a real-world document.

## Build configuration

- Shared settings (target framework, `LangVersion`, `ImplicitUsings`, `AnalysisLevel`, version/company/product metadata) live in `Directory.Build.props`. There are no `AssemblyInfo.cs` files; assembly attributes come from csproj properties.
- `.editorconfig` is authoritative for formatting (4-space indent, CRLF, UTF-8 without BOM, file-scoped namespaces) and analyzer severities. `.gitattributes` stores text as LF in the repo and checks source/project files out as CRLF.
- `AnalysisLevel` is `latest-recommended`: the build must stay at zero warnings, and under `CI=true` warnings are errors.
- `Nullable` is deliberately **disabled**: the code predates nullable annotations and enabling it yields ~250 warnings.
- `GlobalSuppressions.cs` suppresses CA1010 because the collections intentionally implement only the non-generic `IEnumerable`.

## Architecture

- **`RangeWrapper`** is the central abstraction: a (start, end) pair of `OpenXmlElement`s. Every collection (`ParagraphWrapperCollection`, `TableWrapperCollection`, `BookmarkWrapperCollection`, `CommentWrapperCollection`, Table Row/Cell/Column collections) is constructed from a `RangeWrapper` and lazily queries elements between `Start` and `End` via `GetElements<T>(end)` — nothing is cached. `RangeWrapper.Text` dispatches on the start element's type (Paragraph, TableCell, TableRow, Table, BookmarkStart, CommentRangeStart, Comment) and emulates Word's markers: `\r` for end of paragraph, `\a` for end of cell/row. Setting `Text` is supported only for TableCell, Paragraph, and BookmarkStart.
- **`WordprocessingDocumentWrapper`** is the entry point (`Open`/`Create`); its `Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments` all build a `RangeWrapper` over `MainDocumentPart.Document`. It also has part-copying helpers (`CopyAllPartsTo`, `CopyLayoutTo`).
- **Element wrappers** (`ParagraphWrapper`, `TableWrapper`, `BookmarkWrapper`, `CommentWrapper`, ...) hold the underlying SDK element, expose it via a property of the same name, and expose a `Range`.
- **`OpenXmlExtensions`** holds the low-level tree-walking logic the wrappers depend on (`GetElements` between two elements, `GetLowestCommonAncestor`, `GetElementsBetween`, table merge/width helpers, `Paragraph.SetText`, etc.). Put new XML-traversal primitives here rather than in wrappers.
- Collections implement non-generic `IEnumerable` with a hand-written nested enumerator that indexes into the collection; follow that pattern for new collections.

## Conventions

- Follow `.editorconfig`. Code uses explicit types rather than `var`, and `String.Format(CultureInfo.InvariantCulture, ...)` rather than interpolation (CA1305 requires an explicit culture).
- The Open XML SDK 3.x enum-like types (e.g. `BreakValues`, `TableWidthUnitValues`) are structs, so compare them with `==` in `if`/`else if` rather than `switch`.
- Test classes are named `<Class>Test`, one per wrapper (`<Class>CollectionTest` for collections).
