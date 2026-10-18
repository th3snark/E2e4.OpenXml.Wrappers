# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A C# library (`OpenXml.Wrappers`, namespace `OpenXml.Wrappers.Word`) that wraps the Open XML SDK's WordprocessingML types to give a Word-object-model-like API (`Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments`, etc.). Solution file: `E2e4.OpenXml.Wrappers.slnx`. Hosted in the public GitHub repo `th3snark/E2e4.OpenXml.Wrappers` (default branch `main`). See `README.md` for usage. CI is `.github/workflows/ci.yml` (format check, Release build with warnings as errors, tests, pack to an artifact only; nothing is published to NuGet by it).

## Build & test

The library targets `net10.0`, `net8.0` and `netstandard2.0` (`Compat.cs` and `Polyfills.cs` hold the few newer BCL members the code uses, for `netstandard2.0`; use `Guard.ThrowIfNull` and friends rather than `ArgumentNullException.ThrowIfNull`, and avoid `Span`, `Index`/`Range` and `HashCode`-style APIs that `netstandard2.0` lacks without extra packages: `DocumentFormat.OpenXml` must stay the only dependency). The tests and samples target `net10.0`, `net8.0` and `net48`, so `net48` runs the `netstandard2.0` build (Windows only; CI runs on Windows). It builds with the .NET 10 SDK and needs the .NET 8 runtime and `DocumentFormat.OpenXml` 3.x; the `dotnet` CLI is all that is needed:

```bash
dotnet build E2e4.OpenXml.Wrappers.slnx
dotnet test E2e4.OpenXml.Wrappers.slnx
dotnet test E2e4.OpenXml.Wrappers.slnx --filter "FullyQualifiedName~ParagraphWrapperTest"
dotnet format E2e4.OpenXml.Wrappers.slnx --verify-no-changes
```

- Tests are **xUnit**. Per-test setup/teardown is the constructor / `IDisposable.Dispose`. Parallelization is disabled in `Properties\xunit.assemblyinfo.cs` because test classes share fixture files on disk (some opened editable).
- Tests open fixture documents in `OpenXml.Wrappers.Test\TestFiles\` (`Document1.docm`, `Document2.docx`, `Document3.docx`), referenced through constants in `Utils\TestFile.cs` and copied to the output directory by a `CopyToOutputDirectory` item in the test csproj. Tests assert against the specific content of these documents. `Document2.docx` is a minimal synthetic fixture (a 2x3 table holding the `confReportType` and `StudyNumber` bookmarks); do not replace it with a real-world document.
- `samples\OpenXml.Wrappers.Samples` is a console project of 21 examples (`Examples.All`: 13 Basic, 5 Intermediate, 3 Advanced) that each write a `.docx`; `SampleSupport.cs` has the shared helpers. `OpenXml.Wrappers.Test\SamplesTest.cs` runs every example and validates the output with `OpenXmlValidator`, and asserts the 12/5/3 counts, so add a test-visible entry to `Examples.All` for any new example. Keep them openable in real Word.
- Licensed MIT (`LICENSE`); the icon is in `assets\` (`icon.png` is packed into the NuGet package). Nothing is published to NuGet or elsewhere until the user says so.

## Build configuration

- Shared settings (target framework, `LangVersion`, `ImplicitUsings`, `AnalysisLevel`, version/company/product metadata) live in `Directory.Build.props` (the target frameworks are in each project file). There are no `AssemblyInfo.cs` files; assembly attributes come from csproj properties.
- `.editorconfig` is authoritative for formatting (4-space indent, CRLF, UTF-8 without BOM, file-scoped namespaces) and analyzer severities. `.gitattributes` stores text as LF in the repo and checks source/project files out as CRLF.
- `AnalysisLevel` is `latest-recommended`: the build must stay at zero warnings, and under `CI=true` warnings are errors.
- `Nullable` is deliberately **disabled**: the code predates nullable annotations and enabling it yields ~250 warnings.

## Architecture

- **`RangeWrapper`** is the central abstraction: a (start, end) pair of `OpenXmlElement`s (`StartElement`/`EndElement`), or a span of characters (`Start`/`End` integers over a story, see `RangeWrapper.Span*.cs`, `StoryMap` and `SpanAnchors.cs`). A span stores two `Anchor`s (a text element and an offset), not integers, so it follows edits; the library tells `SpanRegistry` when it splits a text element, edits one in place, splits or joins a paragraph or removes text, and the registry corrects the anchors of every live span. Work inside one paragraph, table cell or table uses a `StoryMap` of just that part (`GetScoped`), and a span across consecutive paragraphs of one parent uses a window map (`StoryMap.BuildWindow`); only the rest builds the map of the whole story. `Start`/`End` and `SetRange` count through the story without a map (`StoryMap.Enumerate`). Anchors are immutable; each also remembers the pieces of text near it on both sides (`Anchor.Before`/`After`) so that a range can collapse to where its text was when something removes the text behind the library's back, and `SpanRegistry` finds the ranges to correct by the element their ends are on (it is thread-safe). `StoryMap.Enumerate` is the single description of what takes up positions (text, tabs, breaks, marks, field marks and codes, objects); `UnitNavigator` moves by units block by block; `PositionIndex` is the public bulk reader of positions. Every collection (`ParagraphWrapperCollection`, `TableWrapperCollection`, `BookmarkWrapperCollection`, `CommentWrapperCollection`, Table Row/Cell/Column collections) is constructed from a `RangeWrapper` and lazily queries elements between `Start` and `End` via `GetElements<T>(end)` — nothing is cached. `RangeWrapper.Text` dispatches on the start element's type (Paragraph, TableCell, TableRow, Table, BookmarkStart, CommentRangeStart, Comment) and emulates Word's markers: `\r` for end of paragraph, `\a` for end of cell/row. Setting `Text` is supported only for TableCell, Paragraph, and BookmarkStart.
- **`RangeWrapper` is a partial class**, one file per feature: `.Insert` (`InsertBefore`/`InsertAfter`), `.Delete`, `.Find`, `.Font`, `.ParagraphFormat` and `.Style`. The feature objects (`FindWrapper`, `FontWrapper`, `ParagraphFormatWrapper`) are thin views over a range (a span splits runs so formatting can apply to part of one). They write direct formatting but read the effective formatting like Word does, through the internal `FormattingResolver` (document defaults, table style, paragraph style, character style, direct, with Word's toggle rules; `.Direct` reads direct formatting only), return null for mixed values, and write into the schema order (`SetRunProperty` in `OpenXmlExtensions` for run properties; the typed properties of `ParagraphProperties` for paragraphs). All public members use points (Word's unit) and convert to twips inside (`ConversionHelper.PointsToTwips`); only the helpers that work directly on Open XML elements (`OpenXmlExtensions`, `ConversionHelper.CentimeterToDxa`) keep twips/dxa.
- **Built-in styles** are defined in `BuiltInStyles` (the internal catalogue, with the `BuiltInStyle` enum using `WdBuiltinStyle` values). `Styles[name]` and applying a style by name (`StyleHelper.FindStyleId`) add the definition to the document on demand; `StyleHelper.GetName` maps stored names ("heading 1") to the names Word shows ("Heading 1"). Add new definitions to the catalogue in schema order and keep `BuiltInStylesTest` (which validates them all) passing.
- **Indexes are 0-based** everywhere (collection indexers, `Index`, `RowIndex`, `ColumnIndex`, `Table.Cell`) by decision, for C# developers; do not introduce 1-based indexes or an abstraction over the base.
- **`WordprocessingDocumentWrapper`** is the entry point (`Open`/`Create`); its `Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments` all build a `RangeWrapper` over `MainDocumentPart.Document`. It also has part-copying helpers (`CopyAllPartsTo`, `CopyLayoutTo`).
- **Element wrappers** (`ParagraphWrapper`, `TableWrapper`, `BookmarkWrapper`, `CommentWrapper`, ...) hold the underlying SDK element, expose it via a property of the same name, and expose a `Range`.
- **`OpenXmlExtensions`** holds the low-level tree-walking logic the wrappers depend on (`GetElements` between two elements, `GetLowestCommonAncestor`, `GetElementsBetween`, table merge/width helpers, `Paragraph.SetText`, etc.). Put new XML-traversal primitives here rather than in wrappers.
- Collections implement `IEnumerable<TWrapper>` (the non-generic `IEnumerable` is an explicit implementation) with an iterator method over a `GetElementList()` snapshot taken when enumeration starts; follow that pattern for new collections. The indexer and `Count` re-query the tree each time, so they are live.
- Wrappers implement `IEquatable<T>` and `==`/`!=` by reference equality of the wrapped element(s); give any new wrapper the same `Equality` region.
- `WordprocessingDocumentWrapper` is `IDisposable`; `Close()` is the same as `Dispose()`. `Create` initialises the main document and comments roots.

## Conventions

- Follow `.editorconfig`. Code uses explicit types rather than `var`, and `String.Format(CultureInfo.InvariantCulture, ...)` rather than interpolation (CA1305 requires an explicit culture).
- The Open XML SDK 3.x enum-like types (e.g. `BreakValues`, `TableWidthUnitValues`) are structs, so compare them with `==` in `if`/`else if` rather than `switch`.
- Test classes are named `<Class>Test`, one per wrapper (`<Class>CollectionTest` for collections).
- The tests for a `RangeWrapper` feature file are in `RangeWrapper<Feature>Test` (for example `RangeWrapperFindTest`). Most use in-memory element trees rather than the fixture documents; tests that need styles or comments create a temporary document with `WordprocessingDocumentWrapper.Create`.
