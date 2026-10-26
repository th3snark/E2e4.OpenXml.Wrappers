# Changelog

All notable changes are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/) and the project uses [Semantic Versioning](https://semver.org/).

## [1.0.1] - 2026-10-26

### Fixed
- The package icon now shows on nuget.org: the packed README is a copy with the icon as an absolute image link and relative links made absolute, pinned to the release tag. The README in the repository is unchanged.

### Changed
- No change to the public API or behaviour. The source now follows the predefined type keyword rule (`string` rather than `String`), enforced in CI, and the performance tests compare run time at two document sizes instead of using fixed time limits, so they no longer fail on slow machines.

## [1.0.0] - 2026-10-18

The first release.

### Added
- A Word-object-model-like API over the Open XML SDK for `.docx` files, with `DocumentFormat.OpenXml` as the only dependency. Builds for `net10.0`, `net8.0` and `netstandard2.0` (so .NET Framework 4.6.1 and later and every .NET version are supported); the tests run on `net10.0`, `net8.0` and `net48`.
- **Document**: `Open`, `Create`, `Save`, `SaveAs`, `Close(SaveOptions)`, `Content`, `Range()` and `Range(start, end)`, `Name`, `FullName`, `Path`, `ReadOnly`, `PageSetup`, `Characters`/`Words`/`Sentences`, `ComputeStatistics`, `CompatibilityMode`, `Protect`/`Unprotect`/`ProtectionType`, `Variables`, built-in and custom document properties. `TrackRevisions` can be read and switched off.
- **Ranges** of elements or of characters, with Word's positions: `Start`, `End`, `SetRange`, `Collapse`, `MoveStart`, `MoveEnd`, `Move`, `Expand`, `Duplicate`, `Characters`, `Words`, `Sentences`, `Text`, `InsertBefore`/`InsertAfter`, `InsertParagraphBefore`/`After`, `InsertBreak`, `Delete()` and `Delete(unit, count)`. A span is tied to its text, so it follows edits however they are made.
- **Find and replace** (`Find.Execute` with Word's parameters in Word's order, wildcards, `^` codes, backwards search, wrap).
- **Formatting**: `Font` and `ParagraphFormat` (effective values as Word reads them, `.Direct` for direct formatting, `Reset()`, tab stops), `Style`/`StyleObject`, colours (`WordColor`, `ColorHelper.RGB`, `NamedColor`, `ColorIndex`), `LineWidth`/`BorderLineWidth`, `Borders` and `Shading`.
- **Paragraphs, tables, rows, columns and cells**: adding, deleting, merging, splitting, sizing, shading, borders, vertical alignment, repeating header rows and the rest of the members listed in the README.
- **Bookmarks, comments (with replies and `Done`), fields (`Update` works out the fields that need no page layout), sections, page setup, headers and footers, page numbers, styles (the built-in styles always exist) and document properties.**
- 21 runnable examples in `samples/OpenXml.Wrappers.Samples`, validated against the Open XML schema by the test suite.
- MIT license, package metadata and icon.

### Not supported
Page layout (pages, lines, `Information`, `Selection`), tracked changes (`TrackRevisions = true` throws), hyperlinks, footnotes and endnotes, lists, content controls, pictures and shapes, tables of contents, and the other limits in the README under "Not supported".
