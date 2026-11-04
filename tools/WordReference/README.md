# Word reference harness

Tools that compare what the library writes with what Word itself writes. They need Word (Word 365 on Windows) and PowerShell 7, so they
run on a developer's machine; the tests and CI only read the snapshots committed under
`OpenXml.Wrappers.Test/TestFiles/WordReference/`.

| File | What it does |
| --- | --- |
| `New-WordReferences.ps1` | Runs every scenario in Word and rewrites the snapshots. `-Scenario <name>` runs only some. |
| `scenarios.ps1` | The scenarios: named scriptblocks of Word object model calls on a new document. |
| `Test-OpensInWord.ps1` | Opens every sample document in Word; fails on a COM error or a repair and reports `Range.Text` differences. `-Folder` opens other documents. |
| `WordAutomation.ps1` | Shared helpers: starting and quitting Word, guarding `Normal.dotm`. |
| `Tool/` | A small console program (not in the solution) that links the test project's `MarkupNormalizer.cs`, so that the snapshots and the tests share one normaliser. It also reads a document's text through the library for `Test-OpensInWord.ps1`. |

## Regenerating the snapshots

```powershell
pwsh tools/WordReference/New-WordReferences.ps1
```

Close Word first: the script starts its own Word (`WINWORD.EXE /a /automation`) and attaches to it through COM, and refuses to run while
another Word is open, so that it never drives your documents. `/a` means add-ins and your `Normal.dotm` are not loaded, so the documents
are what a fresh Word writes. That fallback still has the Office 2013 - 2022 theme (Calibri), so each new document gets the current Office
theme (`Document Themes 16\Office Theme.thmx`, Aptos) applied first, as a new document in Word 365 has it. For the run, `Application.UserName` and `UserInitials` are set to `Reference` and `R` (and
`Options.UseLocalUserInfo` to true, so that comments do not take the signed-in account's name or address); they are restored afterwards.
`Normal.dotm` is copied to `%TEMP%\WordReference-<time>\` first and compared afterwards (restored and reported if it changed), and the
Office user name in the registry is checked too.

The normaliser puts adjacent runs of the same formatting together, on both sides, because Word keeps runs apart that differ only in their revision
ids; `RunMarkupTest` checks directly that the library merges them.

Each scenario's document is saved as .docx; every XML part of it is normalised and written to `<scenario>/<part>`, and the Word build goes in
`TestFiles/WordReference/README.md`. The parts of `blank-document` and the styles of `built-in-styles` also go to
`OpenXml.Wrappers/Word/Templates`, without revision and document ids, statistics or the application's name: these are what the library's
`Create` and built-in styles write. Review the diff before committing: the snapshots reflect the machine's Word build and locale (page size,
`w:lang`, `themeFontLang`).

## Normalisation

`OpenXml.Wrappers.Test/Utils/MarkupNormalizer.cs` removes what changes on every save (`w:rsid*`, `w14:paraId` and `w14:textId`, proofing
marks, `w:lastRenderedPageBreak`, the `_GoBack` bookmark), replaces ids with ordinal placeholders (`{rId1}`, `{bookmark1}`, `{comment1}`,
`{para1}`, ...) so that what refers to what is kept, replaces dates, document ids, hashes and salts with fixed placeholders, sorts attributes,
namespace declarations, relationships and content types, drops unused namespace declarations and `mc:Ignorable` prefixes, makes
relationship targets absolute and writes indented XML. Part files are named after their path below `word/` (`document.xml`,
`theme.theme1.xml`, `document.xml.rels`), plus `content-types.xml` and `package.rels`.

## Adding a scenario

1. Add a scriptblock to `scenarios.ps1` (Word's indexes are 1-based).
2. Add a case with the same name to `MarkupConformanceTest`, doing the same through the library (0-based), with the parts or elements to
   compare and, if the library cannot match Word yet, the plan step that will make it.
3. Run `New-WordReferences.ps1 -Scenario <name>` and commit the new folder.
