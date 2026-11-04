# Excel counterpart: assessment

> **Status (November 2026, after Word 1.0.1).** An assessment, not a plan of record. Nothing here is built until the Word gap work ([word-om-gap-review.md](word-om-gap-review.md), steps 1 to 20) is done. SpreadsheetML details were checked against the Open XML SDK documentation on Microsoft Learn on 4 November 2026.

Question: can this repository grow an `OpenXml.Wrappers.Excel` that does for `.xlsx` files what `OpenXml.Wrappers.Word` does for `.docx` files (an Excel-object-model-like API over the Open XML SDK), and how should it be laid out?

## 1. Verdict

**Partly well positioned.** The names leave room for Excel, and the shared plumbing is small and internal. The Word engine, some public types and all of the packaging and documentation are Word-only, so Excel needs its own project, package and engine rather than a folder in the existing one.

What already fits:

* The namespace is `OpenXml.Wrappers.Word`, under the `OpenXml.Wrappers/Word` folder, so `OpenXml.Wrappers.Excel` sits beside it without a rename.
* The package id `E2e4.OpenXml.Wrappers` and the assembly name `OpenXml.Wrappers` are generic, and so is `Product` in `Directory.Build.props`.
* `Compat.cs` (`Guard`) and `Polyfills.cs` (`CallerArgumentExpressionAttribute`, `HashCode`, `ReferenceEqualityComparer` for `netstandard2.0`) are internal and have nothing to do with Word.
* `DocumentFormat.OpenXml` 3.x, the only dependency, already contains the SpreadsheetML classes (`DocumentFormat.OpenXml.Spreadsheet`, `SpreadsheetDocument`). Excel needs no new package.

## 2. What is bound to Word

| Piece | Why it is Word-only | Consequence for Excel |
|---|---|---|
| `OpenXmlExtensions` (internal) | Imports `DocumentFormat.OpenXml.Wordprocessing`; most members are about bookmarks, comments, paragraphs, runs and tables. | Excel gets its own internal extensions. Only the few element-tree helpers (`GetLowestCommonAncestor`, `GetElements`) are generic, and they are cheap to copy. |
| `ConversionHelper` (internal) | Converts points to twips and table percentages. SpreadsheetML uses points for row heights and character widths for columns, not twips. | Not shared. Excel needs a column-width conversion (characters of the default font to points and back) instead. |
| `WordColor`, `ColorIndex`, `NamedColor`, `ColorHelper`, `LineWidth`, `BorderLineWidth` (public) | Public types in `OpenXml.Wrappers.Word`. Moving them to a shared namespace is a breaking change (SemVer major), and type forwarding cannot change a namespace. | Excel defines its own: Excel's `Color` is the same BGR `Long` as Word's, but `ColorIndex` is a 56-entry palette, not `WdColorIndex`, and border weights are `XlBorderWeight`, not `WdLineWidth`. |
| `DocumentPropertiesWrapper`, `CustomDocumentProperty`, `CustomDocumentPropertyCollection`, `DocumentPropertyType` (public) | Public Word-namespace types whose constructors take a `WordprocessingDocument`. The parts they read (`docProps/core.xml`, `app.xml`, `custom.xml`) are the same in every Office package. | The reading and writing logic moves into shared internal code over `OpenXmlPackage`; the Word types stay as they are and call it, and Excel gets its own public types. |
| Package metadata, `README.md`, `CLAUDE.md` | `Description` and `PackageTags` describe Word; the packed README is the Word README; `CLAUDE.md` describes only the Word engine and the Word test fixtures. | Excel needs its own description, tags and package README, and a section (or file) of its own in `CLAUDE.md`. |
| The test project | `OpenXml.Wrappers.Test` is flat: one folder of Word test classes and `.docx`/`.docm` fixtures, with `InternalsVisibleTo` from the Word assembly. | A separate Excel test project (section 3). |
| The engine | `RangeWrapper` spans over a story, `StoryMap`, `SpanRegistry`, anchors, `UnitNavigator` and `FormattingResolver` model text that flows through paragraphs and runs, with character positions. | Nothing of it carries over. A worksheet is a sparse grid addressed by row and column, with formatting in a shared stylesheet (section 5). |

## 3. Recommended structure

```
OpenXml.Wrappers/               Word library, package E2e4.OpenXml.Wrappers (unchanged)
OpenXml.Wrappers.Test/          Word tests (unchanged)
OpenXml.Wrappers.Excel/         Excel library, package E2e4.OpenXml.Wrappers.Excel
OpenXml.Wrappers.Excel.Test/    Excel tests and .xlsx fixtures
shared/                         internal source linked into both libraries
  Compat.cs, Polyfills.cs, PackageProperties.cs (core, app and custom properties over OpenXmlPackage)
```

* **A separate project and package**, `OpenXml.Wrappers.Excel` (namespace `OpenXml.Wrappers.Excel`), packed as `E2e4.OpenXml.Wrappers.Excel`, with its own `<Version>` in its csproj starting at `0.1.0-preview.1`. Plan step 1 single-sources the Word version in its csproj, so each package then has exactly one version property of its own; `AssemblyVersion` moves per project with it.
* **A separate test project**, `OpenXml.Wrappers.Excel.Test`, with its own fixtures. Besides keeping the suites apart, this avoids a real conflict: the linked internal types (`System.HashCode`, `ReferenceEqualityComparer` and so on) exist once in each library, and a single test project that sees the internals of both would get two copies of each and ambiguous-type errors.
* **A `shared/` folder of internal source**, linked into both libraries with `<Compile Include="..\shared\*.cs" Link="Shared\%(Filename)%(Extension)" />`. Each assembly compiles its own internal copy, so neither package references the other, and moving `Compat.cs` and `Polyfills.cs` there changes nothing public. Only code with no format-specific types goes there.
* **The Word package stays untouched**: same id, same namespace, same public surface, same tags.
* **Releases**: Excel tags are `excel-vX.Y.Z`, checked by `eng/Check-Version.ps1 -TagPrefix excel-v` (the version check that plan step 1 introduces, given a tag prefix). The Word publish workflow triggers on `v*`, which an `excel-v` tag does not match, so a separate workflow (or job) packs only the Excel project. Publishing still needs the owner's say-so, as for Word.

Why this layout:

* **Independent SemVer.** Excel starts at a preview and will make breaking changes before 1.0; Word is at 1.x and makes none. One package would force a Word major for every Excel break, or freeze Excel's design at its first release. Separate packages also let a Word user take Word fixes without Excel's surface and size, and the other way round.
* **No new dependency.** Both libraries depend only on `DocumentFormat.OpenXml` 3.x, which already has the SpreadsheetML classes. A third "common" package would be a new dependency for every Word user, and a public one would make its types a SemVer contract of their own; linked internal source avoids both.

## 4. First milestone (0.1)

The goal is the code a macro or report generator writes most: open or create a workbook, find a sheet, write values, formulas and formats into ranges, and save. Indexes are 0-based, as in the Word library (`Worksheets[0]`, `Cells(0, 0)` is A1); A1 references keep Excel's notation, so `Range("A1")` and `Cells(0, 0)` are the same cell.

| Object | Members in 0.1 |
|---|---|
| `Workbook` | `Open`, `Create`, `Save`, `SaveAs`, `Close` (saving unless told not to, as the Word library does), `Worksheets`, `Names`, `BuiltinDocumentProperties`, `CustomDocumentProperties`, `Name`, `FullName`, `Path` |
| `Worksheets` | `Add` (before or after a sheet), `this[int]`, `this[string]`, `Count`, enumeration |
| `Worksheet` | `Name` (Excel's rules: at most 31 characters, none of `[ ] : * ? / \`, unique ignoring case), `Visible` (visible, hidden, very hidden), `Delete` (not the last sheet, and not the last visible one), `Index`, `Range("A1:B2")`, `Cells(row, column)`, `Rows`, `Columns`, `Names` (sheet-scoped) |
| `Range` | `Value`, `Value2`, `Formula`, `Text`, `NumberFormat`, `Font`, `Interior`, `Borders`, `HorizontalAlignment`, `MergeCells`, `Merge`, `UnMerge`, `Rows`, `Columns`, `Cells`, `Count`, `Row`, `Column`, `Address`, `Worksheet`, `ClearContents` |
| `Rows`, `Columns` (and a row or column range) | `RowHeight`, `ColumnWidth`, `Hidden` |
| `Font` | `Name`, `Size`, `Bold`, `Italic`, `Underline`, `Strikethrough`, `Color` |
| `Interior` | `Color`, `Pattern` (solid and none) |
| `Borders`, `Border` | `LineStyle`, `Weight`, `Color`, by edge |
| `Names`, `Name` | `Add`, `this[...]`, `Count`, `Delete`, `Name`, `RefersTo`, workbook and sheet scope |
| Document properties | The same built-in and custom properties as the Word library, from the shared code |

Behaviour that 0.1 settles:

* `Value2` is the stored value: a `double` for numbers and dates, `string`, `bool`, or an error value. `Value` is the same except that a date-formatted number reads as a `DateTime`, as in Excel. Writing a `DateTime` stores a serial number and gives the cell a date format if it has none.
* `Formula` writes the formula and no cached value (section 5); `Value` of a formula cell reads the cached result if the file has one.
* `Text` is the displayed text, so it needs number-format rendering. 0.1 renders `General`, the fixed, thousands, percent and scientific formats and the built-in date and time formats, and throws `NotSupportedException` for format codes it cannot render rather than guessing.
* A range read is a 2-D array for a multi-cell range, as in Excel.

## 5. Engine pieces Excel needs

Nothing of the Word engine is reused. Excel needs these, all internal:

* **A sparse grid.** `sheetData` holds `row` elements with a 1-based index `r`, and each row holds `c` cells with A1 references; rows and cells that hold nothing need not be written, and they must stay in ascending order. Reading never creates a row or cell; writing inserts it in order. Learn: [Working with sheets](https://learn.microsoft.com/office/open-xml/spreadsheet/working-with-sheets).
* **An A1 parser.** Cell, range, whole-column (`A:A`), whole-row (`1:1`), absolute (`$A$1`) and sheet-qualified (`'Sheet 1'!A1`) references, with Excel's limits (columns `A` to `XFD`, rows 1 to 1,048,576), and the conversion between A1 and 0-based row and column indexes.
* **A shared-strings manager.** A cell of type `s` holds a 0-based index into the shared string table (`sst`/`si`); the writer looks a string up before adding it. The reader also accepts inline strings (`inlineStr`), formula strings (`str`), booleans (`b`), errors (`e`) and ISO dates (`d`), which are the other `CellValues`. Learn: [Working with the shared string table](https://learn.microsoft.com/office/open-xml/spreadsheet/working-with-the-shared-string-table), [CellValues](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.cellvalues).
* **Stylesheet deduplication.** A cell's `s` attribute indexes a cell format (`cellXfs/xf`) that points at shared fonts, fills, borders and number formats. Setting `Font.Bold` on one cell must not change other cells that share its format, so a write finds or adds the font, then finds or adds the cell format, through a cache keyed on the values. The cache is built once per workbook and kept in step with the stylesheet. The stylesheet keeps the two fills Excel requires first (none and `gray125`).
* **`<cols>` span splitting.** Column widths and hidden flags are `col` elements that cover a range of columns (`min` to `max`). Setting the width of one column inside a span splits the span into up to three, and the spans must stay sorted and must not overlap. Learn: [Columns](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.columns), [Column.Min](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.column.min).
* **OADate handling.** Dates are serial numbers. In the default 1900 date system Excel treats 1900 as a leap year, so serial 60 is a 29 February 1900 that never existed; `DateTime.ToOADate`/`FromOADate` agree with Excel from 1 March 1900 (serial 61) on, and the converter handles the serials before it. Workbooks with `date1904` set count from 1904, 1,462 days later. Learn: [Excel incorrectly assumes that the year 1900 is a leap year](https://learn.microsoft.com/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year), [Workbook.Date1904](https://learn.microsoft.com/office/vba/api/excel.workbook.date1904).
* **No calculation engine.** The library does not evaluate formulas. Writing a formula clears the cell's cached value and sets `fullCalcOnLoad` on the workbook's `calcPr`, so Excel recalculates everything when it opens the file; the calculation chain part is dropped, because Excel rebuilds it and a stale one can list cells that no longer have formulas. Learn: [CalculationProperties.FullCalculationOnLoad](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.calculationproperties.fullcalculationonload), [Working with the calculation chain](https://learn.microsoft.com/office/open-xml/spreadsheet/working-with-the-calculation-chain).

Sheet visibility is the `state` attribute of the workbook's `sheet` element (`visible`, `hidden`, `veryHidden`; [SheetStateValues](https://learn.microsoft.com/dotnet/api/documentformat.openxml.spreadsheet.sheetstatevalues)). Merged cells are `mergeCell` references in the worksheet's `mergeCells` list; names are `definedName` elements in the workbook, with `localSheetId` for sheet scope.

## 6. Deferred

Not in 0.1, and each a milestone of its own if wanted: charts, pivot tables, conditional formatting, and tables (`ListObjects`). Also later: data validation, comments and notes, hyperlinks, images, freeze panes and other sheet views, print settings, protection, sorting and filtering.

## 7. Testing against Excel

The Word reference harness of plan step 2 compares what the library writes with what Word 365 writes for the same operation, driven through Word COM on a maintainer's machine. The same approach works for Excel through Excel COM: run the equivalent VBA-style operations in Excel, save, and compare the parts (styles, shared strings, worksheet XML) with the library's output, normalised the same way. Building the harness so that the Office application is a parameter makes the Excel version mostly reuse.
