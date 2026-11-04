# Changelog

All notable changes are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/) and the project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Behaviour changes
- `Create` (and `InitializeMainDocumentPart`) now make what Word 365 writes for a new blank document, rather than a body and a compatibility setting: styles with document defaults and latent styles (33 styles, among them Normal, Heading 1 to 9, Title and Quote), the current Office theme (Aptos), font table, settings, web settings and document properties. Text in a new document is therefore Aptos 12 point rather than Times New Roman 10 point, and margins are one inch all round.
- **Code that adds one of these parts itself after `Create` now throws** "Only one instance of the type is allowed for this parent": change the part that is there (`MainDocumentPart.StyleDefinitionsPart`) instead of calling `AddNewPart<StyleDefinitionsPart>()`. Code that counted on a new document having no styles (`Styles.Count`, `Styles.Exists`) sees Word's. A document that was opened rather than created is not affected.
- A new document takes its paper size, header and footer distance, column spacing, language and decimal and list separators from `CultureInfo.CurrentCulture`, as Word takes them from Windows' regional settings: A4 and 1.25 cm in en-GB, Letter and half an inch in en-US. Set the page up explicitly if a document must not depend on the machine.
- The built-in styles are Word 365's own definitions, so they look different from 1.0's: Heading 1 is 20 point rather than 16, for instance. Adding a paragraph style also adds its linked character style ("Heading 1 Char"), and applying a style unhides it, as in Word.
- `Font.Kerning` of unformatted text in a new document reads 1, from Word's document defaults.
- Formatting is written as Word writes it, so code that reads the XML directly may see a difference. On and off properties are bare (`<w:b/>`). A value the text or paragraph has anyway from its styles is not written: `Bold = false`, `Size` back to the inherited size, `Alignment = Left`, `Style = "Normal"` and so on remove the direct formatting, and `w:val="0"` is only written where a style would otherwise turn the property on. `Font.Bold`, `Italic` and `Size` no longer also set the complex script properties (`w:bCs`, `w:iCs`, `w:szCs`), and `Font.Name` only sets the font of Latin text (`w:ascii`, `w:hAnsi`).
- Runs with the same formatting are merged after formatting and editing, so a `Run` taken from the XML before an edit may no longer be in the document.
- A new paragraph follows Word's rules rather than being plain: after a paragraph that has one following it, it has the formatting of the following one; at the end of a story it continues the paragraph, or takes the next style (a paragraph after a Heading 1 is Normal). `Paragraphs.Add` and `InsertParagraphBefore` do the same. Its paragraph mark takes the formatting of the character next to it.
- `Range.InsertBreak` on a span of characters writes a page break in a paragraph of its own, with the text after it in a new paragraph, and a column break at the start of a paragraph, as Word does. The text of the range around it changes accordingly.
- The paragraph mark that ends a section reads as `\f` in `Range.Text` and `Paragraph.Text`, as in Word, rather than `\r`.
- `Fields.Add` writes `w:fldSimple` when the result can be worked out and the text has no formatting of its own, as Word does, and field codes are spaced as Word spaces them, so `Field.Code.Text` of a typed field is `TITLE  \* MERGEFORMAT`. A field whose result needs page layout has no result until Word updates it.
- `ParagraphFormat.WidowControl` reads true when nothing says otherwise, which is Word's default. It used to read false.

- **`Tables.Add` makes the table Word makes.** By default (`Word8TableBehavior`) that means **no borders**, no style and fixed columns, where it used to draw borders. `Word9TableBehavior` gives the Table Grid style. Pass that, or set `Table.Borders.Enable` or `Table.Style`, to keep borders. The width is `w:tblW` automatic rather than a number of twips, so `Table.PreferredWidth` reads 0 where it used to read the width; `Table.Width` works it out from the columns. A table in a cell fits inside the margins of the cell.
- `Comment.Date` is local time, as Word stores it, rather than UTC; the UTC time is kept beside it. A comment on a collapsed range takes in the word it is in, and a comment with no range or no text has no added space.
- `PageNumbers.Add` writes what Word writes: the number in a frame at the start of the header or footer, with the text that was there indented to make room. The paragraph is not aligned, so `Paragraph.Alignment` does not say where the number is.
- Setting `Row.HeadingFormat` also sets the rows above, and clearing it clears the rows below, because heading rows are the first rows of a table. `Rows.Alignment` also aligns each row. Setting `Table.PreferredWidth` scales the columns. Adding a column makes `w:tblW` the total width.
- A section added with `Sections.Add` or `InsertSectionBreak` inherits its headers and footers from the section before rather than having its own references to the same parts, and a break takes the properties of the section it splits rather than those of the last section.
- `Row.Copy` leaves bookmarks and comments with the original, because they can only be in one place.

### Fixed
- **A password given to `Protect` now works in Word.** It was written with a hash that Word could not check, in attributes it does not use (`w:hashValue`, `w:saltValue`), so Word could not unprotect the document with it. It now uses Word's algorithm, checked against documents Word protected and by unprotecting in Word. `Unprotect` still reads what 1.0 wrote.
- Comments write everything Word writes for them: `commentsExtended`, `commentsIds`, `commentsExtensible` and `people`, with their entries removed when a comment is deleted, including when it goes with the text it is on. The end of the range of a reply follows the reference of the comment before it.
- Bookmark ids are unique across the document, not just within a part.
- A header unlinked with `LinkToPrevious = false` keeps what it refers to (pictures, links), and bookmarks stay with the original. `CopyAllPartsTo` and `CopyLayoutTo` copy what the parts refer to, and `CopyLayoutTo` leaves no references to headers it has deleted.
- `Cell.Delete` and `Column.Delete` remove bookmarks and comments in the cells rather than leaving half of them behind.
- Headers and footers are in the Header and Footer styles. Page borders are measured from the edge of the page, as Word's are; without `w:offsetFrom` they were drawn in the wrong place. Line numbers show, because `w:countBy` is written.
- Values Word does not write are left out: `w:orient="portrait"`, `w:type="nextPage"`, `w:hRule="atLeast"`, `w:equalWidth`, an empty `w:trPr`, and `w:tblInd` of 0. `w:tcW` and `w:gridSpan` are always in schema order.
- Comments add the styles they use (Comment Reference, Comment Text and Comment Subject). They used to refer to styles the document did not have.
- Core properties are written to `/docProps/core.xml`, where Word keeps them, also when they are set on a document that has none. They used to go in a part named by the packaging layer (`/package/services/metadata/core-properties/*.psmdcp`).
- Indentation keeps Word's attribute names (`w:left`, `w:right`, `w:leftChars`, `w:rightChars`) when a document is saved. The Open XML SDK writes back the ones it has loaded as `w:start` and so on, which is not what Word writes.
- A default style added to a document with latent styles goes after them, as the schema requires.
- `xml:space="preserve"` is only written on text that starts or ends with white space.
- A carriage return in text given to a paragraph, cell, header, footer or the document starts a new paragraph. It used to end up in the `w:t` of the run.
- Setting `Text` on a paragraph or bookmark leaves no empty hyperlinks, simple fields or content controls behind, and the text of a paragraph stays inside a bookmark round it.
- Joining paragraphs keeps a section break of the second one and drops the one of the first, as deleting a section break does in Word. `Font` and `ParagraphFormat` changes leave no empty `w:pPr`. `Paragraph.DeleteBreaks` leaves no empty runs, a line break is written as `<w:br/>`, and a next page section break no longer writes `w:type`.

### Changed
- The version now has one source, `<Version>` in the library's csproj; `AssemblyVersion` and `FileVersion` are derived from it, and the published package takes its version from it rather than from the tag. CI and the publish workflow check that the package, the assembly versions, this changelog and the release tag agree (`eng/Check-Version.ps1`), prerelease tags such as `v1.2.0-preview.1` can be published, and publishing a version that is already on nuget.org now fails instead of being skipped. No change to the library itself.

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
