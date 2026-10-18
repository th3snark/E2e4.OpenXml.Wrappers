# Changelog

All notable changes are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/) and the project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Changed
- Word names (breaking, before the first release): `Font.Strikethrough` is `StrikeThrough`, `KeepNext` of paragraphs, rows and cells is `KeepWithNext`, `Comment.Initials` is `Initial`, and `Comment.Range` is now the text of the comment while the commented range is `Comment.Scope`, as in Word. `Styles.Delete(style)` is `Style.Delete()`. `StyleWrapper.Name` is the name Word shows ("Heading 1").
- The values of `TextUnit`, `CollapseDirection`, `BorderPosition` and `DocumentPropertyType` are those of Word's `WdUnits`, `WdCollapseDirection`, `WdBorderType` and `MsoDocProperties`; `ParagraphAlignment` has `Distribute`.
- `Font` and `ParagraphFormat` read the effective formatting, as in Word (document defaults, table, paragraph and character styles, direct formatting), including for `Style.Font` and `Style.ParagraphFormat`; `.Direct` reads only the direct formatting. Unset size, name and colour read as Word's defaults (10 points, Times New Roman, auto).
- `CopyAllPartsTo` replaces the parts the target already has, and `CopyLayoutTo` drops header and footer references the source does not have. The collection indexers of bookmarks and comments report the index in their out of range message.

### Added
- `Range.InsertBreak(BreakType)` (page, column, line and section breaks); column breaks are the shift out character in text, as in Word.
- `Paragraph` has the properties of its format (`Alignment`, `LeftIndent`, `RightIndent`, `FirstLineIndent`, `LineSpacing`, `LineSpacingRule`, `KeepTogether`, `PageBreakBefore`, `WidowControl`, `Style`), `Next(count)` and `Previous(count)`; `Paragraphs` has `First`, `Last`, `Format` and `Style`.
- `Font`: `UnderlineColor`, `Superscript`, `Subscript`, `AllCaps`, `SmallCaps`, `Hidden`, `DoubleStrikeThrough`, `Emboss`, `Engrave`, `Outline`, `Shadow`, `Spacing`, `Scaling`, `Position`, `Kerning` and `Reset()`. `Font.Underline` is an `UnderlineStyle` (breaking, before the first release).
- `ParagraphFormat`: `OutlineLevel`, `NoLineNumber`, `MirrorIndents`, `Hyphenation`, `SpaceBeforeAuto`, `SpaceAfterAuto`, `TabStops` and `Reset()`.
- `Find`: `Forward`, `Wrap`, `MatchWildcards` (Word patterns), `^&`, `^w`, `^#`, `^$`, `^?`, `^s`, `^+`, `^=`, `^-` and `^n`, `ClearFormatting()`, and `Execute` with Word's parameters in Word's order (a parameter that is left out leaves its property as it is).
- The built-in styles: `Styles["Heading 1"]` and `Styles[BuiltInStyle.Heading1]` (the `BuiltInStyle` enum has the values of `WdBuiltinStyle`) find the style even when the document has not used it yet, and applying one by name adds its definition to the document.
- `Font.Direct` and `ParagraphFormat.Direct`.

## [1.0.0-preview.1] - 2026-10-17

### Added
- Builds for `net10.0`, `net8.0` and `netstandard2.0` (so .NET Framework 4.6.1 and later and every .NET version are supported), with `DocumentFormat.OpenXml` as the only dependency. The tests run on `net10.0`, `net8.0` and `net48`.
- `Section.Parent` (the sections of the document) and `Sections.Parent` (the range of the document), so every collection and wrapper has a `Parent`.
- Word-object-model API over Open XML: ranges (insert, delete, find and replace, font, paragraph format, style, `InsertParagraphAfter`/`Before`), paragraphs, tables, rows, cells (merge, split, borders), columns, bookmarks (including tables), comments, fields, sections and page setup, headers and footers, styles, built-in and custom document properties, navigation and `Parent`.
- 21 runnable examples in `samples/OpenXml.Wrappers.Samples`, validated by the test suite.
- MIT license, package metadata and icon.

### Changed
- Character ranges (breaking, before the first release): `Range.Start`/`End` are character positions, the element properties are `StartElement`/`EndElement`, `Document.Range` is `Document.Content` and `Document.Range(start, end)` a span. Added `SetRange`, `Collapse`, `MoveStart`, `MoveEnd`, `Move`, `Expand`, `Characters`, `Words` and `Sentences`. `Find.Execute` narrows the range to the match, matches can span paragraphs and replacements can contain tabs, breaks and paragraph marks. `Font`, `Style`, `Delete`, `InsertBefore`/`After`, `Bookmarks.Add`, `Comments.Add` and `Fields.Add` work on any span. A span is anchored to its text, so it follows the text when anything else changes the document, and operations inside one paragraph only look at that paragraph. Positions count field marks and codes (including hyperlinks and `w:fldSimple`), content controls, tracked-deleted text, pictures and references as Word does. `PositionIndex` reads the positions of many ranges at once. Moves by units, windows across tables and finds with `^p` work block by block. Deleting across the edge of cells, and setting the text of such a range, behave as in Word. `Paragraphs[i]` and `Count` no longer look inside paragraphs, and paragraphs in text boxes are not counted, as in Word.
- Units (breaking, before the first release): every public width and spacing is in points. `ParagraphWrapper.SpaceBefore`/`SpaceAfter` are `double`, `TableCellWrapper.CellWidth` is now `Width` (`double`), `TableColumnWrapper.Width` and `TableWrapper.Width` are `double`. Values that are negative, not a number or more than 31680 points are rejected.
- Removed (breaking, before the first release) the members that only duplicated newer ones: `ParagraphWrapper.SetFontBold`/`SetFontColor` and `TableCellWrapper.SetFontBold`/`SetFontColor` (use `Range.Font`), `TableCellWrapper.SetShadingFillColor` (use `BackgroundColor`), `TableRowWrapper.SetTopBorder`/`SetBottomBorder` (use `Borders`) and the `SetFontBold`/`SetFontColor` extension methods on run properties.
- `Paragraph.Delete`, `Row.Delete` and `Table.Delete` use `Range.Delete`, so bookmarks and comments anchored in them are removed and a cell, body or table is never left invalid. `InsertSectionBreak` supports all five section types. `Cell.MergeRight`/`MergeDown` respect grid spans, and `Paragraph.KeepNext` has a getter.
- `PageSetup.OddAndEvenPagesHeaderFooter`; `Range.Delete` for headers and footers; `Cell.Merge` absorbs vertical merges inside the rectangle; `Column.Delete` and `Column.Width` work with merged cells; borders are read from the table and table style; `Paragraph.Text` no longer ends with the cell marker.
- `Cell.Split` handles vertically merged cells (into rows or columns) and `Cell.Merge` cuts vertical merges that stick out of the rectangle; borders from the first/last row and column conditional formatting of a table style are read; `KeepNext` of rows and cells can be read.
- New documents are created with `compatibilityMode` 15 so Word does not open them in Compatibility Mode.
