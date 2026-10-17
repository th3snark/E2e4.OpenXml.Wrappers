# Changelog

All notable changes are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/) and the project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- Word-object-model API over Open XML: ranges (insert, delete, find and replace, font, paragraph format, style, `InsertParagraphAfter`/`Before`), paragraphs, tables, rows, cells (merge, split, borders), columns, bookmarks (including tables), comments, fields, sections and page setup, headers and footers, styles, built-in and custom document properties, navigation and `Parent`.
- 20 runnable examples in `samples/OpenXml.Wrappers.Samples`, validated by the test suite.
- MIT license, package metadata and icon.

### Changed
- `Paragraph.Delete`, `Row.Delete` and `Table.Delete` use `Range.Delete`, so bookmarks and comments anchored in them are removed and a cell, body or table is never left invalid. `InsertSectionBreak` supports all five section types. `Cell.MergeRight`/`MergeDown` respect grid spans, and `Paragraph.KeepNext` has a getter.
- `PageSetup.OddAndEvenPagesHeaderFooter`; `Range.Delete` for headers and footers; `Cell.Merge` absorbs vertical merges inside the rectangle; `Column.Delete` and `Column.Width` work with merged cells; borders are read from the table and table style; `Paragraph.Text` no longer ends with the cell marker.
- `Cell.Split` handles vertically merged cells (into rows or columns) and `Cell.Merge` cuts vertical merges that stick out of the rectangle; borders from the first/last row and column conditional formatting of a table style are read; `KeepNext` of rows and cells can be read.
- New documents are created with `compatibilityMode` 15 so Word does not open them in Compatibility Mode.
