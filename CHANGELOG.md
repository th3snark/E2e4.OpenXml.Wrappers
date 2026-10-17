# Changelog

All notable changes are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/) and the project uses [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- Word-object-model API over Open XML: ranges (insert, delete, find and replace, font, paragraph format, style, `InsertParagraphAfter`/`Before`), paragraphs, tables, rows, cells (merge, split, borders), columns, bookmarks (including tables), comments, fields, sections and page setup, headers and footers, styles, built-in and custom document properties, navigation and `Parent`.
- 20 runnable examples in `samples/OpenXml.Wrappers.Samples`, validated by the test suite.
- MIT license, package metadata and icon.

### Changed
- New documents are created with `compatibilityMode` 15 so Word does not open them in Compatibility Mode.
