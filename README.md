# E2e4.OpenXml.Wrappers

[![CI](https://github.com/th3snark/E2e4.OpenXml.Wrappers/actions/workflows/ci.yml/badge.svg)](https://github.com/th3snark/E2e4.OpenXml.Wrappers/actions/workflows/ci.yml) [![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

<img src="assets/icon.png" alt="E2e4.OpenXml.Wrappers icon" width="96" align="right" />

A .NET library that wraps the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) WordprocessingML types in a Word-object-model-like API: `Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments` and friends.

## What it is for

The Open XML SDK gives complete, exact access to a `.docx` file, but it works at the level of the file format: a visible paragraph is a `w:p` holding `w:r` runs holding `w:t` text, a merged cell is a `gridSpan` plus a `vMerge`, and the element order is fixed by the schema. That is the right level for a general tool, and a poor one for the common job of "open this document, change these things, save it".

People who have written Word macros already know a better vocabulary: `ActiveDocument.Tables(1).Cell(2, 3).Range.Text`, `Range.Find.Execute`, `Bookmarks("Total").Range`, `Range.Font.Bold`. This library offers that vocabulary over Open XML, so code (and the people writing it) can work with the document as Word presents it, **without Word installed**, on a server, in a container or in a build pipeline.

- **Familiar names**: `Range`, `Paragraphs`, `Tables`, `Rows`, `Cells`, `Columns`, `Bookmarks`, `Comments`, `Fields`, `Sections`, `PageSetup`, `Styles`, `Headers`/`Footers`, `BuiltInDocumentProperties`.
- **Safe by construction**: output follows the schema order, so documents open in Word without repair. The sample documents are validated with the SDK's `OpenXmlValidator` in the test suite and have been opened in Word itself.
- **Not a replacement for the SDK**: every wrapper exposes the element it wraps, so you can drop down to the SDK for anything the library does not cover and carry on.
- **Honest about the gaps**: where Word's behaviour is not implemented the library throws rather than quietly doing something else (see [Differences from the Word object model](#differences-from-the-word-object-model)).

Typical uses: filling templates and mail merges, generating reports, invoices and agreements from data, find and replace across documents, reviewing documents by adding comments, and porting existing Word VBA to a service.

## Requirements

- .NET 10 SDK
- `DocumentFormat.OpenXml` 3.x (restored automatically)

## Usage

```csharp
using OpenXml.Wrappers.Word;

using (WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open("report.docx", isEditable: true))
{
    foreach (ParagraphWrapper paragraph in document.Paragraphs)
    {
        Console.WriteLine(paragraph.Text);
    }

    document.Bookmarks["Title"].Range.Text = "Quarterly report";
    document.Save();
}
```

`Create` makes a new document (from a path or a stream) with an empty body, and `SaveAs` writes a copy to a new file. Editable documents are also saved when disposed or closed. New documents are marked as Word 2013+ (`compatibilityMode` 15), so Word does not open them in Compatibility Mode.

## Examples

The [`samples`](samples/OpenXml.Wrappers.Samples) project holds 21 runnable examples. Each one builds a complete `.docx` that opens in Word, and each is checked by the test suite (it is created and validated against the Open XML schema). Run them all with:

```bash
dotnet run --project samples/OpenXml.Wrappers.Samples [output folder]
```

The documents are written to `output` (or the folder you give). The code in each file is commented to be read as a tutorial; [`SampleSupport.cs`](samples/OpenXml.Wrappers.Samples/SampleSupport.cs) holds the small helpers they share (styles for a new document, a table builder).

**Basic: one object each**

| Example | Shows |
| --- | --- |
| [Range](samples/OpenXml.Wrappers.Samples/Basic/RangeExample.cs) | Insert, find and replace, format and delete |
| [Paragraph](samples/OpenXml.Wrappers.Samples/Basic/ParagraphExample.cs) | Adding, formatting and moving between paragraphs |
| [Table](samples/OpenXml.Wrappers.Samples/Basic/TableExample.cs) | Adding, filling, sizing, aligning and bordering a table and its columns |
| [Row](samples/OpenXml.Wrappers.Samples/Basic/RowExample.cs) | Height, repeating header rows, inserting rows, row borders |
| [Cell](samples/OpenXml.Wrappers.Samples/Basic/CellExample.cs) | Merge, split, shading, alignment, borders |
| [Bookmark](samples/OpenXml.Wrappers.Samples/Basic/BookmarkExample.cs) | Naming a place and reading or replacing it later |
| [Comment](samples/OpenXml.Wrappers.Samples/Basic/CommentExample.cs) | Adding, changing and deleting comments |
| [Section](samples/OpenXml.Wrappers.Samples/Basic/SectionExample.cs) | Page size, orientation and margins; a landscape section |
| [Style](samples/OpenXml.Wrappers.Samples/Basic/StyleExample.cs) | Listing, adding, formatting and applying styles |
| [Field](samples/OpenXml.Wrappers.Samples/Basic/FieldExample.cs) | Adding fields and reading their code and result |
| [Header and footer](samples/OpenXml.Wrappers.Samples/Basic/HeaderFooterExample.cs) | Text, page numbers, a different first page |
| [Properties](samples/OpenXml.Wrappers.Samples/Basic/PropertiesExample.cs) | Built-in and custom document properties |
| [Character ranges](samples/OpenXml.Wrappers.Samples/Basic/CharactersExample.cs) | Format part of a sentence, words and sentences, find and narrow, bookmark and comment on a few characters |

**Intermediate: several objects together**

| Example | Shows |
| --- | --- |
| [Letter](samples/OpenXml.Wrappers.Samples/Intermediate/LetterExample.cs) | A mail merge from a template: placeholders, bookmarks, find and replace, properties |
| [Invoice](samples/OpenXml.Wrappers.Samples/Intermediate/InvoiceExample.cs) | Address and line-item tables, totals worked out from data, merged cells, shading |
| [Minutes](samples/OpenXml.Wrappers.Samples/Intermediate/MinutesExample.cs) | Headings, an action table, a bookmarked decision with a cross reference, comments, page-number footer |
| [Report](samples/OpenXml.Wrappers.Samples/Intermediate/ReportExample.cs) | Cover page, table of contents, "Page x of y", a landscape appendix |
| [Editing](samples/OpenXml.Wrappers.Samples/Intermediate/EditingExample.cs) | Open an existing document, replace text, fix a table, add comments, remove stale content |

**Advanced: medium-sized documents**

| Example | Shows |
| --- | --- |
| [Quarterly report](samples/OpenXml.Wrappers.Samples/Advanced/QuarterlyReportExample.cs) | Title page, contents, headings, data tables, comments, fields, headers and footers, a landscape section, properties |
| [Catalogue](samples/OpenXml.Wrappers.Samples/Advanced/CatalogueExample.cs) | Generated from a list of products: a section per category with its own header, nested tables, a bookmark per product, an index table |
| [Agreement](samples/OpenXml.Wrappers.Samples/Advanced/AgreementExample.cs) | A template with placeholders, numbered clauses (`SEQ`) and cross references (`REF`), merged-cell schedule, signature block, reviewer comments |

## Differences from the Word object model

The API follows the shape of the Word object model, but it is C# over Open XML rather than Word itself, so some things differ:

- **Indexes are 0-based.** `Paragraphs[0]` is the first paragraph, where Word's `Paragraphs(1)` is. This applies to every collection, to `Index`, `RowIndex` and `ColumnIndex`, and to `Table.Cell(row, column)`. There is no `Item()` method; use the indexer. An index out of range throws `ArgumentOutOfRangeException`. `Bookmarks["name"]` returns `null` when the bookmark does not exist.
- **Collections are `IEnumerable<T>`**, so `foreach` and LINQ work without casts. Enumeration works on a snapshot taken when it starts, so it is safe to delete items while enumerating. `Count` and the indexers are live: they look at the document each time.
- **A range is made of elements or of a span of characters.** `Range.Start` and `Range.End` are character positions as in Word (see [Character ranges](#character-ranges)); the Open XML elements that bound a range of elements are `StartElement` and `EndElement`. `document.Content` is the whole document, because `document.Range(start, end)` is a span of it.
- **Wrappers compare by the item they wrap.** Two wrappers of the same paragraph, table, bookmark and so on are `==` and have the same hash code, even though each property access creates a new wrapper.
- **Units are mostly the raw Open XML values**, which means the widths and spacing of the older members (`ParagraphWrapper.SpaceBefore`, `CellWidth`, column and table widths) are in twentieths of a point (twips), not points. `Font` and `ParagraphFormat` use points as Word does.
- **`Range.Text` uses Word's markers**: `\r` ends a paragraph, `\a` ends a table cell, `\r\a` ends a table row, `\v` is a line break and `\f` a page break. Reading text is supported for every kind of range. Setting it is supported for table cells, paragraphs, bookmarks, headers and footers, the whole document and spans of characters; rows and tables throw `InvalidOperationException`.
- **Not everything is implemented.** Features that are not supported throw `NotSupportedException`, `NotImplementedException` or `InvalidOperationException` rather than doing nothing, for example adding a comment to a range other than a bookmark, paragraph or table cell.

## Editing a range

A `RangeWrapper` can be edited much as a Word `Range` can:

```csharp
RangeWrapper range = document.Bookmarks["Summary"].Range;

range.InsertBefore("Note: ");
range.InsertAfter(" (draft)");

range.Font.Bold = true;
range.Font.Size = 11;               // points
range.Font.Color = "C00000";        // RGB hex
range.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
range.ParagraphFormat.SpaceAfter = 6;   // points
range.Style = "Heading 2";

// Find and replace, in a range, a table or the whole document.
document.Content.Find.Execute("colour", matchCase: true, replaceWith: "color", replace: ReplaceMode.All);

document.Paragraphs[3].Range.Delete();
```

| Member | Notes |
| --- | --- |
| `InsertBefore`, `InsertAfter` | Inside the markers of a bookmark or comment range. For a paragraph, table cell or the document the text goes in the first or last paragraph, before the paragraph mark (Word puts the text of `InsertAfter` on a whole paragraph in the next paragraph). Not supported for tables and rows. |
| `Delete` | Removes paragraphs, rows and tables; empties table cells and the document (a cell and the body must keep a paragraph); removes the runs of a bookmark or comment range. Bookmarks and comments anchored in deleted content are removed with it. |
| `Find` | `Execute` returns whether the text was found and narrows the range to the match. A replacement keeps the formatting of the first character it replaces. Matches can span runs and paragraphs. See [Character ranges](#character-ranges). |
| `Font` | `Bold`, `Italic`, `Underline`, `Strikethrough`, `Size`, `Name` and `Color`. Applies to the characters of the range, splitting runs where it starts and ends in the middle of one. |
| `ParagraphFormat` | `Alignment`, indents, `SpaceBefore`, `SpaceAfter`, `LineSpacing`, `LineSpacingRule`, `KeepWithNext`, `KeepTogether`, `PageBreakBefore` and `WidowControl`, in points. Also available as `Paragraph.Format`. |
| `Style` | A paragraph or character style by name. `Table.Style` sets a table style. |
| `InsertParagraphAfter`, `InsertParagraphBefore` | Add an empty paragraph next to a paragraph, table, bookmark, comment, table cell, the document or a header/footer, and return it. |

`Font` and `ParagraphFormat` properties are `null` when the range has mixed values, and setting one to `null` removes the direct formatting. They only see direct formatting, not values inherited from styles.

## Character ranges

A range is either made of whole elements (a paragraph, table, bookmark and so on) or, as in Word, a span of characters. `range.Start` and `range.End` are integer positions counted from the start of the text, which is the main text of the document or one header, footer or comment (a range in a header counts from the start of that header). A position counts each character, each tab and line break, each end of paragraph (which is also the end of a table cell for the last paragraph of a cell) and each end of table row. Pictures and field codes are not counted.

```csharp
RangeWrapper range = document.Range(6, 11);       // 'world' in 'Hello world'
range.Font.Bold = true;                            // splits the run; nothing else is touched
range.Text = "there";                              // replaces the text, keeping the formatting
range.Collapse(CollapseDirection.End);             // an insertion point
range.InsertAfter(", friend");

RangeWrapper paragraph = document.Paragraphs[2].Range;
paragraph.Words[3].Font.Italic = true;             // Characters, Words and Sentences are ranges
paragraph.MoveEnd(TextUnit.Word, 2);               // MoveStart, MoveEnd, Move, Expand and SetRange
```

- `document.Content` is the whole document and `document.Range(start, end)` a span of it. `SetRange`, `Collapse`, `MoveStart`, `MoveEnd`, `Move` and `Expand` turn a range of elements into a span in place, as in Word, and `Duplicate()` copies one. Indexes of the `Characters`, `Words` and `Sentences` collections are 0-based. A word includes the spaces after it, a sentence the spaces after its full stop, and a word or sentence only partly inside the range is included whole.
- A span is made of whole runs when it is used: reading or changing the formatting of part of a run, or deleting part of one, first splits the run in two. That does not change how the document looks, and it means `Font`, `Style` and `Find` work on exactly the characters of the range.
- **Find** narrows the range to the match when it finds the text, so a loop works as in VBA: `while (range.Find.Execute("text")) { range.Font.Bold = true; }` carries on from the end of each match up to the end of the range it started in (a collapsed range is searched to the end of the text). A match can span paragraphs and the text can use `^p` (paragraph mark), `^t` (tab), `^l` (line break), `^m` (page break) and `^^` (a caret). After a replace all the range keeps its start and covers the same region with the replacements in it.
- `InsertBefore` and `InsertAfter` on a span insert at its start or end and grow the range to include the text, and a carriage return in the text starts a new paragraph. `Delete` removes the characters: a range that takes in paragraph marks joins the paragraphs, whole paragraphs, rows and tables go, the end of a table cell stays, and a field that is only partly in the range is removed whole. `Bookmarks.Add`, `Comments.Add`, `Fields.Add`, `Style`, `ParagraphFormat` (which applies to every paragraph the range touches) and `InsertParagraphAfter` all work on a span, including a collapsed one.
- `Text` of a table row or table cannot be set.
- **A range follows its text.** The ends of a span are tied to the text itself, not to numbers, so when text is inserted or removed before, inside or after a range the range stays on the same text, however the change was made: through another range, through the other wrappers or directly with the Open XML SDK. Text inserted inside a range is part of it. If the text a range covered is removed through the library the range collapses to where it was; if it is removed behind the library's back (by deleting a paragraph element with the SDK, say) the range keeps its last known position.
- **Cost.** Reading `Start` and `End` as numbers looks through the whole story, so keep the ranges (or their words) rather than the numbers. Everything else is cheap for a range inside one paragraph: formatting, replacing, inserting, deleting, `Words` and `Sentences`, bookmarks and comments only look at that paragraph, and a `Find` that does not include a paragraph mark carries on from where it stopped, so looping over the matches of a document of thousands of paragraphs takes about a second. Changing the text across paragraphs, moving by units (`MoveStart`, `MoveEnd`, `Move`) and `SetRange` look through the whole story each time.
## Table formatting

These use points like `Font` and `ParagraphFormat`: `Row.Height`, `Row.HeightRule` (`Auto`, `AtLeast`, `Exactly`), `Row.HeadingFormat`, `Row.AllowBreakAcrossPages`, `Cell.VerticalAlignment`, `Cell.BackgroundColor` (six hex digits), `Table.Alignment`, `Table.PreferredWidth` (null unless a fixed width) and `Column.Index`. `Column.Delete` removes the grid column and the cell at its position in every row; a cell that spans the column loses one column of its span instead, and the start of a vertical merge passes to the next row. It throws `NotSupportedException` only for the older `hMerge` kind of merge.

`Table.Borders` and `Cell.Borders` work like Word's: `Borders[BorderPosition.Top]` is a border with `LineStyle`, `LineWidth` (points, 0.25 to 12) and `Color` (hex or "auto"), and `SetOutside` / `SetInside` set several at once. Positions are `Top`, `Left`, `Bottom`, `Right`, `Horizontal` and `Vertical` (the lines between rows and columns), plus the two diagonals for cells only. Setting `LineStyle` to `None` writes an explicit "no border"; `Clear()` removes the direct setting so the table style applies. A border with no direct setting reads from the table (a cell takes the table's outside borders on the edges of the table and its inside borders elsewhere) and from the table style and the styles it is based on; the conditional formatting of a style for the first and last row and column is read when the table's look switches it on (banded rows are not). `Row.Borders` works the same way but is stored on the cells of the row: `Top` and `Bottom` go on every cell, `Left` on the first, `Right` on the last and `Vertical` between the cells; a row has no `Horizontal` or diagonal border. A row border whose cells disagree reads `LineStyle` as `Other`, `LineWidth` as 0 and `Color` as null. `Row.SetTopBorder` and `SetBottomBorder` still work as before.

## Sections, styles and document properties

- `document.Sections` lists the sections in order (a document always has at least one). `Section.PageSetup` has `PageWidth`, `PageHeight`, `Orientation` (changing it swaps the size, as Word does), `TopMargin`, `BottomMargin`, `LeftMargin`, `RightMargin`, `Gutter`, `HeaderDistance`, `FooterDistance`, `SectionStart` and `DifferentFirstPageHeaderFooter`, all in points. A section without a setting reads Word's defaults (Letter, 72 point margins) and writes the setting when you set it.
- `document.Styles` lists the styles (`Name`, `StyleId`, `Type`, `BuiltIn`, `BaseStyle`, `NextParagraphStyle`), looks them up by name or id, and has `Add(name, type)` and `Delete(style)`. Built-in and default styles cannot be deleted; deleting a style removes the references to it. `Style.Font` and `Style.ParagraphFormat` work like the range versions but on the formatting set in the style itself (nothing inherited from the style it is based on is read); setting a value to null removes it.
- `document.BuiltInDocumentProperties` has `Title`, `Subject`, `Author`, `Keywords`, `Comments`, `Category`, `LastAuthor`, `CreationDate`, `LastSaveTime`, `Company` and `Manager`. `document.CustomDocumentProperties` holds the custom properties: `Add(name, type, value)` (types `Boolean`, `Date` in UTC, `FloatingPoint`, `Number`, `Text`), lookup by name (ignoring case) or position, `Exists`, and `Value`, `Type`, `Name` and `Delete()` on each. Looking never creates the part, and the link-to-content option of Word is not supported.

## Table structure

- `Tables.Add(range, numRows, numColumns)` inserts an empty table with single 0.5 point borders immediately **before** the range (a paragraph, table, cell, bookmark, comment or the whole document) and fills the width between the margins or of the cell. Unlike Word it does not replace the text of the range.
- `Rows.Add(beforeRow)` inserts an empty row with the cells and formatting of `beforeRow` before it; with `null` it appends.
- `Cell.ColumnIndex` and `Table.Cell(row, column)` count the cells of a row, as Word does, so in a row with merged cells they are not grid columns. `Cell.Merge(otherCell)` merges the rectangle between two cells (in either order) using `gridSpan` and vertical merges, and moves the text of the merged cells into the first one.
- `Cell.Split(numRows, numColumns)` divides a cell. Splitting into columns divides the grid column, so cells in other rows that span it get wider spans; splitting into rows adds rows and vertically merges the rest of the row.

A merge absorbs vertical merges that lie inside the rectangle, and cuts the ones that stick out of it (what is left above or below stays merged, or becomes an ordinary cell when it is a single row). A vertically merged cell can be split into columns (every row of the merge is split) or into rows: into as many rows as it has it is unmerged, into more rows are added below with the other cells running through them, and into fewer the rows are shared out. A cell next to a vertically merged cell can be split into rows (the merge gets longer). Merging and splitting throw `NotSupportedException` rather than guess in the remaining cases: splitting a merged cell into rows and columns at once, splitting a cell that spans columns into a number that does not divide the span, and tables with the older `hMerge` horizontal merge, which Word no longer writes.

## Bookmarks and comments

`Bookmarks.Add(name, range)` bookmarks a paragraph, table, table cell, bookmark or comment range (an existing bookmark of the same name is moved), and `Bookmarks.Exists(name)` and `Bookmarks["name"]` ignore case. Names follow Word's rules: a letter first, then letters, digits and underscores, at most 40 characters. `Comment.Text` can be set (it replaces the body of the comment), `Comment.Date` is UTC, and `Comment.Delete()` removes the comment together with its range markers and reference.

## Headers and footers

`section.Headers[HeaderFooterIndex.Primary]` (also `FirstPage` and `EvenPages`) and `section.Footers[...]` give a `HeaderFooter` with `Range`, `Exists`, `LinkToPrevious`, `IsHeader`, `Index` and `Parent`. They are indexed by `HeaderFooterIndex` as in Word, not by position, and the collections always hold the three kinds.

```csharp
document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Quarterly report";
FieldWrapperCollection fields = document.Sections[0].Footers[HeaderFooterIndex.Primary].Range.Fields;
```

Reading `Range` creates an empty header or footer if the section and the sections before it have none; a section without its own shows (and edits) that of the nearest earlier section, as `LinkToPrevious` says. `LinkToPrevious = false` gives the section a copy of what it was showing; `true` removes its own. `Exists` is true only for a header or footer the section has of its own. A header range supports `Text`, `InsertBefore`/`InsertAfter`, `Find`, `Font`, `Fields` and the other paragraph-level features; setting `Text` replaces the content with one plain paragraph per line.

Creating a first-page header does not switch on `DifferentFirstPageHeaderFooter`, and even-page headers only show when the document's odd/even setting is on: `section.PageSetup.OddAndEvenPagesHeaderFooter = true` (a setting of the whole document).

## Fields

`document.Fields` and `range.Fields` list the fields in document order: complex fields (begin, separate and end characters around a code) and simple fields, including fields nested in the result of another. A field has `Code` (read and set, without the surrounding spaces), `Type` (a `FieldType`, `Other` for kinds not in the enum), `Result` (the stored text, read and set), `Range`, `Parent` and `Delete()`. `Fields.Add(range, FieldType.Page, "\* MERGEFORMAT")` or `Fields.Add(range, "MERGEFIELD Name")` adds a complex field at the end of the range, by the rules of `InsertAfter`, with an empty result and the dirty flag so Word updates it when the document opens.

Fields are never calculated by the library, so `Result` is whatever the document stores. Setting a result needs the whole result to be inside one paragraph; the result takes the formatting of the first run it replaces.

## Navigation

`Paragraph.Next()` / `Previous()` walk the paragraphs of the document in order, through table cells, and return `null` at either end. `Row.Next()` / `Previous()` / `IsFirst` / `IsLast` and `Cell.Next()` / `Previous()` work within a table (a cell's neighbour wraps to the adjacent row). `Range.Duplicate()` returns a new range over the same text; `Collapse`, `SetRange`, `MoveStart`, `MoveEnd`, `Move` and `Expand` move through characters, words, sentences and paragraphs (see [Character ranges](#character-ranges)).

## Parent

Every wrapper and collection has a `Parent` that answers "what contains me": a cell's is its row, a row's and a column's its table, a table's the table it is nested in (or null), and the table collections' the table. A paragraph's, bookmark's and comment's `Parent` is a `RangeWrapper` over the element that contains it (the table cell or body, or the paragraph), and the paragraph, bookmark and comment collections return the `RangeWrapper` they were taken from.

## Build and test

```bash
dotnet build E2e4.OpenXml.Wrappers.slnx
dotnet test E2e4.OpenXml.Wrappers.slnx
dotnet format E2e4.OpenXml.Wrappers.slnx --verify-no-changes
```

## Layout

| Path | Contents |
| --- | --- |
| `OpenXml.Wrappers/Word` | The wrappers, their collections and the `OpenXmlExtensions` tree-walking helpers |
| `OpenXml.Wrappers/Helpers` | Unit conversion helpers |
| `OpenXml.Wrappers.Test` | xUnit tests and the fixture documents in `TestFiles` |
| `samples/OpenXml.Wrappers.Samples` | The 21 runnable examples |
| `assets` | The icon (`icon.svg`, and `icon.png` used for the NuGet package) |

See [CLAUDE.md](CLAUDE.md) for architecture notes and conventions.

## License

[MIT](LICENSE). Copyright © 2026 GitHub: th3snark (E2e4).
