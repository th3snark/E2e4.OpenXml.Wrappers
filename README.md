# E2e4.OpenXml.Wrappers

A .NET library that wraps the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) WordprocessingML types in a Word-object-model-like API: `Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments` and friends.

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

`Create` makes a new document (from a path or a stream) with an empty body, and `SaveAs` writes a copy to a new file. Editable documents are also saved when disposed or closed.

## Differences from the Word object model

The API follows the shape of the Word object model, but it is C# over Open XML rather than Word itself, so some things differ:

- **Indexes are 0-based.** `Paragraphs[0]` is the first paragraph, where Word's `Paragraphs(1)` is. This applies to every collection, to `Index`, `RowIndex` and `ColumnIndex`, and to `Table.Cell(row, column)`. There is no `Item()` method; use the indexer. An index out of range throws `ArgumentOutOfRangeException`. `Bookmarks["name"]` returns `null` when the bookmark does not exist.
- **Collections are `IEnumerable<T>`**, so `foreach` and LINQ work without casts. Enumeration works on a snapshot taken when it starts, so it is safe to delete items while enumerating. `Count` and the indexers are live: they look at the document each time.
- **Ranges are made of elements, not character positions.** `Range.Start` and `Range.End` are the Open XML elements that bound the range.
- **Wrappers compare by the item they wrap.** Two wrappers of the same paragraph, table, bookmark and so on are `==` and have the same hash code, even though each property access creates a new wrapper.
- **Units are mostly the raw Open XML values**, which means the widths and spacing of the older members (`ParagraphWrapper.SpaceBefore`, `CellWidth`, column and table widths) are in twentieths of a point (twips), not points. `Font` and `ParagraphFormat` use points as Word does.
- **`Range.Text` uses Word's markers**: `\r` ends a paragraph, `\a` ends a table cell, `\r\a` ends a table row, `\v` is a line break and `\f` a page break. Reading text is supported for paragraphs, table cells, rows and tables, bookmarks and comments (not for a whole document range). Setting it is supported for table cells, paragraphs and bookmarks; other types throw `InvalidOperationException`.
- **Some properties are set-only**, for example `KeepNext` on paragraphs, rows and cells.
- **Not everything is implemented.** Features that are not supported throw `NotSupportedException`, `NotImplementedException` or `InvalidOperationException` rather than doing nothing, for example `InsertSectionBreak` for any type but `"nextPage"`, or adding a comment to a range other than a bookmark, paragraph or table cell.

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
document.Range.Find.Execute("colour", matchCase: true, replaceWith: "color", replace: ReplaceMode.All);

document.Paragraphs[3].Range.Delete();
```

| Member | Notes |
| --- | --- |
| `InsertBefore`, `InsertAfter` | Inside the markers of a bookmark or comment range. For a paragraph, table cell or the document the text goes in the first or last paragraph, before the paragraph mark (Word puts the text of `InsertAfter` on a whole paragraph in the next paragraph). Not supported for tables and rows. |
| `Delete` | Removes paragraphs, rows and tables; empties table cells and the document (a cell and the body must keep a paragraph); removes the runs of a bookmark or comment range. Bookmarks and comments anchored in deleted content are removed with it. |
| `Find` | `Execute` returns whether the text was found. It never changes the range or the document, because a range is made of whole elements and cannot be narrowed to the match. Replace keeps the formatting of the first run of the match. Matches can span runs but not paragraphs. |
| `Font` | `Bold`, `Italic`, `Underline`, `Strikethrough`, `Size`, `Name` and `Color`. Applies to whole runs, so it cannot format part of a run. |
| `ParagraphFormat` | `Alignment`, indents, `SpaceBefore`, `SpaceAfter`, `LineSpacing`, `LineSpacingRule`, `KeepWithNext`, `KeepTogether`, `PageBreakBefore` and `WidowControl`, in points. Also available as `Paragraph.Format`. |
| `Style` | A paragraph or character style by name. `Table.Style` sets a table style. |

`Font` and `ParagraphFormat` properties are `null` when the range has mixed values, and setting one to `null` removes the direct formatting. They only see direct formatting, not values inherited from styles.

## Table formatting

These use points like `Font` and `ParagraphFormat`: `Row.Height`, `Row.HeightRule` (`Auto`, `AtLeast`, `Exactly`), `Row.HeadingFormat`, `Row.AllowBreakAcrossPages`, `Cell.VerticalAlignment`, `Cell.BackgroundColor` (six hex digits), `Table.Alignment`, `Table.PreferredWidth` (null unless a fixed width) and `Column.Index`. `Column.Delete` removes the grid column and the cell at its position in every row; it throws `NotSupportedException` for tables with horizontally merged cells.

`Table.Borders` and `Cell.Borders` work like Word's: `Borders[BorderPosition.Top]` is a border with `LineStyle`, `LineWidth` (points, 0.25 to 12) and `Color` (hex or "auto"), and `SetOutside` / `SetInside` set several at once. Positions are `Top`, `Left`, `Bottom`, `Right`, `Horizontal` and `Vertical` (the lines between rows and columns), plus the two diagonals for cells only. Setting `LineStyle` to `None` writes an explicit "no border"; `Clear()` removes the direct setting so the table style applies. Only direct formatting is read, so a border that comes from a table style reads as `None`. `Row.SetTopBorder` and `SetBottomBorder` still work as before.

## Sections, styles and document properties

- `document.Sections` lists the sections in order (a document always has at least one). `Section.PageSetup` has `PageWidth`, `PageHeight`, `Orientation` (changing it swaps the size, as Word does), `TopMargin`, `BottomMargin`, `LeftMargin`, `RightMargin`, `Gutter`, `HeaderDistance`, `FooterDistance`, `SectionStart` and `DifferentFirstPageHeaderFooter`, all in points. A section without a setting reads Word's defaults (Letter, 72 point margins) and writes the setting when you set it.
- `document.Styles` lists the styles (`Name`, `StyleId`, `Type`, `BuiltIn`, `BaseStyle`, `NextParagraphStyle`), looks them up by name or id, and has `Add(name, type)` and `Delete(style)`. Built-in and default styles cannot be deleted; deleting a style removes the references to it. Style formatting is not exposed.
- `document.BuiltInDocumentProperties` has `Title`, `Subject`, `Author`, `Keywords`, `Comments`, `Category`, `LastAuthor`, `CreationDate`, `LastSaveTime`, `Company` and `Manager`. Custom properties are not implemented.

## Table structure

- `Tables.Add(range, numRows, numColumns)` inserts an empty table with single 0.5 point borders immediately **before** the range (a paragraph, table, cell, bookmark, comment or the whole document) and fills the width between the margins or of the cell. Unlike Word it does not replace the text of the range.
- `Rows.Add(beforeRow)` inserts an empty row with the cells and formatting of `beforeRow` before it; with `null` it appends.
- `Cell.Merge(otherCell)` merges the rectangle between two cells (in either order) using `gridSpan` and vertical merges, and moves the text of the merged cells into the first one.
- `Cell.Split(numRows, numColumns)` divides a cell. Splitting into columns divides the grid column, so cells in other rows that span it get wider spans; splitting into rows adds rows and vertically merges the rest of the row.

Merging and splitting throw `NotSupportedException` when they would cut through an existing merge, rather than guess: merging cells that are already vertically merged, splitting a vertically merged cell, splitting a cell with a vertically merged neighbour into rows, or splitting a cell that spans columns into a number that does not divide the span.

## Bookmarks and comments

`Bookmarks.Add(name, range)` bookmarks a paragraph, table cell, bookmark or comment range (an existing bookmark of the same name is moved), and `Bookmarks.Exists(name)` and `Bookmarks["name"]` ignore case. Names follow Word's rules: a letter first, then letters, digits and underscores, at most 40 characters. `Comment.Text` can be set (it replaces the body of the comment), `Comment.Date` is UTC, and `Comment.Delete()` removes the comment together with its range markers and reference.

## Headers and footers

`section.Headers[HeaderFooterIndex.Primary]` (also `FirstPage` and `EvenPages`) and `section.Footers[...]` give a `HeaderFooter` with `Range`, `Exists`, `LinkToPrevious`, `IsHeader`, `Index` and `Parent`. They are indexed by `HeaderFooterIndex` as in Word, not by position, and the collections always hold the three kinds.

```csharp
document.Sections[0].Headers[HeaderFooterIndex.Primary].Range.Text = "Quarterly report";
FieldWrapperCollection fields = document.Sections[0].Footers[HeaderFooterIndex.Primary].Range.Fields;
```

Reading `Range` creates an empty header or footer if the section and the sections before it have none; a section without its own shows (and edits) that of the nearest earlier section, as `LinkToPrevious` says. `LinkToPrevious = false` gives the section a copy of what it was showing; `true` removes its own. `Exists` is true only for a header or footer the section has of its own. A header range supports `Text`, `InsertBefore`/`InsertAfter`, `Find`, `Font`, `Fields` and the other paragraph-level features; setting `Text` replaces the content with one plain paragraph per line.

Creating a first-page header does not switch on `DifferentFirstPageHeaderFooter`, and even-page headers only show when the document's odd/even setting is on, which the library does not expose.

## Fields

`document.Fields` and `range.Fields` list the fields in document order: complex fields (begin, separate and end characters around a code) and simple fields, including fields nested in the result of another. A field has `Code` (read and set, without the surrounding spaces), `Type` (a `FieldType`, `Other` for kinds not in the enum), `Result` (the stored text, read and set), `Range`, `Parent` and `Delete()`. `Fields.Add(range, FieldType.Page, "\* MERGEFORMAT")` or `Fields.Add(range, "MERGEFIELD Name")` adds a complex field at the end of the range, by the rules of `InsertAfter`, with an empty result and the dirty flag so Word updates it when the document opens.

Fields are never calculated by the library, so `Result` is whatever the document stores. Setting a result needs the whole result to be inside one paragraph; the result takes the formatting of the first run it replaces.

## Navigation

`Paragraph.Next()` / `Previous()` walk the paragraphs of the document in order, through table cells, and return `null` at either end. `Row.Next()` / `Previous()` / `IsFirst` / `IsLast` and `Cell.Next()` / `Previous()` work within a table (a cell's neighbour wraps to the adjacent row). `Range.Duplicate()` returns a new range over the same elements. Integer `Start`/`End`, `Collapse` and `SetRange` are not provided: a range is made of whole elements, so there are no character positions to move.

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

See [CLAUDE.md](CLAUDE.md) for architecture notes and conventions.
