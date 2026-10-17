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
- **Units are the raw Open XML values**, which means widths and spacing are in twentieths of a point (twips), not points.
- **`Range.Text` uses Word's markers**: `\r` ends a paragraph, `\a` ends a table cell, `\r\a` ends a table row, `\v` is a line break and `\f` a page break. Reading text is supported for paragraphs, table cells, rows and tables, bookmarks and comments (not for a whole document range). Setting it is supported for table cells, paragraphs and bookmarks; other types throw `InvalidOperationException`.
- **Some properties are set-only**, for example `KeepNext` on paragraphs, rows and cells.
- **Not everything is implemented.** Features that are not supported throw `NotSupportedException`, `NotImplementedException` or `InvalidOperationException` rather than doing nothing, for example `InsertSectionBreak` for any type but `"nextPage"`, or adding a comment to a range other than a bookmark, paragraph or table cell.

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
