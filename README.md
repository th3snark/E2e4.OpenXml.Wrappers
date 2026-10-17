# E2e4.OpenXml.Wrappers

A .NET library that wraps the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) WordprocessingML types in a Word-object-model-like API: `Range`, `Paragraphs`, `Tables`, `Bookmarks`, `Comments` and friends.

## Requirements

- .NET 10 SDK
- `DocumentFormat.OpenXml` 3.x (restored automatically)

## Usage

```csharp
using OpenXml.Wrappers.Word;

WordprocessingDocumentWrapper document = WordprocessingDocumentWrapper.Open("report.docx", isEditable: true);

foreach (ParagraphWrapper paragraph in document.Paragraphs)
{
    Console.WriteLine(paragraph.Text);
}

document.Close();
```

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
