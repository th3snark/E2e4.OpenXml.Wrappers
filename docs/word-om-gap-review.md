# Word object model gap review


> **Status (for 1.0.0).** This review was written before the work it proposes. Done since: C1 (colours: `WordColor`, `ColorHelper.RGB`, `NamedColor`, the `ColorIndex` members), C2 (the `StyleObject` members; the `Style` properties stay names by decision), C3 (`LineWidth`, `BorderLineWidth`), C5 (`Document.Range()`), C7 (`TrackRevisions = true` throws until revisions are recorded), `Range.Delete(unit, count)`, nullable `Paragraph` formatting, `Field.Code`/`Result` with `Range` and `Font`, and the clean-up in section 4. Everything else below is still open and is additive, so it can ship in 1.x releases. The README, not this file, describes what the library does today.
Scope: the public API of `OpenXml.Wrappers.Word` (branch `api-word-fidelity`) against the Word VBA object model, looking first at the objects that already exist, then at the limitations, and last at the objects that are missing.

Method and caveats: this comes from reading the public signatures of every file in `OpenXml.Wrappers\Word`, plus `README.md`, `CHANGELOG.md` and `CLAUDE.md`. I did not run the code or check behaviour, and I compared against the Word object model from memory, not against the current Microsoft reference. Check each item against the docs before building it. Some members I list as missing may sit in a file I only skimmed.

## 1. Summary

The existing objects are already well covered. `Range` navigation, `Find`, `Font`, `ParagraphFormat`, tables, fields, comments, bookmarks, sections and page setup mostly have Word's member names, units and semantics. The gaps fall into four groups:

1. **Cross-cutting conventions that macro code leans on.** These are colours, style objects, line-width constants, and `Selection`/`Application`/`ActiveDocument`. A VBA port fails on these before it reaches any missing feature.
2. **Missing members on objects that exist.** The largest cluster is `Range` itself, then `Cells`/`Rows`/`Columns` as collections, `Find` formatting, `Borders`/`Shading` on paragraphs and styles, and `Style`.
3. **Surface that is not Word's.** The `Paragraph`, `Table` and `Cell` types carry several helpers that Word does not have. Some leak Open XML types.
4. **Documented limitations.** Most need a layout engine and cannot be closed. Some are choices that could change.

## 2. Cross-cutting gaps (fix these first)

| # | Gap | Why it matters | Suggestion |
|---|-----|----------------|------------|
| C1 | **Colours are hex strings** (`Font.Color`, `Border.Color`, `Shading.*Color`, `Cell.BackgroundColor`). Word uses `WdColor` (a BGR `Long`) and `WdColorIndex`. | VBA like `.Font.Color = wdColorRed` or `RGB(255,0,0)` has no direct equivalent. The hex-string convention has to be learned. | Add a `WdColor` enum and a `ColorHelper.RGB(r,g,b)` returning the BGR int. Add `int` overloads, or a `Color` struct with implicit conversions from both, so `Font.Color = RGB(...)` ports. Add `Font.TextColor`, `Font.ColorIndex`, `Border.ColorIndex` and `Shading.*PatternColorIndex`. Keep the string members. |
| C2 | **Style references are strings** (`Range.Style`, `Paragraph.Style`, `Table.Style`, `Style.BaseStyle`, `NextParagraphStyle`). In Word the getter returns a `Style` object and the setter accepts a name, a `Style` or a `WdBuiltinStyle`. | Macros write `Selection.Style = ActiveDocument.Styles("Heading 1")`, or read `.Style.NameLocal`. | Return `StyleWrapper` from getters (or add a typed `StyleObject` and keep the string). Add setter overloads for `StyleWrapper` and `BuiltInStyle`, and `Style.NameLocal`. |
| C3 | **Line widths are doubles in points.** Word's `LineWidth` is the `WdLineWidth` enum (`wdLineWidth050pt` and so on). | Ported constants do not compile. | Add a `LineWidth` enum with Word's values. Accept it, and map it to points internally. |
| C4 | **There is no `Application`, `Documents`, `ActiveDocument` or `Selection`.** Nearly every macro starts at one of them. | The "familiar to VBA authors" goal breaks in the first line of most macros. | See section 6. A light `Application`/`Documents` facade is cheap. `Selection` needs a design decision. |
| C5 | **`Document.Range(start, end)` needs both arguments.** In Word both are optional, and `Range()` is the whole story. | Small, but `doc.Range()` is common. | Make both optional. |
| C6 | **The API uses `Wrapper` suffixes** (`RangeWrapper`, `CommentWrapper`). Word's names are `Range`, `Comment`. | Only matters for readability and ported snippets. | Document `using Range = OpenXml.Wrappers.Word.RangeWrapper;` aliases in the README. Low cost, no code change. |
| C7 | **Tracked changes are not honoured (decided: `TrackRevisions = true` throws until they are).** `TrackRevisions` only sets the setting. Edits made through the library are not recorded as revisions. | With the flag on in Word, the same macro produces revisions. Here it silently produces plain edits. | At minimum, document that. Better: when `TrackRevisions` is true, write `w:ins`/`w:del` for text insert and delete. See the Revisions entry in section 5. |

## 3. Gaps on objects that exist

Ordered by how often macros use them. "Cheap" means the Open XML mapping is direct. "Moderate" means it needs the story-map or span machinery. "Hard" needs new subsystems.

### 3.1 Range (highest value)

Present: `Text`, `Start`, `End`, `SetRange`, `Collapse`, `Move*`, `Expand`, `Characters`/`Words`/`Sentences`, `Paragraphs`, `Tables`, `Cells`, `Rows`, `Bookmarks`, `Comments`, `Fields`, `Find`, `Font`, `ParagraphFormat`, `Style`, `InsertBefore`/`After`, `InsertParagraph*`, `InsertBreak`, `Delete`, `Duplicate`.

| Missing member | Notes | Effort |
|---|---|---|
| `Range.Delete(Unit, Count)` and an `int` return | Word's `Delete` returns the number of units deleted. Ours returns void. | Cheap |
| `Range.Next(Unit, Count)`, `Range.Previous(Unit, Count)` | The unit navigator already exists. | Cheap |
| `MoveStartUntil/While`, `MoveEndUntil/While` | Common in text-processing macros. | Cheap |
| `IsEqual`, `InRange`, `InStory`, `Range.Document`, `Range.Parent`, `StoryType`, `StoryLength` (present) | Comparison and identity. | Cheap |
| `Range.Sections`, `Range.Columns`, `Range.Sentences` (present) | Collections scoped to the range. | Cheap |
| `Range.ParagraphStyle`, `CharacterStyle`, `CharacterFormat` | Split style reads and writes. | Cheap |
| `Range.Information(WdInformation)` | The page and line items need layout. A layout-free subset is possible: `wdWithInTable`, `wdStartOfRangeRowNumber`/`ColumnNumber`, `wdEndOfRange*`, `wdMaximumNumberOfRows`/`Columns`, `wdAtEndOfRowMarker`, `wdHeaderFooterType`, `wdInFootnote` and `wdInEndnote` (when those exist), `wdInCommentPane`. | Moderate |
| `Range.Case` / `ChangeCase(WdCharacterCase)` | Plain text transform per run. | Cheap |
| `Range.Sort`, `Range.SortAscending`/`Descending`, and `Table.Sort` | The README lists `Table.Sort` as unimplemented. | Moderate |
| `Range.FormattedText` and `Range.Copy`/`Cut`/`Paste` | The macro idiom for moving content is `r2.FormattedText = r1.FormattedText`. There is no clipboard, so a `FormattedText` property that copies the elements is what is needed. It must preserve bookmarks, comments and fields inside the copy. | Hard, but high value |
| `Range.InsertFile`, `InsertXML`, `WordOpenXML`, `XML` | `InsertXML` and `WordOpenXML` are natural for an Open XML library. `InsertFile` reuses the `CopyAllPartsTo` machinery. | Moderate |
| `Range.InsertDateTime`, `InsertSymbol`, `InsertCrossReference`, `InsertCaption`, `InsertAlignmentTab`, `InsertAutoText` | `InsertDateTime`, `InsertSymbol` and `InsertCrossReference` map to fields, so they are cheap on top of `Fields.Add`. | Cheap to moderate |
| `Range.GoTo`, `GoToNext`, `GoToPrevious`, `Document.GoTo` | Support the layout-free targets: bookmark, section, table, comment, field, heading, equation, graphic (later). | Moderate |
| `Range.Borders`, `Range.Shading` | Present for table parts and pages only. See 3.4. | Moderate |
| `Range.Select`, `Range.PasteSpecial`, `Range.ShapeRange`, `Range.Editors`, `Range.Language*`, `Range.NoProofing`, `Range.HighlightColorIndex` | Highlight is the one that matters. See 3.2. | Cheap |
| `Range.Revisions`, `Hyperlinks`, `Footnotes`, `Endnotes`, `InlineShapes`, `ContentControls`, `ListFormat`, `ListParagraphs` | Missing objects. See section 5. | n/a |

### 3.2 Font

Present and good: effective-value reading, `Direct`, `Reset`, the toggles and the spacing group.

| Missing | Notes |
|---|---|
| `Highlight` / `Range.HighlightColorIndex` | Probably the most-used missing formatting property. Map to `w:highlight`. |
| `Color` as `WdColor`, `ColorIndex`, `TextColor` | See C1. |
| `NameAscii`, `NameOther`, `NameFarEast`, `NameBi`, `SizeBi`, `BoldBi`, `ItalicBi` | Complex-script and East Asian variants. Open XML has `rFonts` attributes and `bCs`/`iCs`/`szCs`. |
| `Borders`, `Shading` (character) | `w:bdr`, `w:shd` on run properties. |
| `Animation`, `Contrast`, `Fill`, `Glow`, `Line`, `Reflection`, `Shadow` (the object) and `Grow()`/`Shrink()`, `Duplicate` | `Grow`/`Shrink` and `Duplicate` are cheap. The effects objects are 2010 text effects stored in `w14:` extensions. Low priority. |
| `Font.StrikeThrough` / `DoubleStrikeThrough` | Present. |

### 3.3 Paragraph, Paragraphs and ParagraphFormat

| Missing | Notes |
|---|---|
| `Borders`, `Shading` on `Paragraph`, `Paragraphs`, `ParagraphFormat` | `pBdr` and `shd` in paragraph properties. |
| `Paragraph.OutlineLevel`, `Hyphenation` and the other format properties directly on `Paragraph` | `Paragraph` forwards only some of `ParagraphFormat`. In Word, `Paragraph` exposes the whole set. Forward everything, and make `Paragraphs` collection-level setters do the same. |
| `Indent()`, `Outdent()`, `OpenUp()`, `CloseUp()`, `Space1()`, `Space15()`, `Space2()`, `TabHangingIndent()`, `TabIndent()`, `Reset()`, `Paragraph.ID` (not needed) | Cheap. These are used heavily in macros recorded from the ribbon. |
| `LineUnitBefore`/`After`, `CharacterUnitLeftIndent` and the other `*Unit` members, `FarEast*`, `WordWrap`, `AddSpaceBetweenFarEast*`, `Alignment` for East Asian | Open XML has the attributes (`beforeLines`, `leftChars`). |
| `ParagraphFormat.Duplicate`, `ParagraphFormat.Borders`, `Shading` | |
| `Paragraphs.Add(Range)` | Ours takes no argument. Word inserts before the given range. |
| `Paragraphs.Borders/Shading/TabStops/Alignment ...` | Only `Format` and `Style` are collection-level. |
| `Paragraph.CloseUp` family and `Paragraphs.Count` | |

### 3.4 Borders and Shading

* `Borders` exists for tables, rows, columns, cells, sections and pages. It is missing for paragraphs, ranges and styles (`Range.Borders`, `Paragraph.Borders`, `Style.Borders`, `Font.Borders`).
* Missing members on `Borders`: `DistanceFromTop/Left/Bottom/Right`, `Shadow`, `AlwaysInFront`, `JoinBorders`, `DisplayGridlines`, `Item(wdBorderHorizontal ...)` coverage for paragraphs, `ApplyPageBordersToAllSections`, `EnableFirstPageInSection`. Missing members on `Border`: `ColorIndex`, `ArtStyle`/`ArtWidth`, `LineWidth` as an enum (C3), and `Color` as `WdColor` (C1).
* `Shading` is only reachable through tables. Add `Range.Shading`, `Paragraph.Shading`, `Style.Shading`, `Font.Shading`, and the `BackgroundPatternColorIndex`/`ForegroundPatternColorIndex` pair. `Texture` already exists.

### 3.5 Styles

Present: `Name`, `StyleId`, `Type`, `BuiltIn`, `BaseStyle`, `NextParagraphStyle`, `Font`, `ParagraphFormat`, `Delete`, `Styles.Add`, built-in catalogue.

Missing: `NameLocal`, `Description`, `InUse`, `Locked`, `Priority`, `QuickStyle`, `Visibility`, `UnhideWhenUsed` (the catalogue has it internally), `LinkStyle`/`Linked`, `Borders`, `Shading`, `Table` (the `TableStyle` object with conditional formatting), `ListTemplate`, `ListLevelNumber`, `NoSpaceBetweenParagraphsOfSameStyle`, `AutomaticallyUpdate`, `LanguageID`, `Style.Document`, `Styles.Count` of only the used or defined styles (check that it matches Word's rules for latent styles), `Style.Frame`. Also `Styles.Add(name, type)` should return an object that supports all of these. Note that built-in coverage should be checked against the full `WdBuiltinStyle` list. For built-in styles that are not in the catalogue the behaviour is not clear from the API.

### 3.6 Tables, Rows, Columns, Cells

This is the best-covered area, and the collection objects are where it is thinnest.

* **`Cells` is a bare collection.** Word's `Cells` has `Add(BeforeCell)`, `Merge`, `Split`, `Delete`, `Width`, `Height`, `HeightRule`, `Borders`, `Shading`, `VerticalAlignment`, `SetWidth`, `SetHeight`, `DistributeHeight`, `DistributeWidth`, `AutoFit`, `Range` and `Parent`. Ours has an indexer and `Count`. This also applies to `Range.Cells`, which is the common way to format a block of cells.
* **`Rows`**: `DistributeHeight` is missing (also listed in the README as unimplemented), `AllowOverlap`, `WrapAroundText`, `DistanceTop/Bottom/Left/Right`, `HorizontalPosition`, `VerticalPosition`, `RelativeHorizontalPosition`, `RelativeVerticalPosition`, `Borders`, `Range`, `SetLeftIndent`, `SetHeight`, `ConvertToText`.
* **`Row`**: `LeftIndent` (it exists on `Rows` only), `NestingLevel`, `Select`, `Range.Information` rows, `Row.Cells.Add`.
* **`Columns`**: `Width`, `PreferredWidth`, `PreferredWidthType`, `SetWidth(width, RulerStyle)`, `AutoFit`, `Borders`, `Select`, `Delete`, `Column.Next`/`Previous`, `Column.IsFirst` (present).
* **`Cell`**: `Formula`, `Orientation`/`TextOrientation`, `AutoFit`, `SetWidth`, `SetHeight`, `Select`, `Cell.NestingLevel`, `Cell.ID`, `Delete(ShiftCells.Up)` (throws today), and deleting a cell that is part of a vertical merge (throws today). The README says `Cell.Delete` is not implemented, but it is, with those limits. Update the README.
* **`Table`**: `Sort`, `ConvertToText`, `Range.ConvertToTable`, `AutoFormat`, `TableDirection`, `AllowPageBreaks`, `Rows.Wrap*` (above), `Table.Select`, `Table.ID`, `Table.Style` as an object (C2), `ApplyStyle*` (present), `Table.Title`/`Descr` (present).
* **Legacy horizontal merges** (`hMerge`) throw in merge, split and column delete. Files from older Word versions can contain them. A normalising pass that converts `hMerge` to `gridSpan` at open would remove most of these `NotSupportedException`s.

### 3.7 Find / Replacement

* `Find.Font`, `Find.ParagraphFormat`, `Find.Style`, `Find.Highlight`, and the same on `Replacement`, are missing. `Format = true` and `MatchSoundsLike`/`MatchAllWordForms` throw. `Replacement.ClearFormatting()` is a no-op so that macros compile, but a macro that sets `.Replacement.Font.Bold = True` has no `Font` to set. This is the single most common macro idiom that cannot port. Searching by formatting needs a per-run resolver (the `FormattingResolver` exists), and replacing with formatting is `Font`/`ParagraphFormat` applied to the replacement span.
* Missing flags: `MatchPrefix`, `MatchSuffix`, `IgnorePunct`, `IgnoreSpace`, `MatchFuzzy`, `MatchByte`, `CorrectHangulEndings`, `HitHighlight`, `Parent`, `Execute` `ReplaceMode.One` vs `All` (check). `FindWrap.Ask` throws; reasonable.

### 3.8 Fields

* `Field.Code` and `Field.Result` are objects with only `Text`. In Word they are `Range`s, so `Field.Result.Font.Bold = True` and `Field.Code.Text = ...` both work. Making them `RangeWrapper`-backed (or add a `Range` property to the part object) closes the main gap.
* Missing: `Kind` (`wdFieldKindNone/Hot/Warm/Cold`), `ShowCodes` and `Fields.ToggleShowCodes`, `Data`, `Select`, `Copy`/`Cut`, `Field.Parent`, `LinkFormat`, `OLEFormat`, `InlineShape`, `Fields.Update` per-field result codes (check what `Update` returns for a failed field), `FormFields` and `Range.FormFields`.
* The public mutable lists on `FieldWrapper` (`Codes`, `ResultText`, `ResultRuns`, `AllRuns`, `Separate`, `End`) look like internals. If they are public, hide them, because they cannot survive a change of how fields are modelled.
* Layout fields (`PAGE`, `NUMPAGES`, `PAGEREF`, `TOC`) stay dirty by design.

### 3.9 Bookmarks and Comments

* **Bookmark**: `StoryType`, `Column`, `Range` as settable (moving a bookmark), `Select`, `Bookmarks.Add` with an optional range (the selection in Word). `Bookmarks.Exists` and `ShowHidden` are good.
* **Comment**: `Reference`, `ShowTip`, `Edit()`, `Range` setter, and check that `Author`/`Initial`/`Date` are settable as in Word. `Comments.Format`, `ShowBy` and `Comments.Add` over ranges other than the supported ones.

### 3.10 Sections, HeaderFooter, PageSetup

* `Section`: `ProtectedForForms`, `Footnotes`/`Endnotes` settings, `Sections.PageSetup` (Word has it), `Section.Next` (not Word). Good otherwise.
* `HeaderFooter`: `Shapes`, `Range` (present), `Exists` (present), `LinkToPrevious` (present). Check that `Headers[HeaderFooterIndex.FirstPage]` behaves when `DifferentFirstPageHeaderFooter` is off, as Word does (it still returns the object).
* `PageSetup`: `SectionDirection`, `FirstPageTray`/`OtherPagesTray`, `CharsLine`/`LinesPage` (document grid), `LayoutMode`, `TwoPagesOnOne`, `BookFoldPrinting`, `BookFoldRevPrinting`, `BookFoldPrintingSheets`, `GutterStyle`, `TogglePortrait()`, `SetAsTemplateDefault()`. `Document.PageSetup` forwards only the common members.
* `PageNumbers.Add` does not support Inside/Outside.

### 3.11 Document

* Missing: `Saved`, `SaveAs2`, `Type`, `HasPassword`, `Password`, `WritePassword`, `ReadOnlyRecommended`, `DefaultTabStop`, `AttachedTemplate`, `UpdateStylesOnOpen`, `CopyStylesFromTemplate`, `StoryRanges`, `Repaginate`, `Compare`, `Merge`, `AcceptAllRevisions`/`RejectAllRevisions` (with the Revisions object), `Shapes`, `InlineShapes`, `Hyperlinks`, `Footnotes`, `Endnotes`, `Lists`, `ListParagraphs`, `ListTemplates`, `ContentControls`, `TablesOfContents`, `TablesOfFigures`, `Indexes`, `FormFields`, `CustomXMLParts`, `DocumentTheme` (`ApplyTheme`), `Frames`, `Subdocuments`, `Windows`/`ActiveWindow`/`Selection`, `Undo`/`Redo`, `PrintOut`, `ExportAsFixedFormat`, `CheckSpelling`, `Select`, `Range()` with no arguments (C5).
* `BuiltInDocumentProperties` exposes typed members and a name indexer. Word's collection also enumerates and has `Count`, and each item is a `DocumentProperty` with `Name`, `Type`, `Value`, `LinkToContent`. Add enumeration and an item object so `For Each p In BuiltInDocumentProperties` ports.

## 4. Non-Word surface (status after Phase 1)

Done: `Paragraph.InsertBreak` takes `BreakType`; `Cell.BackgroundColor`, `Cell.MergeDown`/`MergeRight` and `Table.InsertParagraphAfterSelf` are removed (each is `Shading`, `Merge` with `Next()` or `Table.Cell`, and `Range.InsertParagraphAfter`); `OpenXmlExtensions` is internal. `StoryMap`, `Anchor`, `FoundMatch` and `UnitNavigator` are internal already.

Kept on purpose, because one or two other members cannot do the same thing:

* `Table.AppendRow` and `Row.Copy`: together they copy a row with its content; `Rows.Add()` adds an empty row.
* `Table.MergeTable`: Word merges tables by deleting the paragraph between them.
* `Table.Width`: it works out a width in points for fixed and percentage widths; `PreferredWidth` is in the unit of `PreferredWidthType`.
* `Row.ClearText`, `Paragraph.DeleteBreaks`, `Paragraph.CopyPropertiesTo`.
* `Paragraph.InsertSectionBreak`: it ends the section with the paragraph. `Sections.Add(range)` and `Range.InsertBreak` insert a new paragraph for the break.

`ConversionHelper` and `PositionIndex` (with the two `CreatePositionIndex` methods) are internal too. Still public but not Word's: `FieldPartWrapper` (stands in for the `Range` that `Field.Code`/`Result` are in Word), `LineWidth`, `WordColor`, `ColorHelper`, `NamedColor`, `ColorIndex` and `BorderLineWidth` (Word's constants and `RGB`, in C# form).

## 5. Limitations of the library as it stands

**Cannot be closed without a layout engine.** `Pages` and `Lines` statistics, page and line `Information`, `GoTo` by page or line, `Selection` view-dependent movement, `Repaginate`, `PAGE`/`NUMPAGES`/`PAGEREF`/`TOC` results, printing, spelling and grammar.

**Not implemented at all (the objects that do not exist yet).** Hyperlinks, footnotes and endnotes (and their stories), lists and `ListFormat`/`ListTemplate`, content controls, inline shapes and shapes (pictures, text boxes, drawings), revisions, tables of contents, figures and authorities, indexes, form fields, frames, subdocuments, charts, equations (`OMaths`), document themes, custom XML parts, `Windows`/`View`/`Selection`, `Application`/`Documents`/`Templates`/`AutoCorrect`/`Options`.

**Behavioural limits of objects that do exist.**

* Track changes: edits are never recorded as revisions, so `TrackRevisions = true` throws (reading it and setting it to false work). Tracked-deleted text is counted in positions but is not exposed as `Revisions`.
* `Find` cannot search or replace by formatting and has no sounds-like or word-forms matching.
* Legacy `hMerge` tables are rejected by merge, split and column delete. `Cell.Delete` supports `ShiftCells.Left` only and not on vertically merged cells.
* `Range.Text` can be set for cells, paragraphs, bookmarks, headers and footers, the document and spans, but not for rows and tables.
* `Comments.Add` and `Bookmarks.Add` accept only some range kinds. Adding a table, paragraph or insert inside a hyperlink or other run-holding element, and a paragraph split inside a hyperlink, throw.
* `Field.Code`/`Result` are text only. Field results are computed for a fixed set of field types.
* Styles: names are matched in English, and `Styles[...]` for a built-in name depends on the catalogue. A document in another language that stores localised names may not match.
* Colours, line widths and style references use library types, not Word's (section 2).
* Indexes are 0-based by decision. Error types are .NET exceptions, not Word's COM error numbers, so `On Error` translations will differ.
* Not thread-safe per document (the span registry is, the document is not). Large documents: a span outside one paragraph, cell or table builds a story map of the whole story on each operation, so macros that loop over `Range(i, i+1)` in a long document will be slow. Worth a benchmark before publishing.
* Output is validated for Word via `OpenXmlValidator` in the tests; there is no round trip through real Word in CI.

**Documentation that is stale.** The README says `Cell.Delete` is not implemented (it is, partly). Check the "Not supported" list against the code after each phase.

## 6. Suggested order of work

Phase 1 changes shape and so should come before the first publish. Everything after it is additive.

**Phase 1 — conventions (breaking, done before publish).**
1. C1 colours: `WdColor`, `RGB()`, int overloads, `ColorIndex`.
2. C2 style objects: getters return `StyleWrapper`, setters accept name, object or `BuiltInStyle`.
3. C3 `LineWidth` enum.
4. Section 4: make `Paragraph.InsertBreak` use `BreakType`, hide internals, remove `Cell.BackgroundColor` duplicate, fix `Document.Range()` optional arguments.
5. `Field.Code`/`Field.Result` as ranges.

**Phase 2 — fill in the existing objects (additive).**
1. `Range`: `Delete(Unit, Count)`, `Next`/`Previous`, `MoveStartUntil` and friends, `IsEqual`/`InRange`, `Case`, layout-free `Information`, `GoTo`, `Sections`/`Columns`, `FormattedText`, `InsertDateTime`/`Symbol`/`CrossReference`/`InsertXML`/`InsertFile`.
2. `Font.Highlight` and bidi/East Asian font members; `Borders`/`Shading` on `Range`, `Paragraph` and `Style`.
3. `Paragraph`/`ParagraphFormat` completeness and ribbon helpers (`Indent`, `Space15`, `OpenUp` ...).
4. `Cells` / `Rows` / `Columns` as full collections (set-wise formatting, `Cells.Merge`, `SetWidth`, `DistributeHeight`), `Table.Sort`, `ConvertToText`, `Range.ConvertToTable`.
5. `Find.Font/ParagraphFormat/Style/Highlight` and `Replacement.*`.
6. `Style` completeness (`Description`, `Priority`, `Visibility`, `Linked`, `Table`).
7. Document members, `DocumentProperties` enumeration, `Saved`.
8. A normalising pass for `hMerge`.

**Phase 3 — behaviour that macros expect.**
1. Track changes: write `w:ins`/`w:del` when `TrackRevisions` is on, then a `Revisions` object (`Accept`, `Reject`, `Author`, `Date`, `Type`).
2. A light `Application` / `Documents` / `ActiveDocument` facade, and a decision about `Selection` (below).

**Phase 4 — new objects, in this order.** Hyperlinks, then lists (`ListFormat`, `ListTemplate`), footnotes and endnotes, content controls, inline shapes and pictures, TOC, and shapes.

**Decision for you: `Selection`.** Much recorded VBA is `Selection.TypeText`, `Selection.MoveDown`, `Selection.Find`. A logical-selection object (a `RangeWrapper` plus insertion-point state, with `TypeText`, `TypeParagraph`, `Move*`, `HomeKey`/`EndKey`, `Collapse` and `Find`) could be built on the existing span machinery without layout. Vertical movement and page keys would stay unsupported. It would make ported recorded macros far more familiar, but it adds mutable global state to an otherwise functional API. I would build it after Phase 2, and only if you want recorded macros to be a target.
