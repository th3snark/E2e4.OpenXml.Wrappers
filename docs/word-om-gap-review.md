# Word object model gap review

> **Status (November 2026, after 1.0.1).** This is the checklist for the gap work in releases 1.1 to 1.5. Every member of every Word object below is marked done, planned (with its plan step), out of scope or decided against. When a step lands, move its members to "Done" in the same pull request. Done in 1.0 since the first version of this review: colours (`WordColor`, `ColorHelper.RGB`, `NamedColor`, `ColorIndex`), style objects (`StyleObject`), `LineWidth`/`BorderLineWidth`, `Document.Range()`, `Range.Delete(unit, count)`, nullable `Paragraph` formatting, `Field.Code`/`Result` with `Range` and `Font`, and `TrackRevisions = true` throwing. The README, not this file, describes what the library does today.

Scope: the public API of `OpenXml.Wrappers.Word` at 1.0.1 against the Word VBA object model, first the objects that exist, then the objects that are missing, then the limitations and the order of work.

Method: the member lists come from the Microsoft Learn Word VBA reference (`https://learn.microsoft.com/en-us/office/vba/api/word.<object>`), fetched on 4 November 2026. Several object pages (`Font`, `ParagraphFormat`, `Paragraph` among them) carry no member list, so the lists were taken from the reference's table of contents, which has one page per member, and cross-checked against the object pages that do list members. `CustomXMLParts`, `CustomXMLPart`, `DocumentProperties` and `DocumentProperty` are Office objects (`office.<object>`). The library side is every public member of `OpenXml.Wrappers/Word/*.cs`. Two things in the current reference are worth knowing:

* It has no page for `Font.Color`, `Comment.Author`, `Comment.Initial` or `Comment.Delete` (the comment author is now `Comment.Contact`). The Word type library still exposes them and so does this library; they are not listed below.
* `Item(index)` is the indexer here (README: there is no `Item()` method). A collection whose indexer exists counts `Item` as done.

## 1. Summary

The objects that exist are well covered where macros use them most: `Range` navigation and editing, `Find`, `Font`, `ParagraphFormat`, tables, fields, comments, bookmarks, sections and page setup have Word's names, units and semantics. What is left falls into five groups:

1. **Missing members on objects that exist.** The largest clusters are `Range` (step 6), the paragraph formatting members that `Paragraph` and `Paragraphs` do not forward and the borders and shading of paragraphs, characters and styles (step 7), `Cells`/`Rows`/`Columns` as formatting collections (step 8), and formatting in `Find`/`Replacement` plus the rest of `Style` (step 9).
2. **Objects that do not exist yet.** Hyperlinks, lists, footnotes and endnotes, content controls and custom XML parts, inline shapes, tables of contents, floating shapes, a logical `Selection` and an `Application`/`Documents` facade (steps 11 to 20).
3. **Members that need a layout engine.** Pages, lines, `Repaginate`, page and line `Information`, page-number results. Out of scope.
4. **Members that belong to Word the application, not the file.** UI, windows, printing, the clipboard, proofing, events, VBA, servers and mail. Out of scope.
5. **Features excluded by decision.** Track changes and revisions, mail merge, frames, subdocuments, indexes and tables of authorities, legacy form fields, equations (`OMaths`), charts and SmartArt, and themes beyond the default. Out of scope.

## 2. Conventions

These cut across objects and are not repeated in the tables.

| # | Convention | Status |
|---|---|---|
| C1 | Colours: `WordColor` (hex string, `RGB` int or `NamedColor`), `ColorIndex`, `ColorHelper.RGB`. | Done in 1.0. |
| C2 | Style references: the `Style` properties are names, the object is `StyleObject`. | Done in 1.0. Setters accepting a `StyleWrapper` or `BuiltInStyle` are planned in step 9. |
| C3 | Line widths: `LineWidth` (points) with Word's `BorderLineWidth` constants. | Done in 1.0. |
| C4 | `Application`, `Documents`, `ActiveDocument`, `Selection`. | Planned: `Selection` in step 19, the facade in step 20. |
| C5 | `Document.Range()` with optional arguments. | Done in 1.0. |
| C6 | `Wrapper` suffixes (`RangeWrapper` for `Range`). | Kept. Documenting `using Range = OpenXml.Wrappers.Word.RangeWrapper;` aliases in the README is planned in step 20. |
| C7 | `TrackRevisions = true` throws because edits are not recorded as revisions. | Done in 1.0. Recording revisions is out of scope. |
| C8 | Every object has `Application` and `Creator`. | `Application` is planned in step 20 (it returns the facade). `Creator` is decided against: it is a 32-bit Macintosh creator code with no meaning outside COM. Neither is repeated in the tables below. |
| C9 | 0-based indexes, points, .NET exceptions instead of COM error numbers. | Done, by decision (see section 5). |

## 3. Objects that exist

"Done" means the member exists with Word's meaning; a note says where it differs. In "Planned", the step number is the plan step (section 6). Out of scope and decided-against rows give the reason once per group.

### 3.1 Range
<!-- learn: range -->

| Status | Members |
|---|---|
| Done | `Collapse`, `Delete` (also `Delete(unit, count)`, which returns the count), `Expand`, `InsertAfter`, `InsertBefore`, `InsertBreak`, `InsertParagraphAfter`, `InsertParagraphBefore`, `Move`, `MoveEnd`, `MoveStart`, `SetRange`, `Bookmarks`, `Cells`, `Characters`, `Comments`, `Duplicate`, `End`, `Fields`, `Find`, `Font`, `ParagraphFormat`, `Paragraphs`, `Rows`, `Sentences`, `Start`, `StoryLength`, `Style` (the name; the object is `StyleObject`), `Tables`, `Text`, `Words`, `ComputeStatistics` (words, characters, characters with spaces and paragraphs; lines and pages throw. Word returns 0 for a range that includes the end of a cell, which its own page calls inaccurate; the cell is counted here), `EndOf`, `InRange`, `InStory`, `InsertParagraph`, `IsEqual`, `MoveEndUntil`, `MoveEndWhile`, `MoveStartUntil`, `MoveStartWhile`, `MoveUntil`, `MoveWhile`, `Next`, `Previous`, `StartOf`, `WholeStory`, `Case` (lower, upper, title word, title sentence and toggle; `CharacterCase`), `InsertSymbol` (a `w:sym` for an ANSI symbol; a character in a font for a Unicode one; no `Bias`), `InsertAlignmentTab` (a `w:ptab`, one position, no text; Word shows it as the character 0), `InsertDateTime` (a `DATE` field as Word writes it, or text; no full-width, language or calendar arguments), `Information` for the layout-free types (`WithInTable`, the row and column numbers and the most rows and columns, `InCommentPane`, `InHeaderFooter`, `HeaderFooterType`, `InFootnote`, `InEndnote`; `AtEndOfRowMarker` is not done: Word gave False for the end-of-row mark), `BookmarkID`, `PreviousBookmarkID`, `Sections`, `FormattedText` (copies paragraphs, tables, fields, pictures, links and the styles they use, between ranges and documents; bookmarks stay with the original, as in Word; comments are not copied, where Word copies them; a table the range only touches is copied whole; numbering is dropped when copying to another document; a section break is not copied), `WordOpenXML` (a flat OPC package), `InsertXML` (flat OPC only, not Word 2003 XML), `InsertFile` (no range, confirm, link or attachment arguments), `ImportFragment`, `ExportFragment` (documents only). `Move`, `MoveStart`, `MoveEnd` and `Expand` also take `TextUnit.Story`. `Expand` returns nothing, where Word returns the number of characters added |
| Done (step 6, fourth part) | `Document` (the wrapper the document was opened or created with; null for elements that are not in a document), `Parent` (the same), `TextRetrievalMode` (`IncludeFieldCodes`, `IncludeHiddenText`, `ViewType`, `Duplicate`) |
| Not yet | `InsertCaption` and `InsertCrossReference` (see the note under Range) |
| Planned, step 7 | `Bold`, `Italic`, `Underline`, `BoldBi`, `ItalicBi`, `HighlightColorIndex`, `Borders`, `Shading`, `LanguageID`, `LanguageIDFarEast`, `LanguageIDOther`, `NoProofing`, `EmphasisMark`, `FitTextWidth`, `DisableCharacterSpaceGrid` |
| Done (step 8e) | `ConvertToTable` |
| Planned, step 8 | `Sort`, `SortAscending`, `SortDescending` (paragraphs; the table `Sort` is done), `Columns`, `TopLevelTables`, `IsEndOfRowMark`, `Orientation` |
| Done (step 9f) | `CharacterStyle`, `ParagraphStyle`, `TableStyle` (read only, as in Word; names, with `CharacterStyleObject`, `ParagraphStyleObject` and `TableStyleObject`). `CharacterStyle` is "Default Paragraph Font" for plain text, null where characters differ, and a collapsed range takes the character before it; spaces at the end of a range are ignored. `TableStyle` is the table's style for a range inside one table and "Table Normal" for any other. Not measured: several trailing spaces, and `TableStyle` after `SetDefaultTableStyle` |
| Planned, step 10 | `PageSetup` |
| Done (step 11) | `StoryType`, `NextStoryRange` (the same header or footer in the next section that has one of its own; null for other stories) |
| Done (step 12) | `Hyperlinks` |
| Done | `ListFormat` (step 13), `Footnotes`, `Endnotes`, `FootnoteOptions`, `EndnoteOptions` (step 14), `ContentControls`, `ParentContentControl` (step 15a) |
| Done (step 16) | `InlineShapes` |
| Planned, steps 13c to 19 | `ListParagraphs`, `ListStyle`, `SetListLevel` (13c); `ShapeRange` (18); `Select` (19) |
| Out of scope | Layout and rendering: `TextVisibleOnScreen`, `ExportAsFixedFormat`, `ExportAsFixedFormat2`, `ExportAsFixedFormat3`, `EnhMetaFileBits`, `CopyAsPicture`. Clipboard and UI: `Copy`, `Cut`, `Paste`, `PasteAndFormat`, `PasteAppendTable`, `PasteAsNestedTable`, `PasteExcelTable`, `PasteSpecial`, `ShowAll`, `Relocate` (outline view), `GoToEditableRange`, `LookupNameProperties` (address book), `InsertAutoText` (needs the Normal template's building blocks). Proofing: `CheckGrammar`, `CheckSpelling`, `CheckSynonyms`, `GetSpellingSuggestions`, `DetectLanguage`, `LanguageDetected`, `GrammarChecked`, `GrammaticalErrors`, `SpellingChecked`, `SpellingErrors`, `ReadabilityStatistics`, `SynonymInfo`, `ConvertHangulAndHanja`, `TCSCConverter`. Track changes and co-authoring: `Revisions`, `Conflicts`, `Locks`, `Updates`. Excluded features: `FormFields`, `Frames`, `Subdocuments`, `NextSubdocument`, `PreviousSubdocument`, `OMaths`, `InsertDatabase` (mail merge), `HTMLDivisions`, `Scripts` (HTML) |
| Decided against | `AutoFormat` (Word's AutoFormat rules are undocumented), `Calculate` (evaluates arithmetic in the text; rare), `SortByHeadings` (rare), `Editors` (editing permissions), `ID` (an HTML id), `XML` (the Word 2003 XML format). East Asian layout, which needs reference behaviour we cannot test: `CharacterWidth`, `CombineCharacters`, `HorizontalInVertical`, `TwoLinesInOne`, `Kana`, `PhoneticGuide`, `ModifyEnclosure` |

**InsertCaption and InsertCrossReference: not yet.** Word 16 behaves differently from what the documentation says, so the behaviour is recorded here first. A caption is a paragraph in the Caption style with a `SEQ <label> \* ARABIC` field: `InsertCaption('Figure', ': first', '', 0, False)` on a collapsed range at the start of a paragraph put `Figure 1: first` in that paragraph, which became a Caption paragraph, without adding a paragraph; on a whole paragraph with the position below it added a paragraph after it. Excluding the label gave only the number; an unknown label raised 'Command failed'. A cross reference writes a hidden bookmark `_Ref<number>` (round the label and number, the caption text or the whole caption, or a heading's text) and a `REF _Ref... \h` field (`PAGEREF` for the page number, `\w` for a heading's number, `\p` for above or below), or a `REF` to a bookmark by name. The kinds `wdOnlyLabelAndNumber` and `wdAboveBelow` with a caption raised 'Command failed', and `GetCrossReferenceItems` listed an Equation caption under 'Table'. They are not in a step yet; the markup would be checked scenario by scenario as the other fields are.

**GoTo, GoToNext and GoToPrevious: not planned yet.** Word 16 gives results that do not follow one rule, so they are recorded here rather than copied. What it does, measured on a document with three headings (at 6, 20 and 44), two comments, two fields, two tables and three sections: the result is a new collapsed range at the start of the item (a bookmark gives the bookmark's range and needs a name, and a field can be found by the first word of its code); the range it is called on does not move. Next is the first item that starts after the range's start. Previous is the last item that ends at or before it, so from inside a heading it skips that heading. When nothing is found the result is the range collapsed at its start, except for sections, where a number past the last gives the last section. `Last` gives the last field, comment, table and section, but nothing for headings, and Previous for a table before the range gave nothing in one test. Bookmarks with `Next` or `Previous` raise "This bookmark does not exist". A Which of `wdGoToLast` with a bookmark is an error, and `GoToNext(wdGoToBookmark)` is "Bad parameter". Layout targets (page, line, percent, errors) are out of scope.

### 3.2 Document
<!-- learn: document -->

`WordprocessingDocumentWrapper` is the `Document`.

| Status | Members |
|---|---|
| Done | `Close`, `Save`, `SaveAs2` (as `SaveAs(path, format)`), `Protect`, `Unprotect`, `Range`, `ComputeStatistics` (words, characters, paragraphs; lines and pages throw), `Bookmarks`, `BuiltInDocumentProperties`, `Characters`, `Comments`, `CompatibilityMode`, `Content`, `CustomDocumentProperties`, `Fields`, `FullName`, `Name`, `PageSetup` (reads the first section, sets every section), `Paragraphs`, `Path`, `ProtectionType`, `ReadOnly`, `Sections`, `Sentences`, `Styles`, `Tables`, `TrackRevisions` (read, and switch off; `true` throws), `Variables`, `Words`, `WordOpenXML` |
| Planned, step 6 | `GoTo`, `GetCrossReferenceItems` |
| Done (step 9f) | `DefaultTableStyle` (null where none is set; Word gives a style with no name), `SetDefaultTableStyle` (writes `w:defaultTableStyle`, as Word does, even for "Table Normal"; `setInTemplate` has no effect; a style that is not a table style throws) |
| Done (step 10a) | `DefaultTabStop` (points; at most 1638, the most the SDK stores), `AutoHyphenation`, `HyphenateCaps`, `ConsecutiveHyphensLimit`, `UpdateStylesOnOpen` (`w:linkStyles`), `EmbedTrueTypeFonts`, `SaveSubsetFonts`, `DoNotEmbedSystemFonts`, `RemovePersonalInformation`, `RemoveDateAndTime` (these three only keep the setting; Word acts on them when it saves), `Type` (`WordDocumentType`, from the package type), `SaveFormat`, `DeleteAllComments`, `SetCompatibilityMode`. All written as Word 16 writes them, in the `document-settings` scenario |
| Done (step 10e) | `Convert` (to compatibility mode 15; throws when the document is in it already, where Word makes the command unavailable; matches the `document-convert` scenario) |
| Decided against | `Compatibility` (the legacy compatibility options: Word 16 answers false to every one in a new document and refuses to set them in a .docx), `Final` (Word asks before marking a document as final, so it cannot be measured unattended), `Saved` (needs every edit to be tracked as a change), `RemoveDocumentInformation` (a privacy sweep over comments, properties and more; `DeleteAllComments`, the document properties and `RemovePersonalInformation` cover the parts that matter) |
| Decided against | `AttachedTemplate`, `CopyStylesFromTemplate`, `UpdateStyles` (no templates), `WritePassword`, `WriteReserved`, `ReadOnlyRecommended` (Word 16 did not write them to `settings.xml` when set; they belong to saving), `HyphenationZone` (Word 16 ignored a value of 18 points when set) |
| Planned, steps 11 to 20 | `StoryRanges` (done in step 11); `Hyperlinks` (done in step 12); `ConvertNumbersToText`, `CountNumberedItems`, `RemoveNumbers` (13c); `CustomXMLParts` (done in step 15c); `SelectLinkedControls`, `SelectUnlinkedControls` (planned later, with `CustomXMLNode`); `ContentControls`, `SelectContentControlsByTag`, `SelectContentControlsByTitle` (done in step 15a); `InlineShapes` (done in step 16); `TablesOfContents`, `TablesOfFigures` (17); `Shapes`, `Background` (18); `Select` (19); `Parent`, `Activate` (20) |
| Out of scope | Layout: `Repaginate`, `FitToPages`, `FreezeLayout`, `ReadingModeLayoutFrozen`, `ConvertAutoHyphens`, `ManualHyphenation`, `ExportAsFixedFormat`, `ExportAsFixedFormat2`, `ExportAsFixedFormat3`. Printing: `PrintOut`, `PrintPreview`, `ClosePrintPreview`, `PrintFormsData`, `PrintPostScriptOverText`, `PrintRevisions`. UI and windows: `ActiveWindow`, `Windows`, `Undo`, `Redo`, `UndoClear`, `ViewCode`, `ViewPropertyBrowser`, `WebPagePreview`, `ToggleFormsDesign`, `FormsDesign`, `ReturnToLastReadPosition`, `ReadingLayoutSizeX`, `ReadingLayoutSizeY`, `PresentIt`, `Broadcast`, `CommandBars`, `FormattingShowClear`, `FormattingShowFilter`, `FormattingShowFont`, `FormattingShowNextLevel`, `FormattingShowNumbering`, `FormattingShowParagraph`, `FormattingShowUserStyleName`, `ClickAndTypeParagraphStyle`, `StyleSortMethod`, `DeleteAllCommentsShown`, `FollowHyperlink`, `DisableFeatures`, `DisableFeaturesIntroducedAfter`, `UserControl`, `CheckConsistency`, `DocumentInspectors`. Proofing: `CheckGrammar`, `CheckSpelling`, `DetectLanguage`, `LanguageDetected`, `GrammarChecked`, `SpellingChecked`, `GrammaticalErrors`, `SpellingErrors`, `ShowGrammaticalErrors`, `ShowSpellingErrors`, `ReadabilityStatistics`, `EmbedLinguisticData`, `ActiveWritingStyle`. Track changes and review: `Revisions`, `AcceptAllRevisions`, `AcceptAllRevisionsShown`, `RejectAllRevisions`, `RejectAllRevisionsShown`, `TrackFormatting`, `TrackMoves`, `Compare`, `Merge`, `EndReview`, `ReplyWithChanges`, `SendForReview`, `OriginalDocumentTitle`, `RevisedDocumentTitle`, `CoAuthoring`. Servers, mail and the host: `AddToFavorites`, `CanCheckin`, `CheckIn`, `CheckInWithVersion`, `LockServerFile`, `GetWorkflowTasks`, `GetWorkflowTemplates`, `Post`, `SendFax`, `SendFaxOverInternet`, `SendMail`, `Email`, `MailEnvelope`, `ServerPolicy`, `DocumentLibraryVersions`, `ContentTypeProperties`, `Sync`, `Permission`, `Signatures`, `SensitivityLabel`, `Reload`, `ReloadAs`, `IsInAutosave`, `AutoSaveOn`, `Container`, `Research`, `SmartDocument`. Encryption (the Open XML SDK does not encrypt packages): `HasPassword`, `Password`, `EncryptionProvider`, `PasswordEncryptionAlgorithm`, `PasswordEncryptionFileProperties`, `PasswordEncryptionKeyLength`, `PasswordEncryptionProvider`, `SetPasswordEncryptionOptions`. Mail merge, letters and legacy forms: `MailMerge`, `Envelope`, `DataForm`, `CreateLetterContent`, `GetLetterContent`, `SetLetterContent`, `RunLetterWizard`, `FormFields`, `ResetFormFields`, `SaveFormsData`. VBA and events: `RunAutoMacro`, `HasVBProject`, `VBProject`, `VBASigned`, `CodeName`, and the events `BuildingBlockInsert`, `ContentControlAfterAdd`, `ContentControlBeforeContentUpdate`, `ContentControlBeforeDelete`, `ContentControlBeforeStoreUpdate`, `ContentControlOnEnter`, `ContentControlOnExit`, `New`, `Open`, `XMLAfterInsert`, `XMLBeforeDelete` (and the Close and Sync events). Themes: `ApplyDocumentTheme`, `ApplyTheme`, `RemoveTheme`, `ActiveTheme`, `ActiveThemeDisplayName`, `DocumentTheme`, `LockTheme`. Excluded objects: `Frames`, `Subdocuments`, `IsMasterDocument`, `IsSubdocument`, `Indexes`, `TablesOfAuthorities`, `TablesOfAuthoritiesCategories`, `OMaths`, `OMathBreakBin`, `OMathBreakSub`, `OMathFontName`, `OMathIntSubSupLim`, `OMathJc`, `OMathLeftMargin`, `OMathNarySupSubLim`, `OMathRightMargin`, `OMathSmallFrac`, `OMathWrap`, `UseMathDefaults`, `ChartDataPointTrack`, `DeleteAllInkAnnotations`. HTML, text and XML formats (the library writes packages only): `HTMLDivisions`, `Frameset`, `DefaultTargetFrame`, `WebOptions`, `Scripts`, `StyleSheets`, `OpenEncoding`, `SaveEncoding`, `TextEncoding`, `TextLineEnding`, `OptimizeForWord97`, `TransformDocument`, `XMLSaveThroughXSLT`, `XMLUseXSLTWhenSaving`, `XMLShowAdvancedErrors`, `XMLSchemaReferences`, `SelectNodes`, `SelectSingleNode` (custom XML markup, which Word no longer supports). Templates: `MakeCompatibilityDefault` |
| Decided against | AutoFormat and style galleries: `AutoFormat`, `AutoFormatOverride`, `Kind`, `ApplyQuickStyleSet2`, `SaveAsQuickStyleSet`, `LockQuickStyleSet`. Formatting restrictions and editing permissions (document protection covers editing): `EnforceStyle`, `RemoveLockedStyles`, `DeleteAllEditableRanges`, `SelectAllEditableRanges`. Legacy conversions: `DowngradeDocument`, `ConvertVietDoc`. `Bibliography` (rare), `CurrentRsid` (internal to Word). Drawing grid settings: `GridDistanceHorizontal`, `GridDistanceVertical`, `GridOriginFromMargin`, `GridOriginHorizontal`, `GridOriginVertical`, `GridSpaceBetweenHorizontalLines`, `GridSpaceBetweenVerticalLines`, `SnapToGrid`, `SnapToShapes`. East Asian line breaking: `JustificationMode`, `KerningByAlgorithm`, `FarEastLineBreakLanguage`, `FarEastLineBreakLevel`, `NoLineBreakAfter`, `NoLineBreakBefore` |

### 3.3 Font
<!-- learn: font -->

Reads give effective formatting, `Direct` gives direct formatting, and a mixed value reads as null.

| Status | Members |
|---|---|
| Done | `Reset`, `AllCaps`, `Bold`, `ColorIndex`, `DoubleStrikeThrough`, `Emboss`, `Engrave`, `Hidden`, `Italic`, `Kerning`, `Name`, `Outline`, `Position`, `Scaling`, `Shadow`, `Size`, `SmallCaps`, `Spacing`, `StrikeThrough`, `Subscript`, `Superscript`, `Underline`, `UnderlineColor`, and from step 7: `Grow` and `Shrink` (Word's sizes; the complex script size follows), `Duplicate` (also `Range.Font = font`), `BoldBi`, `ItalicBi`, `SizeBi`, `ColorIndexBi` (the same as `ColorIndex`: Word keeps one colour), `NameAscii`, `NameOther`, `NameFarEast`, `NameBi` (theme fonts read `+Body`, `+Headings` and `+Body CS`), `EmphasisMark`, `DisableCharacterSpaceGrid`. `Name` also changes the complex script font when that was set on its own, as Word does |
| Done (step 7c) | `Borders` (one border all round the text: only `Top`) and `Shading` of characters, direct formatting only |
| Planned, step 7 | `Parent`, `TextColor`, and the OpenType features stored as single `w14` elements: `Ligatures`, `NumberForm`, `NumberSpacing`, `StylisticSet`, `ContextualAlternates` |
| Out of scope | `SetAsTemplateDefault` (writes the attached template) |
| Decided against | The 2010 text effects, each an object of its own in `w14` markup: `Fill`, `Glow`, `Line`, `Reflection`, `TextShadow`, `ThreeD`. `DiacriticColor` (complex-script only) |

### 3.4 ParagraphFormat
<!-- learn: paragraphformat -->

| Status | Members |
|---|---|
| Done | `Reset`, `Alignment`, `FirstLineIndent`, `Hyphenation`, `KeepTogether`, `KeepWithNext`, `LeftIndent`, `LineSpacing`, `LineSpacingRule`, `MirrorIndents`, `NoLineNumber`, `OutlineLevel`, `PageBreakBefore`, `RightIndent`, `SpaceAfter`, `SpaceAfterAuto`, `SpaceBefore`, `SpaceBeforeAuto`, `TabStops`, `WidowControl` |
| Done (step 7b) | `OpenUp`, `CloseUp`, `OpenOrCloseUp` (12 points before), `Space1`, `Space15`, `Space2`, `TabIndent`, `TabHangingIndent` (by default tab stops, which is `Settings.DefaultTabStop`), `Duplicate` (also `Range.ParagraphFormat = format`), `ReadingOrder`, `WordWrap`, `AutoAdjustRightIndent`, `BaseLineAlignment`, `TextboxTightWrap`, and the East Asian switches `AddSpaceBetweenFarEastAndAlpha`, `AddSpaceBetweenFarEastAndDigit`, `DisableLineHeightGrid`, `FarEastLineBreakControl`, `HalfWidthPunctuationOnTopOfLine`, `HangingPunctuation` (each a single element, defaults as Word's) |
| Done (step 7c) | `Borders` (the four edges) and `Shading` of paragraphs, direct formatting only (a style's borders are not read, and a style's `ParagraphFormat.Borders` throws) |
| Planned, step 7 | `Parent`, `CollapsedByDefault` |
| Decided against | `CharacterUnitFirstLineIndent`, `CharacterUnitLeftIndent`, `CharacterUnitRightIndent`, `LineUnitAfter`, `LineUnitBefore`, `IndentCharWidth`, `IndentFirstLineCharWidth`: they store a count of characters or lines for the East Asian grid and Word 16 is inconsistent about them (`IndentCharWidth(3)` changed nothing, `IndentFirstLineCharWidth(2)` set the indent but not the unit). They need reference behaviour we cannot reproduce |
| Planned, step 9 | `Style` |

### 3.5 Paragraph
<!-- learn: paragraph -->

In Word, `Paragraph` exposes the whole of `ParagraphFormat`; here it forwards only part of it. Step 7 forwards the rest.

| Status | Members |
|---|---|
| Done | `Next`, `Previous`, `Alignment`, `FirstLineIndent`, `Format`, `KeepTogether`, `KeepWithNext`, `LeftIndent`, `LineSpacing`, `LineSpacingRule`, `PageBreakBefore`, `Parent`, `Range`, `RightIndent`, `SpaceAfter`, `SpaceBefore`, `Style`, `WidowControl` |
| Done (step 7b) | everything `ParagraphFormat` has that is done, and `Indent` / `Outdent` (to the next or previous default tab stop), `Reset`, `OpenUp`, `CloseUp`, `OpenOrCloseUp`, `Space1`, `Space15`, `Space2`, `TabIndent`, `TabHangingIndent`, `Format` can be assigned |
| Done (step 7c) | `Borders`, `Shading` |
| Done (step 7d) | `OutlineDemote`, `OutlinePromote` and `OutlineDemoteToBody`: they change the style only (direct formatting stays). Heading n moves to n ± 1 within 1 to 9; a paragraph with a level of its own moves that level; body text becomes a heading one level below (demote) or at the level of (promote) the nearest heading before it, Heading 2 or Heading 1 if there is none |
| Planned, step 7 | `CollapseHeadingByDefault`, `IsStyleSeparator` |
| Decided against | the character and line unit members and `IndentCharWidth`, `IndentFirstLineCharWidth` (see ParagraphFormat) |
| Planned, step 13 | `JoinList`, `SeparateList`, `ListAdvanceTo`, `ResetAdvanceTo`, `ListNumberOriginal` |
| Out of scope | `CollapsedState` (view state), `SelectNumber` (UI), `DropCap` (drop caps are frames) |
| Decided against | `ID` (an HTML id) |

### 3.6 Paragraphs
<!-- learn: paragraphs -->

| Status | Members |
|---|---|
| Done | `Add` (also `Add(range)`, step 7d), `Item`, `Count`, `First`, `Last`, `Format`, `Parent`, `Style` |
| Done (step 7d) | `CloseUp`, `OpenUp`, `OpenOrCloseUp`, `DecreaseSpacing`, `IncreaseSpacing` (6 points before and after, not below 0), `Indent`, `Outdent`, `Reset`, `Space1`, `Space15`, `Space2`, `TabHangingIndent`, `TabIndent`, `OutlineDemote`, `OutlineDemoteToBody`, `OutlinePromote`, and the format properties of the collection (`Alignment`, `LeftIndent`, `KeepWithNext`, `WordWrap`, `Borders`, `Shading` and so on), each reading null where the paragraphs differ |
| Decided against | `IndentCharWidth`, `IndentFirstLineCharWidth`, `CharacterUnit*` and `LineUnit*` (see ParagraphFormat) |

### 3.7 Find
<!-- learn: find -->

| Status | Members |
|---|---|
| Done | `ClearFormatting`, `Execute` (Word's parameters in Word's order), `Forward`, `Found`, `MatchCase`, `MatchWholeWord`, `MatchWildcards`, `Replacement`, `Text`, `Wrap` (`FindWrap.Ask` throws) |
| Done (step 9b) | `Format`, `Font`, `ParagraphFormat`, `Style` (the name), `Highlight` (true or false; null for either) |
| Done (step 9e) | `MatchPrefix`, `MatchSuffix`, `IgnorePunct`, `IgnoreSpace`, `NoProofing`, `LanguageID`, `LanguageIDFarEast` |
| Done (step 9h) | `MatchPhrase` (white space in the text matches any white space or control characters, as measured: a run must be matched by one at least as long, the whole run is used when text follows; not used with wildcards) |
| Decided against | `LanguageIDOther` (Word 16 does not expose it on `Find` or `Replacement` over COM), `Parent` (returns the owning Word object; the library has no `Application` hierarchy for it until step 20) |
| Out of scope | `MatchSoundsLike`, `MatchAllWordForms` (need Word's linguistic data; they exist and throw when set to `true`), `HitHighlight`, `ClearHitHighlight` (UI), `Frame` (frames) |
| Decided against | `Execute2007` (the same as Execute). Script-specific matching, which needs reference behaviour we cannot test: `MatchFuzzy`, `SetAllFuzzyOptions`, `ClearAllFuzzyOptions`, `MatchByte`, `MatchAlefHamza`, `MatchControl`, `MatchDiacritics`, `MatchKashida`, `CorrectHangulEndings`, `HanjaPhoneticHangul` |

### 3.8 Replacement
<!-- learn: replacement -->

| Status | Members |
|---|---|
| Done | `ClearFormatting`, `Text` |
| Done (step 9b) | `Font`, `ParagraphFormat`, `Style` (the name), `Highlight` (true, false or null) |
| Done (step 9e) | `LanguageID`, `LanguageIDFarEast`, `NoProofing` |
| Decided against | `Parent` (see `Find`) |
| Out of scope | `Frame` (frames) |

### 3.9 Style, Styles, TableStyle, ConditionalStyle
<!-- learn: style -->

**Style.** `Name` and `StyleId` are extra and stay.

| Status | Members |
|---|---|
| Done | `Delete`, `BaseStyle`, `BuiltIn`, `Font`, `NameLocal`, `NextParagraphStyle`, `ParagraphFormat`, `Type` |
| Done (step 9a) | `AutomaticallyUpdate`, `InUse`, `LanguageID`, `LanguageIDFarEast`, `NoProofing`, `Linked`, `LinkStyle` (read only), `Locked`, `Priority`, `QuickStyle`, `UnhideWhenUsed`, `Visibility`, `NoSpaceBetweenParagraphsOfSameStyle`; `BaseStyle` and `NextParagraphStyle` can be set |
| Done (step 9c) | `Borders`, `Shading` (paragraph styles; Word gives `Nothing` for a character style, so they are null there. `Font.Borders`, `Font.Shading`, `ParagraphFormat.Borders` and `ParagraphFormat.Shading` of a style work too) |
| Done (step 9d) | `Table` (null for a style that is not a table style) |
| Decided against | `Parent` (see `Find`) |
| Planned, step 13 | `ListLevelNumber`, `ListTemplate`, `LinkToListTemplate` |
| Out of scope | `Frame` (frames) |
| Decided against | `Description` (the summary Word generates for its UI, such as "Font: Bold, Style: Quick Style") |

Style members as Word 16 behaves, measured: `Styles.Add` of a paragraph style also adds the character style linked to it (`<name> Char`, based on Default Paragraph Font), and a character or table style is based on the default of its type. `Priority` is one more than `w:uiPriority` (which is left out for 1). `Visibility` is true when the style is hidden (`w:semiHidden`): Word 16 reads false for Normal and Heading 1 and true for Heading 5. Language and no-proofing set on a paragraph style are also written to the linked style, and setting `BaseStyle` gives the linked style the character formatting of the new base (its own, not what it inherits), and so does `Priority` for the priority. Word does not mirror other changes by the library's `Font` of a style to the linked style, which it does for those it writes itself (not compared). `LinkStyle` is read only: Word refused to link a paragraph style to a character style that was made on its own.

Borders and shading of styles as Word 16 writes them: a paragraph style has `w:pBdr` and `w:shd` in its paragraph properties, a character style the border (`w:bdr`) and shading in its run properties. Word gives the shading of a paragraph style to the linked character style as well (as run shading), and the border of the font of a paragraph style, but not the shading of its font. `Style.Borders` and `Style.Shading` of a character style are null (Word returns `Nothing` for `Borders`; `Shading` was not measured). Not compared: borders of a table style.

<!-- learn: styles -->

**Styles.**

| Status | Members |
|---|---|
| Done | `Add`, `Item` (by index, name or `BuiltInStyle`), `Count` |
| Done (step 9g) | `BuiltInStyle` covers all of `WdBuiltinStyle` except List Bullet and List Number (1 to 5), whose numbering Word writes as a numbering part (planned with lists, step 13): index, TOA, list, list continue, envelope, HTML, body text variants, Message Header, Note Heading, Macro Text, Normal Indent, Block Text, Document Map, Normal (Web), `NormalObject` (Word 365 shows No Spacing for it) and the 20 built-in table styles, checked in the `built-in-styles` scenario |
| Decided against | `Parent` (see `Find`) |

<!-- learn: tablestyle -->

**TableStyle** (`Style.Table`).

| Status | Members |
|---|---|
| Done (step 9d) | `Condition` (`ConditionCode`), `Alignment`, `AllowBreakAcrossPage`, `Borders`, `BottomPadding`, `ColumnStripe`, `LeftIndent`, `LeftPadding`, `RightPadding`, `RowStripe`, `Shading`, `Spacing`, `TopPadding` |
| Decided against | `Parent` (see `Find`) |
| Decided against | `AllowPageBreaks` (not measured; `AllowBreakAcrossPage` is the row property), `TableDirection` (setting it through automation changed nothing in Word 16) |

Table styles as Word 16 writes them: `Styles.Add` of a table style writes an empty `w:tblPr`. `Alignment` and `Spacing` are written to the style's table properties and to its row properties, `AllowBreakAcrossPage = false` is `w:cantSplit` in the row properties, `LeftIndent` is `w:tblInd`, the stripes are `w:tblStyleRowBandSize` and `w:tblStyleColBandSize`, and `Borders` are `w:tblBorders`. The paddings and the shading go to the style's cell properties (`w:tcMar` and `w:shd`), not to `w:tblCellMar` or the table properties; a padding that is set writes only its own side. Reading gives only what the style sets itself.

<!-- learn: conditionalstyle -->

**ConditionalStyle** (`TableStyle.Condition`).

| Status | Members |
|---|---|
| Done (step 9d) | `Borders`, `BottomPadding`, `Font`, `LeftPadding`, `ParagraphFormat`, `RightPadding`, `Shading`, `TopPadding` (`Font.Borders`, `Font.Shading`, `Font.Duplicate` and the borders and shading of `ParagraphFormat` throw `NotSupportedException`; use `Borders` and `Shading` of the conditional format) |
| Decided against | `Parent` (see `Find`) |

Conditional formats as Word 16 writes them (`w:tblStylePr`, in the order whole table, first row, last row, first column, last column, vertical bands, horizontal bands, then the corner cells): `Font` is `w:rPr`, `ParagraphFormat` is `w:pPr`, and the shading, the borders and the paddings are in the cell properties with an empty `w:tblPr` beside them. Setting the shading or a border makes the borders that are not set `nil`; setting a padding makes the other three sides `nil`. Word's own results for combinations of several of these in one session were not the sum of the parts (a later border or padding dropped what an earlier one wrote), so only each on its own, and font, paragraph format and shading together, are compared.

### 3.10 Table and Tables
<!-- learn: table -->

**Table.** `AppendRow`, `MergeTable` and `Width` are extra (section 4).

| Status | Members |
|---|---|
| Done | `AutoFitBehavior`, `Cell`, `Delete`, `Split`, `AllowAutoFit`, `ApplyStyleColumnBands`, `ApplyStyleFirstColumn`, `ApplyStyleHeadingRows`, `ApplyStyleLastColumn`, `ApplyStyleLastRow`, `ApplyStyleRowBands`, `Borders`, `BottomPadding`, `Columns`, `Descr`, `LeftPadding`, `NestingLevel`, `Parent`, `PreferredWidth`, `PreferredWidthType`, `Range`, `RightPadding`, `Rows`, `Shading`, `Spacing`, `Style`, `Tables`, `Title`, `TopPadding`, `Uniform` |
| Done (step 8e) | `ConvertToText`, `Sort` |
| Planned, step 8 | `SortAscending`, `SortDescending` (Word sorts by the column of the selection, which the library does not have; use `Sort`), `ApplyStyleDirectFormatting` |
| Not done | `TableDirection`: set through automation Word 16 changed nothing in the table, so there is no markup to follow |
| Planned, step 19 | `Select` |
| Decided against | `AutoFormat`, `AutoFormatType`, `UpdateAutoFormat` (Word 97 table AutoFormat; table styles replace it), `ID` (an HTML id) |

<!-- learn: tables -->

**Tables.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count`, `NestingLevel`, `Parent` |

### 3.11 Rows, Row, Columns, Column, Cells, Cell
<!-- learn: rows -->

The collections are where tables are thinnest: Word's `Rows`, `Columns` and `Cells` format the whole set, and `Range.Cells` is the common way to format a block. Older `hMerge` merges are turned into `gridSpan` the first time a table is merged, split or loses a column (step 8f), so those no longer throw for them.

**Rows.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Alignment`, `AllowBreakAcrossPages`, `Count`, `First`, `HeadingFormat`, `Height`, `HeightRule`, `Last`, `LeftIndent`, `NestingLevel`, `Parent`, `Shading` |
| Done (step 8a) | `Delete`, `SetHeight` |
| Done (step 8c) | `Borders` |
| Done (step 8e) | `ConvertToText` |
| Planned, step 8 | `DistributeHeight` (Word measures the laid-out height, so it needs layout), `SetLeftIndent` (for all the rows: Word also sets the table indent; each `Row.SetLeftIndent` is done), `AllowOverlap`, `DistanceBottom`, `DistanceLeft`, `DistanceRight`, `DistanceTop`, `HorizontalPosition`, `VerticalPosition`, `RelativeHorizontalPosition`, `RelativeVerticalPosition`, `WrapAroundText`, `SpaceBetweenColumns`, `TableDirection` |
| Planned, step 19 | `Select` |

<!-- learn: row -->

**Row.** `ClearText` and `Copy` are extra (section 4).

| Status | Members |
|---|---|
| Done | `Delete`, `SetHeight`, `Alignment`, `AllowBreakAcrossPages`, `Borders`, `Cells`, `HeadingFormat`, `Height`, `HeightRule`, `Index`, `IsFirst`, `IsLast`, `Next`, `Parent`, `Previous`, `Range`, `Shading` |
| Done (step 8a) | `NestingLevel` |
| Done (step 8e) | `ConvertToText` |
| Done (step 8f) | `LeftIndent`, `SetLeftIndent` |
| Planned, step 8 | `SpaceBetweenColumns` |
| Planned, step 19 | `Select` |
| Decided against | `ID` (an HTML id) |

<!-- learn: columns -->

**Columns.**

| Status | Members |
|---|---|
| Done | `Add`, `DistributeWidth`, `Item`, `Count`, `First`, `Last`, `Parent` |
| Done (step 8a) | `Delete`, `NestingLevel`, `PreferredWidth` (in points it is the width of the columns, as in Word), `PreferredWidthType`, `Shading` (also the table's when the columns are all of them), `Width` |
| Done (step 8b) | `SetWidth` |
| Done (step 8c) | `Borders` |
| Needs layout | `AutoFit` (Word measures the text; on empty cells it changes nothing) |
| Planned, step 19 | `Select` |

<!-- learn: column -->

**Column.**

| Status | Members |
|---|---|
| Done | `Delete`, `Borders`, `Cells`, `Index`, `IsFirst`, `IsLast`, `Parent`, `PreferredWidth`, `PreferredWidthType`, `Shading`, `Width` |
| Done (step 8a) | `NestingLevel`, `Next`, `Previous` |
| Done (step 8b) | `SetWidth` |
| Planned, step 8 | `Sort` |
| Needs layout | `AutoFit` |
| Planned, step 19 | `Select` |

<!-- learn: cells -->

**Cells.**

| Status | Members |
|---|---|
| Done | `Item`, `Count`, `NestingLevel`, `Parent` |
| Done (step 8a) | `SetHeight`, `Height`, `HeightRule`, `PreferredWidth` (in points it also sets the columns the cells are all of), `PreferredWidthType`, `Shading` (also the table's when the cells are all of them), `VerticalAlignment` |
| Done (step 8b) | `DistributeWidth`, `SetWidth`, `Width` |
| Done (step 8c) | `Borders` |
| Done (step 8d) | `Add`, `Delete`, `Merge`, `Split` |
| Needs layout | `AutoFit`, `DistributeHeight` (Word measures the laid-out height) |

`Row.SetLeftIndent(indent, ruler style)` leaves space before the first cell of one row, as Word writes it: `gridBefore` and `wBefore` on the row, a grid column where the row starts, and `gridAfter` on the rows that now end earlier. With a ruler style other than `None` the cells give up the indent so that the row ends where it did (`FirstColumn`: the first cell; `Proportional`: in proportion; `SameWidth`: every cell the same whole number of twips). `Row.LeftIndent` reads the space before the row (not the table indent) and sets it as `RulerStyle.None`. A grid column is never made of an edge that is within 2 twips of another, as Word does. `Cell.Formula(formula, numFormat)` puts a formula field at the start of the cell, written in pieces as Word writes it; the result is worked out for `SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX` and `PRODUCT` of `ABOVE`, `BELOW`, `LEFT` or `RIGHT`, which take the numbers up to the first cell that is empty or has no number, and `\#` pictures follow .NET custom formats. Other formulas (cell references, arithmetic) are left for Word to work out. `AutoSum` is the sum of the numbers above or, when there are none directly above, to the left; with neither, Word's result `!Unexpected End of Formula`.

`Table.ConvertToText`, `Rows.ConvertToText` and `Row.ConvertToText` follow Word 16: each cell paragraph is a paragraph for `Paragraphs`; otherwise each row is a paragraph with its cells' text divided by the separator (a tab, a comma and a space, a hyphen for `Default`, or the text given), with left tab stops at the edge of each cell after the first, and a left indent of the cell margin, when tabs separate. A row split out of a table leaves the table in two around the paragraph. Cells with several paragraphs are joined with a space (not checked in Word), and tables inside cells are not converted (the library throws). `Range.ConvertToTable` makes a row of each paragraph and a cell of each field, or fills rows of `NumColumns` cells with paragraphs; the new table is the one `Tables.Add` makes, and the cells hold the text only: the formatting of the text is not kept. `Default` as the separator for `ConvertToTable` is assumed to be the hyphen too. `Table.Sort(excludeHeader, fieldNumber, type, order)` sorts the rows by one column (0-based), keeping equal rows in order and the header row if asked; text is compared by the current culture without regard to case, numbers by the first number in the text, and dates; a table with merged cells throws.

`Cells.Merge` merges the rectangle that holds all the cells. `Cells.Add` adds a cell to a row, before a cell or at the end, like the cell it goes before (or the last), and the grid grows by its width with the other rows left short (`gridAfter`). `Cells.Delete` with `ShiftCells.Left` deletes the columns when the cells are all of them; with `Up` the cells below move up and the bottom one is left empty (also `Cell.Delete(ShiftCells.Up)`); a cell that is merged vertically or spans columns throws. `Cells.Split(rows, columns, mergeBeforeSplit)` splits each cell; whole rows split into rows are new rows, which Word gives the row height it measured: the library writes 252 twips, one line of the default 11 point font. Differences: Word turns a merge of a whole table into one row and one column with a measured height, where the library keeps the merged cell in the table's grid and rows; and `mergeBeforeSplit` can only split into a number of columns that divides the columns the merged cell spans, as `Cell.Split` does. The Word scenarios cover each operation on a row or a 2 by 2 block.

`Cells.Borders`, `Rows.Borders` and `Columns.Borders` work on the block of cells: the edges of the block are the outside borders and the lines between its cells the inside ones. From Word 16: a line on the edge of the block is also written as the opposite border of the cell on its other side, and when the block is a whole table the borders go on the table (`tblBorders`) and nothing is written on the cells. The reference scenarios cover a row and a whole table; a block that is a part of the rows and columns of a table follows the same rule but has not been compared with Word.

A cell given a width of its own makes the table ragged, as in Word: the grid gets a column wherever a cell of any row ends, each cell spans the grid columns between its ends (`gridSpan`), and a row that ends before the table does records the unused columns as `gridAfter` and `wAfter`. `SetWidth` with a ruler style also gives the cells after it the difference (`FirstColumn`: the next cell; `Proportional`: all of them in proportion; `SameWidth`: all of them evenly); with nothing after it, as with `RulerStyle.None`, nothing else changes. `DistributeWidth` shares the width of the cells of each row evenly, in whole twips with the remainder in the last cell. All of this is from Word 16. For `Cells.SetWidth` with a ruler style other than `None` the library applies the style to the cells after the last cell of each row; Word has not been checked for it.

<!-- learn: cell -->

**Cell.**

| Status | Members |
|---|---|
| Done | `Delete` (left, entire row and column; up since step 8d; vertically merged cells still throw), `Merge`, `Split`, `Borders`, `BottomPadding`, `Column`, `ColumnIndex`, `FitText`, `Height`, `HeightRule`, `LeftPadding`, `Next`, `Parent`, `PreferredWidth`, `PreferredWidthType`, `Previous`, `Range`, `RightPadding`, `Row`, `RowIndex`, `Shading`, `Tables`, `TopPadding`, `VerticalAlignment`, `Width`, `WordWrap` |
| Done (step 8a) | `SetHeight`, `NestingLevel` |
| Done (step 8b) | `SetWidth` |
| Done (step 8f) | `AutoSum`, `Formula` |
| Planned, step 19 | `Select` |
| Decided against | `ID` (an HTML id) |

### 3.12 Bookmarks and Comments
<!-- learn: bookmark -->

**Bookmark.**

| Status | Members |
|---|---|
| Done | `Column` (read only: it reads the columns of a bookmark Word made over table cells; `Bookmarks.Add` over cells still writes an ordinary bookmark, where Word writes `colFirst`/`colLast` and puts the end after the row), `Copy`, `Delete`, `Empty`, `End`, `Name`, `Parent`, `Range`, `Start` |
| Done (step 11) | `StoryType` |
| Planned | `Select` (19) |

<!-- learn: bookmarks -->

**Bookmarks.** `DeleteAll` is extra.

| Status | Members |
|---|---|
| Done | `Add`, `Exists`, `Item`, `Count`, `DefaultSorting`, `Parent`, `ShowHidden` |

<!-- learn: comment -->

**Comment.** `Author`, `Initial`, `Text` and `Delete` are kept although the current reference drops them.

| Status | Members |
|---|---|
| Done | `Ancestor`, `Date`, `DeleteRecursively`, `Done`, `Index`, `Parent`, `Range`, `Reference`, `Replies`, `Scope` |
| Decided against | `Contact` (an object over what `Author` and `Initial` already give) |
| Out of scope | `Edit` (opens the comment pane), `IsInk` (ink) |

<!-- learn: comments -->

**Comments.** `DeleteAll` is extra.

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count`, `Parent` |
| Out of scope | `ShowBy` (a view filter) |

### 3.13 Field and Fields
<!-- learn: field -->

**Field.**

| Status | Members |
|---|---|
| Done | `Delete`, `Unlink`, `Update`, `Code`, `Index`, `Kind`, `Locked`, `Next`, `Parent`, `Previous`, `Result`, `Type` |
| Planned | `InlineShape` (16); `Select` (19) |
| Out of scope | `Copy`, `Cut` (clipboard), `DoClick` (UI), `ShowCodes` (view state), `LinkFormat`, `OLEFormat` (OLE links) |
| Decided against | `Data` (the data of mail-merge and similar fields) |
| Decided against | `UpdateSource` (writes a result back into a linked file) |

<!-- learn: fields -->

**Fields.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Update`, `Count`, `Locked` (with `Lock` and `Unlock`), `Parent`, `Unlink` |
| Out of scope | `ToggleShowCodes` (view state) |
| Decided against | `UpdateSource` (as on `Field`) |

### 3.14 Sections, headers and footers, page setup
<!-- learn: section -->

**Section.**

| Status | Members |
|---|---|
| Done | `Borders`, `Footers`, `Headers`, `Index`, `PageSetup`, `Parent`, `ProtectedForForms`, `Range` |

<!-- learn: sections -->

**Sections.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count`, `First`, `Last`, `PageSetup`, `Parent` |

<!-- learn: headerfooter -->

**HeaderFooter.**

| Status | Members |
|---|---|
| Done | `Exists`, `Index`, `IsHeader`, `LinkToPrevious`, `PageNumbers`, `Parent`, `Range` |
| Planned, step 18 | `Shapes` |

<!-- learn: headersfooters -->

**HeadersFooters.**

| Status | Members |
|---|---|
| Done | `Item`, `Count`, `Parent` |

<!-- learn: pagesetup -->

**PageSetup.**

| Status | Members |
|---|---|
| Done | `BottomMargin`, `DifferentFirstPageHeaderFooter`, `FooterDistance`, `Gutter`, `GutterPos`, `HeaderDistance`, `LeftMargin`, `LineNumbering`, `MirrorMargins`, `OddAndEvenPagesHeaderFooter`, `Orientation`, `PageHeight`, `PageWidth`, `PaperSize`, `RightMargin`, `SectionStart`, `SuppressEndnotes`, `TextColumns`, `TopMargin`, `VerticalAlignment` |
| Done (step 10b) | `TogglePortrait`, `SectionDirection` (`SectionDirection`: `Rtl` writes `w:bidi`, `Ltr` is the default and writes nothing; also on `Document.PageSetup`) |
| Decided against | `CharsLine`, `LinesPage`, `LayoutMode`, `GutterStyle` (Word 16 derives `w:docGrid` pitch and character space from layout, and `GutterStyle` did not change what it wrote), `Parent` |
| Out of scope | Printing: `TwoPagesOnOne`, `BookFoldPrinting`, `BookFoldRevPrinting`, `BookFoldPrintingSheets`, `FirstPageTray`, `OtherPagesTray`. View: `ShowGrid`. Templates: `SetAsTemplateDefault` |

<!-- learn: pagenumbers -->

**PageNumbers.**

| Status | Members |
|---|---|
| Done | `Add`, `Count`, `Item`, `ChapterPageSeparator`, `HeadingLevelForChapter`, `IncludeChapterNumber`, `NumberStyle`, `RestartNumberingAtSection`, `ShowFirstPageNumber`, `StartingNumber` |
| Decided against | `Parent` |
| Decided against | `DoubleQuote` (Hebrew numbering only) |

<!-- learn: pagenumber -->

**PageNumber.**

| Status | Members |
|---|---|
| Done | `Alignment` (read only: setting it in Word adds another page number) |
| Planned | `Delete`, `Index`, `Parent` (10); `Select` (19) |
| Out of scope | `Copy`, `Cut` (clipboard) |

<!-- learn: linenumbering -->

**LineNumbering.**

| Status | Members |
|---|---|
| Done | `Active`, `CountBy`, `DistanceFromText`, `RestartMode`, `StartingNumber` |
| Planned, step 10 | `Parent` |

<!-- learn: textcolumns -->

**TextColumns.**

| Status | Members |
|---|---|
| Done | `Add`, `Count`, `EvenlySpaced`, `FlowDirection`, `Item`, `LineBetween`, `SetCount`, `Spacing`, `Width` |
| Decided against | `Parent` |

<!-- learn: textcolumn -->

**TextColumn.**

| Status | Members |
|---|---|
| Done | `SpaceAfter`, `Width` |
| Decided against | `Parent` |

### 3.15 Borders, Border, Shading
<!-- learn: borders -->

`Borders` exists for tables, rows, columns, cells, sections and pages. Step 7 adds it to paragraphs, characters and ranges, step 9 to styles.

**Borders.** `SetOutside`, `SetInside` and `Clear` are extra.

| Status | Members |
|---|---|
| Done | `Item`, `Count`, `Enable`, `InsideColor`, `InsideLineStyle`, `InsideLineWidth`, `OutsideColor`, `OutsideLineStyle`, `OutsideLineWidth` |
| Planned, step 7 | `InsideColorIndex`, `OutsideColorIndex`, `Shadow`, `DistanceFromBottom`, `DistanceFromLeft`, `DistanceFromRight`, `DistanceFromTop`, `Parent` |
| Planned, step 8 | `HasHorizontal`, `HasVertical` |
| Done (step 10b) | `AlwaysInFront` (false is `w:zOrder="back"`), `DistanceFrom` (`BorderDistanceFrom`: text is 0, page edge 1), `EnableFirstPageInSection`, `EnableOtherPagesInSection` (`w:display`; Word keeps at least one on: turning one off while the other is off turns the other on), `SurroundHeader`, `SurroundFooter` (the document setting `w:bordersDoNotSurroundHeader` and `Footer`). Page borders only; the others throw `NotSupportedException` |
| Decided against | `JoinBorders` (Word 16 reads it back but writes nothing), `ApplyPageBordersToAllSections` (Word 16 left no page borders in either section after it in a two-section document, so there is nothing to match) |

<!-- learn: border -->

**Border.** `Clear` is extra.

| Status | Members |
|---|---|
| Done | `Color`, `ColorIndex`, `LineStyle`, `LineWidth`, `Visible` |
| Planned | `Inside`, `Parent` (7); `ArtStyle`, `ArtWidth` (10, page art borders) |

<!-- learn: shading -->

**Shading.** `Clear` is extra.

| Status | Members |
|---|---|
| Done | `BackgroundPatternColor`, `BackgroundPatternColorIndex`, `ForegroundPatternColor`, `ForegroundPatternColorIndex`, `Texture` |
| Planned, step 7 | `Parent` |

### 3.16 TabStops, Variables, text units, StoryRanges, document properties
<!-- learn: tabstop -->

**TabStop.**

| Status | Members |
|---|---|
| Done | `Clear`, `Alignment`, `Leader`, `Position` |
| Planned, step 7 | `CustomTab`, `Next`, `Previous`, `Parent` |

<!-- learn: tabstops -->

**TabStops.**

| Status | Members |
|---|---|
| Done | `Add`, `ClearAll`, `Item`, `Count` |
| Planned, step 7 | `After`, `Before`, `Parent` |

<!-- learn: variable -->

**Variable.**

| Status | Members |
|---|---|
| Done | `Delete`, `Name`, `Value` |
| Planned, step 10 | `Index`, `Parent` |

<!-- learn: variables -->

**Variables.** `Exists` is extra.

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count` |
| Planned, step 10 | `Parent` |

<!-- learn: characters,words,sentences -->

**Characters, Words, Sentences** (all `TextUnitCollection`).

| Status | Members |
|---|---|
| Done | `Item`, `Count`, `Parent` |
| Done | `First`, `Last` (they throw when there are no items) |

<!-- learn: storyranges -->

**StoryRanges.**

| Status | Members |
|---|---|
| Done (step 11) | `Item` (by `StoryType`), `Count`, `Parent`, enumeration. Only the stories the library can range are listed: the main text, the first comment, and the headers and footers of the first section. Word joins all the comments into one story; here each comment is a story of its own, so the comments entry is the first comment. Footnote and endnote stories (and their separators) join in step 14, text boxes in step 18 |

<!-- learn: office:documentproperties -->

**DocumentProperties** (Office). `BuiltInDocumentProperties` has typed members and a name indexer; `CustomDocumentProperties` is a full collection.

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count` (the custom collection; step 10 gives the built-in one a count, enumeration and item objects) |
| Planned, step 10 | `Parent` |

<!-- learn: office:documentproperty -->

**DocumentProperty** (Office).

| Status | Members |
|---|---|
| Done | `Delete`, `Name`, `Type`, `Value` (custom properties; built-in item objects come in step 10) |
| Planned, step 10 | `Parent` |
| Decided against | `LinkSource`, `LinkToContent` (properties linked to a bookmark's text) |

## 4. Objects that do not exist yet

### 4.1 Hyperlinks (step 12)
<!-- learn: hyperlink -->

**Hyperlink.**

| Status | Members |
|---|---|
| Done (step 12) | `Delete`, `Address`, `SubAddress`, `EmailSubject`, `Name`, `Range`, `ScreenTip`, `Target`, `TextToDisplay`, `Type` (always `Range` until shapes exist), `Parent`. A hyperlink made of a HYPERLINK field, which Word counts too, can be read and deleted, not changed |
| Planned, step 18 | `Shape` |
| Out of scope | `AddToFavorites`, `Follow`, `ExtraInfoRequired` (browser and UI) |
| Decided against | `CreateNewDocument` (creates a file as a side effect) |

<!-- learn: hyperlinks -->

**Hyperlinks.**

| Status | Members |
|---|---|
| Done (step 12) | `Add` (one paragraph; a range with no text gets the text to display, or the address), `Item`, `Count`, `Parent`, enumeration. `Document.Hyperlinks` counts the main text only, as Word does; `Range.Hyperlinks` the hyperlinks in the range |

### 4.2 Lists (step 13)

`ApplyBulletDefault`, `ApplyNumberDefault` and `ApplyOutlineNumberDefault` use list definitions that match Word 365's defaults, not the user's galleries.

<!-- learn: listformat -->

**ListFormat.**

| Status | Members |
|---|---|
| Done (step 13a) | Reading: `CountNumberedItems`, `List`, `ListLevelNumber` (from 1, as Word), `ListString`, `ListTemplate`, `ListType`, `ListValue`, `Parent`, `SingleList`, `SingleListTemplate`. The values are worked out as Word works them out (counters per template, deeper levels starting again, letters that repeat, roman numbers); checked against the numbers Word gave for its default bullet, number and outline lists |
| Done (step 13b) | `ApplyBulletDefault`, `ApplyNumberDefault`, `ApplyOutlineNumberDefault`, `ApplyListTemplate(template, continuePreviousList)`, `ConvertNumbersToText`, `ListIndent`, `ListOutdent`, `RemoveNumbers`, and setting `ListLevelNumber`. The defaults are the definitions Word 16 wrote (`Templates/ListDefaults.xml`, from the `list-defaults` scenario); markup of each operation matches Word in twelve scenarios. Bullets reuse the document's bullet list; numbers go on from the paragraph before the range when it is numbered the same way and start again after a gap, as Word does. A direct left indent is moved on by the hanging indent of the first level |
| Planned, step 13c | `ApplyListTemplateWithLevel`, `CanContinuePreviousList` |
| Decided against | `ListPictureBullet` (legacy picture bullets) |

<!-- learn: list -->

**List.**

| Status | Members |
|---|---|
| Done (step 13a) | `CountNumberedItems`, `ListParagraphs`, `SingleListTemplate`, `StyleName` (empty: the library does not link lists to styles) |
| Planned, step 13c | `ApplyListTemplate`, `ApplyListTemplateWithLevel`, `CanContinuePreviousList`, `ConvertNumbersToText`, `RemoveNumbers`, `Range`, `Parent` |

<!-- learn: lists,listparagraphs,listlevels -->

**Lists, ListParagraphs, ListLevels.**

| Status | Members |
|---|---|
| Done (step 13a) | `Item` (0-based), `Count`, `Parent` for `Lists` and `ListLevels`; `Item` and `Count` for `ListParagraphs`. A list is the paragraphs that share a `w:num`, in the order they first appear. `Document.Lists` and `ListParagraphs` look at the main text |

<!-- learn: listtemplate -->

**ListTemplate.**

| Status | Members |
|---|---|
| Done (step 13a) | `ListLevels`, `Name`, `OutlineNumbered` |
| Planned, step 13c | `Convert`, `Parent` |

<!-- learn: listtemplates -->

**ListTemplates.**

| Status | Members |
|---|---|
| Done (step 13a) | `Item` (0-based), `Count`, `Parent`, enumeration |
| Planned, step 13c | `Add` |

<!-- learn: listlevel -->

**ListLevel.**

| Status | Members |
|---|---|
| Done (step 13a) | `Alignment`, `Index` (0-based), `NumberFormat`, `NumberPosition`, `NumberStyle`, `Parent`, `ResetOnHigher`, `StartAt`, `TabPosition` (read only), `TextPosition`, `TrailingCharacter`, `LinkedStyle` (read only) |
| Planned, step 13c | `Font`, setting `TabPosition` and `LinkedStyle` |
| Decided against | `ApplyPictureBullet`, `PictureBullet` (legacy picture bullets) |

### 4.3 Footnotes and endnotes (steps 11 and 14)

Step 11 laid the groundwork every new story needs: `StoryType` (which already tells footnote, endnote and separator notes and text boxes apart), `StoryRanges` and `NextStoryRange`. The story map treats every comment and every note as a story of its own, so the range of a note counts from the start of that note (Word counts one story for all the footnotes). `StoryRanges` lists the first footnote and the first endnote.

<!-- learn: footnote,endnote -->

**Footnote, Endnote.**

| Status | Members |
|---|---|
| Done (step 14a) | `Delete`, `Index` (0-based), `Parent`, `Range` (the text without the mark and the space after it; positions count from the start of the note, which is a story of its own here, where Word joins the notes of a kind), `Reference` |

<!-- learn: footnotes -->

**Footnotes.**

| Status | Members |
|---|---|
| Done (step 14a) | `Add` (the mark goes at the end of the range; a mark of your own is text), `Item` (0-based), `Count`, `Parent`, enumeration. `Document.Footnotes` looks at the main text; `Range.Footnotes` at the marks in the range. Notes are numbered 1, 2, 3 in the order of the text, as Word saves them |
| Done (step 14b) | `Convert`, `SwapWithEndnotes`, `Location`, `NumberingRule`, `NumberStyle`, `StartingNumber`, `Separator`, `ContinuationSeparator`. The options are document-wide: written in the settings and in every section, as Word does. A starting number other than 1 needs continuous numbering, and the other way round, as in Word |
| Decided against | `ContinuationNotice`, `ResetContinuationNotice`, `ResetContinuationSeparator`, `ResetSeparator` (Word leaves the notice out of the file until it is set, and the library never changes the separators) |

<!-- learn: endnotes -->

**Endnotes.**

| Status | Members |
|---|---|
| Done (step 14a) | `Add`, `Item` (0-based), `Count`, `Parent`, enumeration, as for footnotes |
| Done (step 14b) | `Convert`, `SwapWithFootnotes`, `Location` (the end of the document is the default), `NumberingRule`, `NumberStyle` (lowercase roman numbers are the default), `StartingNumber`, `Separator`, `ContinuationSeparator` |
| Decided against | `ContinuationNotice`, `ResetContinuationNotice`, `ResetContinuationSeparator`, `ResetSeparator`, as for footnotes |

<!-- learn: footnoteoptions -->

**FootnoteOptions.**

| Status | Members |
|---|---|
| Done (step 14b) | `Location`, `NumberingRule`, `NumberStyle`, `Parent`, `StartingNumber` (`Range.FootnoteOptions`: the options of the sections the range is in, with the document's where a section has none) |
| Decided against | `LayoutColumns` (the columns footnotes are laid out in need a layout of the page) |

<!-- learn: endnoteoptions -->

**EndnoteOptions.**

| Status | Members |
|---|---|
| Done (step 14b) | `Location` (a setting of the whole document), `NumberingRule`, `NumberStyle`, `Parent`, `StartingNumber` |

### 4.4 Content controls and custom XML parts (step 15)
<!-- learn: contentcontrol -->

**ContentControl.**

| Status | Members |
|---|---|
| Done (step 15a) | `Appearance`, `Delete(deleteContents)`, `ID`, `LockContentControl`, `LockContents`, `ParentContentControl`, `Range`, `ShowingPlaceholderText`, `Tag`, `Title`, `Type`. A control made at a point shows its placeholder (kept in a glossary document as Word keeps it); around text it holds the text. Rich text, text, combo box, drop-down list, date and check box can be made |
| Done (step 15b) | `Checked`, `DateCalendarType`, `DateDisplayFormat`, `DateDisplayLocale`, `DateStorageFormat`, `DropdownListEntries`, `MultiLine`, `SetCheckedSymbol`, `SetUncheckedSymbol`, `Color`, `Temporary`, `Level` (read only), `DefaultTextStyle` (the name of a character style). `Checked`, the symbols and the entries re-show the text the control displays, as Word does. A member that belongs to another kind of control throws `InvalidOperationException` |
| Done (step 15c) | `XMLMapping` (see below) |
| Planned later | `SetPlaceholderText`, `PlaceholderText`, `Parent`. Word's `SetPlaceholderText` could not be called from PowerShell to measure what it writes (it rejects a missing building block), so it waits for a scenario that can |
| Planned later | `Ungroup`, `AllowInsertDeleteSection`, `BuildingBlockCategory`, `BuildingBlockType`, `RepeatingSectionItems`, `RepeatingSectionItemTitle` (picture, group, building block and repeating section controls) |
| Out of scope | `Copy`, `Cut` (clipboard) |

<!-- learn: contentcontrols -->

**ContentControls.**

| Status | Members |
|---|---|
| Done (step 15a) | `Add` (around the runs of one paragraph, or at a point), `Item` (0-based), `Count`, `Parent`, enumeration. `Document.ContentControls` and `Range.ContentControls`; a control inside another counts too |

<!-- learn: contentcontrollistentries -->

**ContentControlListEntries.**

| Status | Members |
|---|---|
| Done (step 15b) | `Add(text, value, index)` (the value is the text when left out; `index` is 0-based), `Clear`, `Item` (0-based), `Count`, enumeration |
| Planned later | `Parent` |

<!-- learn: contentcontrollistentry -->

**ContentControlListEntry.**

| Status | Members |
|---|---|
| Done (step 15b) | `Delete`, `MoveDown`, `MoveUp`, `Index` (0-based), `Text`, `Value`, `Select` (the control then holds the entry's text and no longer shows its placeholder) |
| Planned later | `Parent` |

<!-- learn: xmlmapping -->

**XMLMapping.**

| Status | Members |
|---|---|
| Done (step 15c) | `Delete`, `SetMapping(xPath, prefixMappings, source)` (the XPath is evaluated with the prefixes; false when it matches no element; the control then shows the element's text), `CustomXMLPart`, `IsMapped`, `PrefixMappings`, `XPath` |

<!-- learn: office:customxmlparts -->

**CustomXMLParts** (Office).

| Status | Members |
|---|---|
| Done (step 15c) | `Add(xml)`, `Item` (0-based), `SelectByID`, `SelectByNamespace`, `Count`, enumeration. `Document.CustomXMLParts` |
| Out of scope | The events `PartAfterAdd`, `PartAfterLoad`, `PartBeforeDelete` |

<!-- learn: office:customxmlpart -->

**CustomXMLPart** (Office).

| Status | Members |
|---|---|
| Done (step 15c) | `Delete` (bound controls keep their text), `Load`, `LoadXML`, `BuiltIn` (true for bibliography sources), `Id`, `NamespaceURI`, `XML` |
| Out of scope | The events `NodeAfterDelete`, `NodeAfterInsert`, `NodeAfterReplace` |
| Decided against | `SchemaCollection`, `Errors` (schema validation of the part) |

### 4.5 Inline shapes (step 16)
<!-- learn: inlineshape -->

**InlineShape.**

| Status | Members |
|---|---|
| Done (step 16) | `AlternativeText`, `Delete` (a picture file nothing uses any more leaves the package), `Height`, `LockAspectRatio`, `Range`, `ScaleHeight`, `ScaleWidth`, `Title`, `Type` (picture, linked picture, chart, locked canvas, SmartArt), `Width`. Sizes are in points and come from the file (PNG `pHYs`, JPEG JFIF density, BMP resolution; 96 dpi otherwise), as Word measures them |
| Planned later | `Borders`, `Field`, `Hyperlink`, `Parent`, `PictureFormat` (crop, brightness and the like need picture formatting) |
| Decided against | `Reset`: in Word 16 it left the size of a resized picture alone (it resets picture formatting, which the library does not model) |
| Planned, steps 18 and 19 | `ConvertToShape` (18), `Select` (19) |
| Out of scope | Charts and SmartArt: `Chart`, `HasChart`, `SmartArt`, `HasSmartArt`. OLE: `OLEFormat`, `LinkFormat`. HTML: `Script` |
| Decided against | Picture effects, each an object of its own: `Fill`, `Glow`, `Line`, `Reflection`, `Shadow`, `SoftEdge`, `TextEffect`. `GroupItems`, `HorizontalLineFormat`, `IsPictureBullet` (legacy VML) |

<!-- learn: inlineshapes -->

**InlineShapes.**

| Status | Members |
|---|---|
| Done (step 16) | `AddPicture(fileName, linkToFile, saveWithDocument, range)` (PNG, JPEG, GIF and BMP; the picture goes at the start of the range and the range's text stays; a file already in the document is shared; a linked picture is only linked), `Item` (0-based), `Count`, `Parent`, enumeration. `Document.InlineShapes` and `Range.InlineShapes`; a header or footer picture goes in its own part |
| Out of scope | `AddChart2`, `AddSmartArt` (charts and SmartArt), `AddOLEControl`, `AddOLEObject` (OLE) |
| Decided against | `AddHorizontalLine`, `AddHorizontalLineStandard`, `AddPictureBullet` (legacy VML), `AddWebVideo`, `New` |

### 4.6 Tables of contents and figures (step 17)

A table of contents is a `TOC` field. `Update` rebuilds its entries from the headings or `TC` fields; page numbers need layout, so they are left for Word to fill in (the field stays dirty).

<!-- learn: tableofcontents -->

**TableOfContents.**

| Status | Members |
|---|---|
| Planned, step 17 | `Delete`, `Update`, `HeadingStyles`, `HidePageNumbersInWeb`, `IncludePageNumbers`, `LowerHeadingLevel`, `Parent`, `Range`, `RightAlignPageNumbers`, `TabLeader`, `TableID`, `UpperHeadingLevel`, `UseFields`, `UseHeadingStyles`, `UseHyperlinks` |
| Out of scope | `UpdatePageNumbers` (layout) |

<!-- learn: tablesofcontents,tablesoffigures -->

**TablesOfContents, TablesOfFigures.** A table of figures is the same field with a caption label; its item object follows `TableOfContents`.

| Status | Members |
|---|---|
| Planned, step 17 | `Add`, `Item`, `MarkEntry`, `Count`, `Format`, `Parent` |

### 4.7 Floating shapes (step 18)
<!-- learn: shape -->

**Shape.**

| Status | Members |
|---|---|
| Planned, step 18 | `ConvertToInlineShape`, `Delete`, `Duplicate`, `Flip`, `IncrementLeft`, `IncrementRotation`, `IncrementTop`, `ScaleHeight`, `ScaleWidth`, `Ungroup`, `ZOrder`, `AlternativeText`, `Anchor`, `AutoShapeType`, `Child`, `Decorative`, `Fill`, `GroupItems`, `Height`, `HeightRelative`, `HorizontalFlip`, `Hyperlink`, `ID`, `LayoutInCell`, `Left`, `LeftRelative`, `Line`, `LockAnchor`, `LockAspectRatio`, `Name`, `Parent`, `ParentGroup`, `PictureFormat`, `RelativeHorizontalPosition`, `RelativeHorizontalSize`, `RelativeVerticalPosition`, `RelativeVerticalSize`, `Rotation`, `TextFrame`, `Title`, `Top`, `TopRelative`, `Type`, `VerticalFlip`, `Visible`, `Width`, `WidthRelative`, `WrapFormat`, `ZOrderPosition` |
| Planned, step 19 | `Select` |
| Out of scope | Charts, SmartArt, OLE and 3D: `Chart`, `HasChart`, `SmartArt`, `HasSmartArt`, `OLEFormat`, `LinkFormat`, `Model3D`. HTML: `Script`. Theme-based styles: `ShapeStyle`, `BackgroundStyle`, `GraphicStyle`. The format painter: `Apply`, `PickUp` |
| Decided against | Drawing canvases: `CanvasItems`, `CanvasCropBottom`, `CanvasCropLeft`, `CanvasCropRight`, `CanvasCropTop`. Geometry editing: `Adjustments`, `Nodes`, `Vertices`, `Callout`. Effects: `Glow`, `Reflection`, `Shadow`, `SoftEdge`, `ThreeD`, `TextEffect`. `TextFrame2` (the Office text frame, which TextFrame covers), `SetShapesDefaultProperties` |

<!-- learn: shapes -->

**Shapes.**

| Status | Members |
|---|---|
| Planned, step 18 | `AddLabel`, `AddLine`, `AddPicture`, `AddShape`, `AddTextbox`, `Item`, `Count`, `Parent` |
| Planned, step 19 | `SelectAll` |
| Out of scope | `AddChart2`, `AddSmartArt`, `Add3DModel` (charts, SmartArt, 3D), `AddOLEControl`, `AddOLEObject` (OLE) |
| Decided against | `AddCallout`, `AddCanvas`, `AddCurve`, `AddPolyline`, `BuildFreeform`, `AddTextEffect` (WordArt), `AddWebVideo`, `Range` (multi-shape ranges) |

<!-- learn: textframe -->

**TextFrame.**

| Status | Members |
|---|---|
| Planned, step 18 | `DeleteText`, `AutoSize`, `ContainingRange`, `HasText`, `HorizontalAnchor`, `MarginBottom`, `MarginLeft`, `MarginRight`, `MarginTop`, `NoTextRotation`, `Orientation`, `Parent`, `TextRange`, `VerticalAnchor`, `WordWrap` |
| Out of scope | `Overflowing` (layout) |
| Decided against | Linked text boxes: `BreakForwardLink`, `ValidLinkTarget`, `Next`, `Previous`. `Column`, and the WordArt formats `PathFormat`, `WarpFormat`, `ThreeD` |

<!-- learn: wrapformat -->

**WrapFormat.**

| Status | Members |
|---|---|
| Planned, step 18 | `AllowOverlap`, `DistanceBottom`, `DistanceLeft`, `DistanceRight`, `DistanceTop`, `Parent`, `Side`, `Type` |

### 4.8 Selection (step 19)
<!-- learn: selection-range -->

A logical selection: a `RangeWrapper` plus insertion-point state, one per document, built on the span machinery. It has no view, so movement by lines, screens and pages stays unsupported. A member that `Selection` shares with `Range` has the status given for `Range` in 3.1, with step 19 for those `Range` already has. The members below are the ones only `Selection` has.

| Status | Members |
|---|---|
| Planned, step 19 | `TypeText`, `TypeParagraph`, `TypeBackspace`, `HomeKey`, `EndKey` (story, row and cell units; line units are layout), `MoveLeft`, `MoveRight`, `MoveUp`, `MoveDown` (paragraph units; line and screen units are layout), `Extend`, `ExtendMode`, `Shrink`, `StartIsActive`, `Range`, `Type`, `HeaderFooter`, `BoldRun`, `ItalicRun`, `LtrPara`, `RtlPara`, `LtrRun`, `RtlRun`, `ClearFormatting`, `ClearCharacterAllFormatting`, `ClearCharacterDirectFormatting`, `ClearCharacterStyle`, `ClearParagraphAllFormatting`, `ClearParagraphDirectFormatting`, `ClearParagraphStyle`, `InsertCells`, `InsertColumns`, `InsertColumnsRight`, `InsertRows`, `InsertRowsAbove`, `InsertRowsBelow`, `InsertFormula`, `InsertNewPage`, `InsertStyleSeparator`, `SplitTable`, `SelectCell`, `SelectColumn`, `SelectRow`, `SelectCurrentAlignment`, `SelectCurrentColor`, `SelectCurrentFont`, `SelectCurrentIndent`, `SelectCurrentSpacing`, `SelectCurrentTabs`, `NextField`, `PreviousField`, `CreateTextbox` |
| Out of scope | Layout: `IPAtEndOfLine`. UI: `Active`, `ColumnSelectMode`, `EscapeKey`, `ReadingModeGrowFont`, `ReadingModeShrinkFont`. Clipboard: `CopyFormat`, `PasteFormat`. Templates: `CreateAutoTextEntry`. Track changes: `NextRevision`, `PreviousRevision` |
| Decided against | `Flags`, `ShrinkDiscontiguousSelection` (there are no discontiguous selections), `ToggleCharacterCode`, `ChildShapeRange`, `HasChildShapeRange` (multi-shape ranges) |

### 4.9 Application and Documents (step 20)
<!-- learn: application -->

A light facade so that ported code can start from `Application`, `Documents` or `ActiveDocument`. It owns no process and no UI.

**Application.**

| Status | Members |
|---|---|
| Planned, step 20 | `ActiveDocument`, `Documents`, `Name`, `Version`, `Parent`, `Quit`, `CleanString`, `UserName`, `UserInitials`, `UserAddress`, `CentimetersToPoints`, `InchesToPoints`, `LinesToPoints`, `MillimetersToPoints`, `PicasToPoints`, `PointsToCentimeters`, `PointsToInches`, `PointsToLines`, `PointsToMillimeters`, `PointsToPicas` |
| Planned, step 19 | `Selection` |
| Out of scope | Everything else is Word the program: its windows and UI, add-ins, templates, printing, proofing, mail and mail merge, DDE, VBA, and events: `Activate`, `ActiveEncryptionSession`, `ActivePrinter`, `ActiveProtectedViewWindow`, `ActiveWindow`, `AddAddress`, `AddIns`, `ArbitraryXMLSupportAvailable`, `Assistance`, `AutoCaptions`, `AutoCorrect`, `AutoCorrectEmail`, `AutomaticChange`, `AutomationSecurity`, `BackgroundPrintingStatus`, `BackgroundSavingStatus`, `Bibliography`, `BrowseExtraFileTypes`, `Browser`, `BuildKeyCode`, `COMAddIns`, `CapsLock`, `Caption`, `CaptionLabels`, `ChangeFileOpenDirectory`, `ChartDataPointTrack`, `CheckGrammar`, `CheckLanguage`, `CheckSpelling`, `CommandBars`, `CompareDocuments`, `CustomDictionaries`, `CustomizationContext`, `DDEExecute`, `DDEInitiate`, `DDEPoke`, `DDERequest`, `DDETerminate`, `DDETerminateAll`, `DefaultLegalBlackline`, `DefaultSaveFormat`, `DefaultTableSeparator`, `DefaultWebOptions`, `Dialogs`, `DisplayAlerts`, `DisplayAutoCompleteTips`, `DisplayDocumentInformationPanel`, `DisplayRecentFiles`, `DisplayScreenTips`, `DisplayScrollBars`, `DocumentBeforeClose`, `DocumentBeforePrint`, `DocumentBeforeSave`, `DocumentChange`, `DocumentOpen`, `DocumentSync`, `DontResetInsertionPointProperties`, `EPostageInsert`, `EPostageInsertEx`, `EPostagePropertyDialog`, `EmailOptions`, `EmailTemplate`, `EnableCancelKey`, `FeatureInstall`, `FileConverters`, `FileDialog`, `FileValidation`, `FindKey`, `FocusInMailHeader`, `FontNames`, `GetAddress`, `GetDefaultTheme`, `GetSpellingSuggestions`, `GoBack`, `GoForward`, `HangulHanjaDictionaries`, `Height`, `Help`, `HelpTool`, `International`, `IsObjectValid`, `IsSandboxed`, `KeyBindings`, `KeyString`, `Keyboard`, `KeyboardBidi`, `KeyboardLatin`, `KeysBoundTo`, `LandscapeFontNames`, `Language`, `LanguageSettings`, `Languages`, `Left`, `ListCommands`, `ListGalleries`, `LoadMasterList`, `LookupNameProperties`, `MAPIAvailable`, `MacroContainer`, `MailMergeAfterMerge`, `MailMergeAfterRecordMerge`, `MailMergeBeforeMerge`, `MailMergeBeforeRecordMerge`, `MailMergeDataSourceLoad`, `MailMergeDataSourceValidate`, `MailMergeDataSourceValidate2`, `MailMergeWizardSendToCustom`, `MailMergeWizardStateChange`, `MailMessage`, `MailSystem`, `MailingLabel`, `MathCoprocessorAvailable`, `MergeDocuments`, `MouseAvailable`, `Move`, `NewDocument`, `NewWindow`, `NextLetter`, `NormalTemplate`, `NumLock`, `OMathAutoCorrect`, `OnTime`, `OpenAttachmentsInFullScreen`, `Options`, `OrganizerCopy`, `OrganizerDelete`, `OrganizerRename`, `Path`, `PathSeparator`, `PickerDialog`, `PortraitFontNames`, `PrintOut`, `PrintPreview`, `ProtectedViewWindowActivate`, `ProtectedViewWindowBeforeClose`, `ProtectedViewWindowBeforeEdit`, `ProtectedViewWindowDeactivate`, `ProtectedViewWindowOpen`, `ProtectedViewWindowSize`, `ProtectedViewWindows`, `PutFocusInMailHeader`, `RecentFiles`, `Repeat`, `ResetIgnoreAll`, `Resize`, `RestrictLinkedStyles`, `Run`, `ScreenRefresh`, `ScreenUpdating`, `SensitivityLabelPolicy`, `SetDefaultTheme`, `ShowAnimation`, `ShowClipboard`, `ShowMe`, `ShowStartupDialog`, `ShowStylePreviews`, `ShowVisualBasicEditor`, `SmartArtColors`, `SmartArtLayouts`, `SmartArtQuickStyles`, `SpecialMode`, `StartupPath`, `StatusBar`, `SubstituteFont`, `SynonymInfo`, `System`, `TaskPanes`, `Tasks`, `Templates`, `ToggleKeyboard`, `Top`, `UndoRecord`, `UsableHeight`, `UsableWidth`, `UserControl`, `VBE`, `Visible`, `Width`, `WindowActivate`, `WindowBeforeDoubleClick`, `WindowBeforeRightClick`, `WindowDeactivate`, `WindowSelectionChange`, `WindowSize`, `WindowState`, `Windows`, `WordBasic`, `XMLNamespaces`, `XMLSelectionChange`, `XMLValidationError` |
| Decided against | `PixelsToPoints`, `PointsToPixels` (they depend on the screen), `Build`, `ProductCode` |

<!-- learn: documents -->

**Documents.**

| Status | Members |
|---|---|
| Planned, step 20 | `Add`, `Close`, `Item`, `Open`, `Save`, `Count`, `Parent` |
| Out of scope | `AddBlogDocument`, `CanCheckOut`, `CheckOut` (servers and blogs) |
| Decided against | `OpenNoRepairDialog` (Open without the repair dialog) |

## 5. Not Word's, limitations and decisions

### 5.1 Public surface that is not Word's

Kept on purpose, because no one or two Word members do the same thing:

* `Table.AppendRow` and `Row.Copy`: together they copy a row with its content; `Rows.Add()` adds an empty row.
* `Table.MergeTable`: Word merges tables by deleting the paragraph between them.
* `Table.Width`: a width in points for fixed and percentage widths; `PreferredWidth` is in the unit of `PreferredWidthType`.
* `Row.ClearText`, `Paragraph.DeleteBreaks`, `Paragraph.CopyPropertiesTo`, `Paragraph.InsertSectionBreak`, `Range.FirstParagraph`/`LastParagraph`, `Range.StartElement`/`EndElement`, `Bookmarks.DeleteAll`, `Comments.DeleteAll`, `Variables.Exists`, `Borders.SetOutside`/`SetInside`/`Clear`, `Style.Name`/`StyleId`, the `Direct` views of `Font` and `ParagraphFormat`, and the SDK element properties (`Paragraph.Paragraph`, `Table.Table` and so on).
* `FieldPartWrapper`, which stands for the `Range` that `Field.Code` and `Result` are in Word and has a `Range`.
* `WordColor`, `ColorHelper`, `NamedColor`, `ColorIndex`, `LineWidth` and `BorderLineWidth`: Word's constants and `RGB`, in C# form.
* `WordprocessingDocumentWrapper.Create`, `Open`, `Dispose`, `CopyAllPartsTo`, `CopyLayoutTo`, `InitializeMainDocumentPart`, `InitializeCommentsPart`: the entry points of a file library.

`OpenXmlExtensions`, `ConversionHelper`, `PositionIndex`, `StoryMap`, `Anchor`, `SpanRegistry`, `FormattingResolver` and `UnitNavigator` are internal.

### 5.2 Cannot be closed without a layout engine

`Pages` and `Lines` statistics, page and line `Information`, `GoTo` by page or line, view-dependent `Selection` movement, `Repaginate`, `PAGE`/`NUMPAGES`/`PAGEREF` results and TOC page numbers, printing, and fixed-format export. The library leaves such fields dirty so that Word updates them on open.

### 5.3 Behavioural limits of objects that exist

* Track changes: edits are never recorded as revisions, so `TrackRevisions = true` throws (reading it and setting it to false work). Tracked-deleted text is counted in positions but is not exposed as `Revisions`. Recording revisions is out of scope.
* `Find` has no sounds-like or word-forms matching. As measured in Word 16, `MatchPrefix` finds text that starts a word, `MatchSuffix` text that ends one (both together is a whole word), `IgnorePunct` lets `coop` find `co-op` and `IgnoreSpace` lets `ab` find `a b` (the match ends at its last letter). They are not applied to a wildcard search, where Word turns them off. Searching by formatting compares only the properties that are set in `Find.Font` and `Find.ParagraphFormat` with what each run and paragraph has (from direct formatting and styles, not the conditional formatting of table styles), and highlighting is matched as on or off, not by colour. `Replacement.Highlight = true` highlights in yellow, not the colour of Word's option. A match of paragraph formatting alone is each whole paragraph with its mark, as in Word.
* `Cell.Delete` and `Cells.Delete` do not shift up or left on vertically merged cells, or up on cells that span columns.
* `Range.Text` can be set for cells, paragraphs, bookmarks, headers and footers, the document and spans, but not for rows and tables.
* `Comments.Add` and `Bookmarks.Add` accept only some range kinds. Adding a table, paragraph or insert inside a hyperlink or other run-holding element, and a paragraph split inside a hyperlink, throw.
* Field results are computed for a fixed set of field types.
* Styles: names are matched in English, and `Styles[...]` for a built-in name depends on the catalogue. A document that stores localised names may not match.
* Not thread-safe per document (the span registry is, the document is not). A span outside one paragraph, cell or table builds a story map of the whole story on each operation, so a macro that loops over `Range(i, i + 1)` in a long document is slow.
* Output is validated with `OpenXmlValidator` in the tests. Step 2 adds a Word reference harness that compares generated markup with what Word 365 writes.

### 5.4 Decisions that stay

* Indexes are 0-based everywhere (README). Errors are .NET exceptions, not Word's COM error numbers.
* Unsupported behaviour throws (`NotSupportedException`, `NotImplementedException`, `InvalidOperationException`) rather than doing something else.
* When a convenient design and Word differ, the library does what Word does.

## 6. Order of work

| Step | Work | Release |
|---|---|---|
| 1 | Single-source the version in the csproj | 1.1.0 |
| 2 | Word reference harness (compare with markup that Word 365 writes) | 1.1.0 |
| 3 to 5 | Make generated markup match Word 365 | 1.1.0 |
| 6 | `Range` (3.1), `Characters`/`Words`/`Sentences` | 1.2.0 |
| 7 | `Font`, `ParagraphFormat`, `Paragraph`, `Paragraphs`, borders and shading of paragraphs and characters, `TabStops` | 1.2.0 |
| 8 | Tables: `Rows`, `Columns` and `Cells` as formatting collections, sort and convert, the `hMerge` pass | 1.2.0 |
| 9 | `Find`/`Replacement` formatting, `Style`, `TableStyle`, style setters (C2) | 1.2.0 |
| 10 | Document, sections, page setup, fields, comments, variables and document properties | 1.2.0 |
| 11 | Story groundwork: `StoryType`, footnote, endnote and text-box stories, `StoryRanges` | 1.3.0 |
| 12 | Hyperlinks | 1.3.0 |
| 13 | Lists and `ListFormat` | 1.3.0 |
| 14 | Footnotes and endnotes | 1.3.0 |
| 15 | Custom XML parts and content controls | 1.4.0 |
| 16 | Inline shapes | 1.4.0 |
| 17 | Tables of contents and figures | 1.4.0 |
| 18 | Floating shapes | 1.5.0 |
| 19 | A logical `Selection` | 1.5.0 |
| 20 | `Application`/`Documents` facade, README aliases (C6) | 1.5.0 |

Every step is additive. An Excel counterpart is assessed in [excel-assessment.md](excel-assessment.md) and is built only after the Word work.
