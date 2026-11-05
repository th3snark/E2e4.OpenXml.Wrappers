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
| Planned, step 8 | `ConvertToTable`, `Sort`, `SortAscending`, `SortDescending`, `Columns`, `TopLevelTables`, `IsEndOfRowMark`, `Orientation` |
| Planned, step 9 | `CharacterStyle`, `ParagraphStyle`, `TableStyle` |
| Planned, step 10 | `PageSetup` |
| Planned, step 11 | `StoryType`, `NextStoryRange` |
| Planned, steps 12 to 19 | `Hyperlinks` (12); `ListFormat`, `ListParagraphs`, `ListStyle`, `SetListLevel` (13); `Footnotes`, `Endnotes`, `FootnoteOptions`, `EndnoteOptions` (14); `ContentControls`, `ParentContentControl` (15); `InlineShapes` (16); `ShapeRange` (18); `Select` (19) |
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
| Planned, step 9 | `DefaultTableStyle`, `SetDefaultTableStyle` |
| Planned, step 10 | `Saved`, `SaveFormat`, `Type`, `AttachedTemplate`, `CopyStylesFromTemplate`, `UpdateStyles`, `UpdateStylesOnOpen`, `DefaultTabStop`, `AutoHyphenation`, `HyphenateCaps`, `HyphenationZone`, `ConsecutiveHyphensLimit`, `ReadOnlyRecommended`, `WritePassword`, `WriteReserved`, `EmbedTrueTypeFonts`, `SaveSubsetFonts`, `DoNotEmbedSystemFonts`, `RemovePersonalInformation`, `RemoveDateAndTime`, `RemoveDocumentInformation`, `DeleteAllComments`, `Compatibility`, `SetCompatibilityMode`, `Convert`, `Final` |
| Planned, steps 11 to 20 | `StoryRanges` (11); `Hyperlinks` (12); `Lists`, `ListParagraphs`, `ListTemplates`, `ConvertNumbersToText`, `CountNumberedItems`, `RemoveNumbers` (13); `Footnotes`, `Endnotes` (14); `ContentControls`, `CustomXMLParts`, `SelectContentControlsByTag`, `SelectContentControlsByTitle`, `SelectLinkedControls`, `SelectUnlinkedControls` (15); `InlineShapes` (16); `TablesOfContents`, `TablesOfFigures` (17); `Shapes`, `Background` (18); `Select` (19); `Parent`, `Activate` (20) |
| Out of scope | Layout: `Repaginate`, `FitToPages`, `FreezeLayout`, `ReadingModeLayoutFrozen`, `ConvertAutoHyphens`, `ManualHyphenation`, `ExportAsFixedFormat`, `ExportAsFixedFormat2`, `ExportAsFixedFormat3`. Printing: `PrintOut`, `PrintPreview`, `ClosePrintPreview`, `PrintFormsData`, `PrintPostScriptOverText`, `PrintRevisions`. UI and windows: `ActiveWindow`, `Windows`, `Undo`, `Redo`, `UndoClear`, `ViewCode`, `ViewPropertyBrowser`, `WebPagePreview`, `ToggleFormsDesign`, `FormsDesign`, `ReturnToLastReadPosition`, `ReadingLayoutSizeX`, `ReadingLayoutSizeY`, `PresentIt`, `Broadcast`, `CommandBars`, `FormattingShowClear`, `FormattingShowFilter`, `FormattingShowFont`, `FormattingShowNextLevel`, `FormattingShowNumbering`, `FormattingShowParagraph`, `FormattingShowUserStyleName`, `ClickAndTypeParagraphStyle`, `StyleSortMethod`, `DeleteAllCommentsShown`, `FollowHyperlink`, `DisableFeatures`, `DisableFeaturesIntroducedAfter`, `UserControl`, `CheckConsistency`, `DocumentInspectors`. Proofing: `CheckGrammar`, `CheckSpelling`, `DetectLanguage`, `LanguageDetected`, `GrammarChecked`, `SpellingChecked`, `GrammaticalErrors`, `SpellingErrors`, `ShowGrammaticalErrors`, `ShowSpellingErrors`, `ReadabilityStatistics`, `EmbedLinguisticData`, `ActiveWritingStyle`. Track changes and review: `Revisions`, `AcceptAllRevisions`, `AcceptAllRevisionsShown`, `RejectAllRevisions`, `RejectAllRevisionsShown`, `TrackFormatting`, `TrackMoves`, `Compare`, `Merge`, `EndReview`, `ReplyWithChanges`, `SendForReview`, `OriginalDocumentTitle`, `RevisedDocumentTitle`, `CoAuthoring`. Servers, mail and the host: `AddToFavorites`, `CanCheckin`, `CheckIn`, `CheckInWithVersion`, `LockServerFile`, `GetWorkflowTasks`, `GetWorkflowTemplates`, `Post`, `SendFax`, `SendFaxOverInternet`, `SendMail`, `Email`, `MailEnvelope`, `ServerPolicy`, `DocumentLibraryVersions`, `ContentTypeProperties`, `Sync`, `Permission`, `Signatures`, `SensitivityLabel`, `Reload`, `ReloadAs`, `IsInAutosave`, `AutoSaveOn`, `Container`, `Research`, `SmartDocument`. Encryption (the Open XML SDK does not encrypt packages): `HasPassword`, `Password`, `EncryptionProvider`, `PasswordEncryptionAlgorithm`, `PasswordEncryptionFileProperties`, `PasswordEncryptionKeyLength`, `PasswordEncryptionProvider`, `SetPasswordEncryptionOptions`. Mail merge, letters and legacy forms: `MailMerge`, `Envelope`, `DataForm`, `CreateLetterContent`, `GetLetterContent`, `SetLetterContent`, `RunLetterWizard`, `FormFields`, `ResetFormFields`, `SaveFormsData`. VBA and events: `RunAutoMacro`, `HasVBProject`, `VBProject`, `VBASigned`, `CodeName`, and the events `BuildingBlockInsert`, `ContentControlAfterAdd`, `ContentControlBeforeContentUpdate`, `ContentControlBeforeDelete`, `ContentControlBeforeStoreUpdate`, `ContentControlOnEnter`, `ContentControlOnExit`, `New`, `Open`, `XMLAfterInsert`, `XMLBeforeDelete` (and the Close and Sync events). Themes: `ApplyDocumentTheme`, `ApplyTheme`, `RemoveTheme`, `ActiveTheme`, `ActiveThemeDisplayName`, `DocumentTheme`, `LockTheme`. Excluded objects: `Frames`, `Subdocuments`, `IsMasterDocument`, `IsSubdocument`, `Indexes`, `TablesOfAuthorities`, `TablesOfAuthoritiesCategories`, `OMaths`, `OMathBreakBin`, `OMathBreakSub`, `OMathFontName`, `OMathIntSubSupLim`, `OMathJc`, `OMathLeftMargin`, `OMathNarySupSubLim`, `OMathRightMargin`, `OMathSmallFrac`, `OMathWrap`, `UseMathDefaults`, `ChartDataPointTrack`, `DeleteAllInkAnnotations`. HTML, text and XML formats (the library writes packages only): `HTMLDivisions`, `Frameset`, `DefaultTargetFrame`, `WebOptions`, `Scripts`, `StyleSheets`, `OpenEncoding`, `SaveEncoding`, `TextEncoding`, `TextLineEnding`, `OptimizeForWord97`, `TransformDocument`, `XMLSaveThroughXSLT`, `XMLUseXSLTWhenSaving`, `XMLShowAdvancedErrors`, `XMLSchemaReferences`, `SelectNodes`, `SelectSingleNode` (custom XML markup, which Word no longer supports). Templates: `MakeCompatibilityDefault` |
| Decided against | AutoFormat and style galleries: `AutoFormat`, `AutoFormatOverride`, `Kind`, `ApplyQuickStyleSet2`, `SaveAsQuickStyleSet`, `LockQuickStyleSet`. Formatting restrictions and editing permissions (document protection covers editing): `EnforceStyle`, `RemoveLockedStyles`, `DeleteAllEditableRanges`, `SelectAllEditableRanges`. Legacy conversions: `DowngradeDocument`, `ConvertVietDoc`. `Bibliography` (rare), `CurrentRsid` (internal to Word). Drawing grid settings: `GridDistanceHorizontal`, `GridDistanceVertical`, `GridOriginFromMargin`, `GridOriginHorizontal`, `GridOriginVertical`, `GridSpaceBetweenHorizontalLines`, `GridSpaceBetweenVerticalLines`, `SnapToGrid`, `SnapToShapes`. East Asian line breaking: `JustificationMode`, `KerningByAlgorithm`, `FarEastLineBreakLanguage`, `FarEastLineBreakLevel`, `NoLineBreakAfter`, `NoLineBreakBefore` |

### 3.3 Font
<!-- learn: font -->

Reads give effective formatting, `Direct` gives direct formatting, and a mixed value reads as null.

| Status | Members |
|---|---|
| Done | `Reset`, `AllCaps`, `Bold`, `ColorIndex`, `DoubleStrikeThrough`, `Emboss`, `Engrave`, `Hidden`, `Italic`, `Kerning`, `Name`, `Outline`, `Position`, `Scaling`, `Shadow`, `Size`, `SmallCaps`, `Spacing`, `StrikeThrough`, `Subscript`, `Superscript`, `Underline`, `UnderlineColor`, and from step 7: `Grow` and `Shrink` (Word's sizes; the complex script size follows), `Duplicate` (also `Range.Font = font`), `BoldBi`, `ItalicBi`, `SizeBi`, `ColorIndexBi` (the same as `ColorIndex`: Word keeps one colour), `NameAscii`, `NameOther`, `NameFarEast`, `NameBi` (theme fonts read `+Body`, `+Headings` and `+Body CS`), `EmphasisMark`, `DisableCharacterSpaceGrid`. `Name` also changes the complex script font when that was set on its own, as Word does |
| Planned, step 7 | `Parent`, `TextColor`, `Borders`, `Shading`, and the OpenType features stored as single `w14` elements: `Ligatures`, `NumberForm`, `NumberSpacing`, `StylisticSet`, `ContextualAlternates` |
| Out of scope | `SetAsTemplateDefault` (writes the attached template) |
| Decided against | The 2010 text effects, each an object of its own in `w14` markup: `Fill`, `Glow`, `Line`, `Reflection`, `TextShadow`, `ThreeD`. `DiacriticColor` (complex-script only) |

### 3.4 ParagraphFormat
<!-- learn: paragraphformat -->

| Status | Members |
|---|---|
| Done | `Reset`, `Alignment`, `FirstLineIndent`, `Hyphenation`, `KeepTogether`, `KeepWithNext`, `LeftIndent`, `LineSpacing`, `LineSpacingRule`, `MirrorIndents`, `NoLineNumber`, `OutlineLevel`, `PageBreakBefore`, `RightIndent`, `SpaceAfter`, `SpaceAfterAuto`, `SpaceBefore`, `SpaceBeforeAuto`, `TabStops`, `WidowControl` |
| Done (step 7b) | `OpenUp`, `CloseUp`, `OpenOrCloseUp` (12 points before), `Space1`, `Space15`, `Space2`, `TabIndent`, `TabHangingIndent` (by default tab stops, which is `Settings.DefaultTabStop`), `Duplicate` (also `Range.ParagraphFormat = format`), `ReadingOrder`, `WordWrap`, `AutoAdjustRightIndent`, `BaseLineAlignment`, `TextboxTightWrap`, and the East Asian switches `AddSpaceBetweenFarEastAndAlpha`, `AddSpaceBetweenFarEastAndDigit`, `DisableLineHeightGrid`, `FarEastLineBreakControl`, `HalfWidthPunctuationOnTopOfLine`, `HangingPunctuation` (each a single element, defaults as Word's) |
| Planned, step 7 | `Borders`, `Shading`, `Parent`, `CollapsedByDefault` |
| Decided against | `CharacterUnitFirstLineIndent`, `CharacterUnitLeftIndent`, `CharacterUnitRightIndent`, `LineUnitAfter`, `LineUnitBefore`, `IndentCharWidth`, `IndentFirstLineCharWidth`: they store a count of characters or lines for the East Asian grid and Word 16 is inconsistent about them (`IndentCharWidth(3)` changed nothing, `IndentFirstLineCharWidth(2)` set the indent but not the unit). They need reference behaviour we cannot reproduce |
| Planned, step 9 | `Style` |

### 3.5 Paragraph
<!-- learn: paragraph -->

In Word, `Paragraph` exposes the whole of `ParagraphFormat`; here it forwards only part of it. Step 7 forwards the rest.

| Status | Members |
|---|---|
| Done | `Next`, `Previous`, `Alignment`, `FirstLineIndent`, `Format`, `KeepTogether`, `KeepWithNext`, `LeftIndent`, `LineSpacing`, `LineSpacingRule`, `PageBreakBefore`, `Parent`, `Range`, `RightIndent`, `SpaceAfter`, `SpaceBefore`, `Style`, `WidowControl` |
| Done (step 7b) | everything `ParagraphFormat` has that is done, and `Indent` / `Outdent` (to the next or previous default tab stop), `Reset`, `OpenUp`, `CloseUp`, `OpenOrCloseUp`, `Space1`, `Space15`, `Space2`, `TabIndent`, `TabHangingIndent`, `Format` can be assigned |
| Planned, step 7 | `OutlineDemote`, `OutlineDemoteToBody`, `OutlinePromote`, `Borders`, `Shading`, `CollapseHeadingByDefault`, `IsStyleSeparator` |
| Decided against | the character and line unit members and `IndentCharWidth`, `IndentFirstLineCharWidth` (see ParagraphFormat) |
| Planned, step 13 | `JoinList`, `SeparateList`, `ListAdvanceTo`, `ResetAdvanceTo`, `ListNumberOriginal` |
| Out of scope | `CollapsedState` (view state), `SelectNumber` (UI), `DropCap` (drop caps are frames) |
| Decided against | `ID` (an HTML id) |

### 3.6 Paragraphs
<!-- learn: paragraphs -->

| Status | Members |
|---|---|
| Done | `Add` (without Word's optional range; step 7 adds it), `Item`, `Count`, `First`, `Last`, `Format`, `Parent`, `Style` |
| Planned, step 7 | `CloseUp`, `DecreaseSpacing`, `IncreaseSpacing`, `Indent`, `IndentCharWidth`, `IndentFirstLineCharWidth`, `OpenOrCloseUp`, `OpenUp`, `Outdent`, `OutlineDemote`, `OutlineDemoteToBody`, `OutlinePromote`, `Reset`, `Space1`, `Space15`, `Space2`, `TabHangingIndent`, `TabIndent`, and the collection-level format properties: `AddSpaceBetweenFarEastAndAlpha`, `AddSpaceBetweenFarEastAndDigit`, `Alignment`, `AutoAdjustRightIndent`, `BaseLineAlignment`, `Borders`, `CharacterUnitFirstLineIndent`, `CharacterUnitLeftIndent`, `CharacterUnitRightIndent`, `DisableLineHeightGrid`, `FarEastLineBreakControl`, `FirstLineIndent`, `HalfWidthPunctuationOnTopOfLine`, `HangingPunctuation`, `Hyphenation`, `KeepTogether`, `KeepWithNext`, `LeftIndent`, `LineSpacing`, `LineSpacingRule`, `LineUnitAfter`, `LineUnitBefore`, `NoLineNumber`, `OutlineLevel`, `PageBreakBefore`, `ReadingOrder`, `RightIndent`, `Shading`, `SpaceAfter`, `SpaceAfterAuto`, `SpaceBefore`, `SpaceBeforeAuto`, `TabStops`, `WidowControl`, `WordWrap` |

### 3.7 Find
<!-- learn: find -->

| Status | Members |
|---|---|
| Done | `ClearFormatting`, `Execute` (Word's parameters in Word's order), `Forward`, `Found`, `MatchCase`, `MatchWholeWord`, `MatchWildcards`, `Replacement`, `Text`, `Wrap` (`FindWrap.Ask` throws) |
| Planned, step 9 | `Format` (exists; `true` throws until this step), `Font`, `ParagraphFormat`, `Style`, `Highlight`, `MatchPrefix`, `MatchSuffix`, `MatchPhrase`, `IgnorePunct`, `IgnoreSpace`, `NoProofing`, `LanguageID`, `LanguageIDFarEast`, `LanguageIDOther`, `Parent` |
| Out of scope | `MatchSoundsLike`, `MatchAllWordForms` (need Word's linguistic data; they exist and throw when set to `true`), `HitHighlight`, `ClearHitHighlight` (UI), `Frame` (frames) |
| Decided against | `Execute2007` (the same as Execute). Script-specific matching, which needs reference behaviour we cannot test: `MatchFuzzy`, `SetAllFuzzyOptions`, `ClearAllFuzzyOptions`, `MatchByte`, `MatchAlefHamza`, `MatchControl`, `MatchDiacritics`, `MatchKashida`, `CorrectHangulEndings`, `HanjaPhoneticHangul` |

### 3.8 Replacement
<!-- learn: replacement -->

| Status | Members |
|---|---|
| Done | `ClearFormatting` (does nothing until step 9 gives it formatting to clear), `Text` |
| Planned, step 9 | `Font`, `ParagraphFormat`, `Style`, `Highlight`, `LanguageID`, `LanguageIDFarEast`, `NoProofing`, `Parent` |
| Out of scope | `Frame` (frames) |

### 3.9 Style, Styles, TableStyle, ConditionalStyle
<!-- learn: style -->

**Style.** `Name` and `StyleId` are extra and stay.

| Status | Members |
|---|---|
| Done | `Delete`, `BaseStyle`, `BuiltIn`, `Font`, `NameLocal`, `NextParagraphStyle`, `ParagraphFormat`, `Type` |
| Planned, step 9 | `AutomaticallyUpdate`, `Borders`, `Shading`, `InUse`, `LanguageID`, `LanguageIDFarEast`, `NoProofing`, `Linked`, `LinkStyle`, `Locked`, `Priority`, `QuickStyle`, `UnhideWhenUsed`, `Visibility`, `NoSpaceBetweenParagraphsOfSameStyle`, `Table`, `Parent` |
| Planned, step 13 | `ListLevelNumber`, `ListTemplate`, `LinkToListTemplate` |
| Out of scope | `Frame` (frames) |
| Decided against | `Description` (the summary Word generates for its UI, such as "Font: Bold, Style: Quick Style") |

<!-- learn: styles -->

**Styles.**

| Status | Members |
|---|---|
| Done | `Add`, `Item` (by index, name or `BuiltInStyle`), `Count` |
| Planned, step 9 | `Parent` |

<!-- learn: tablestyle -->

**TableStyle** (`Style.Table`).

| Status | Members |
|---|---|
| Planned, step 9 | `Condition`, `Alignment`, `AllowBreakAcrossPage`, `AllowPageBreaks`, `Borders`, `BottomPadding`, `ColumnStripe`, `LeftIndent`, `LeftPadding`, `Parent`, `RightPadding`, `RowStripe`, `Shading`, `Spacing`, `TableDirection`, `TopPadding` |

<!-- learn: conditionalstyle -->

**ConditionalStyle** (`TableStyle.Condition`).

| Status | Members |
|---|---|
| Planned, step 9 | `Borders`, `BottomPadding`, `Font`, `LeftPadding`, `ParagraphFormat`, `Parent`, `RightPadding`, `Shading`, `TopPadding` |

### 3.10 Table and Tables
<!-- learn: table -->

**Table.** `AppendRow`, `MergeTable` and `Width` are extra (section 4).

| Status | Members |
|---|---|
| Done | `AutoFitBehavior`, `Cell`, `Delete`, `Split`, `AllowAutoFit`, `ApplyStyleColumnBands`, `ApplyStyleFirstColumn`, `ApplyStyleHeadingRows`, `ApplyStyleLastColumn`, `ApplyStyleLastRow`, `ApplyStyleRowBands`, `Borders`, `BottomPadding`, `Columns`, `Descr`, `LeftPadding`, `NestingLevel`, `Parent`, `PreferredWidth`, `PreferredWidthType`, `Range`, `RightPadding`, `Rows`, `Shading`, `Spacing`, `Style`, `Tables`, `Title`, `TopPadding`, `Uniform` |
| Planned, step 8 | `ConvertToText`, `Sort`, `SortAscending`, `SortDescending`, `TableDirection`, `ApplyStyleDirectFormatting` |
| Planned, step 19 | `Select` |
| Decided against | `AutoFormat`, `AutoFormatType`, `UpdateAutoFormat` (Word 97 table AutoFormat; table styles replace it), `ID` (an HTML id) |

<!-- learn: tables -->

**Tables.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count`, `NestingLevel`, `Parent` |

### 3.11 Rows, Row, Columns, Column, Cells, Cell
<!-- learn: rows -->

The collections are where tables are thinnest: Word's `Rows`, `Columns` and `Cells` format the whole set, and `Range.Cells` is the common way to format a block. Step 8 also adds a pass at open that turns legacy `hMerge` merges into `gridSpan`, so merge, split and column delete stop throwing for them.

**Rows.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Alignment`, `AllowBreakAcrossPages`, `Count`, `First`, `HeadingFormat`, `Height`, `HeightRule`, `Last`, `LeftIndent`, `NestingLevel`, `Parent`, `Shading` |
| Planned, step 8 | `ConvertToText`, `Delete`, `DistributeHeight`, `SetHeight`, `SetLeftIndent`, `AllowOverlap`, `Borders`, `DistanceBottom`, `DistanceLeft`, `DistanceRight`, `DistanceTop`, `HorizontalPosition`, `VerticalPosition`, `RelativeHorizontalPosition`, `RelativeVerticalPosition`, `WrapAroundText`, `SpaceBetweenColumns`, `TableDirection` |
| Planned, step 19 | `Select` |

<!-- learn: row -->

**Row.** `ClearText` and `Copy` are extra (section 4).

| Status | Members |
|---|---|
| Done | `Delete`, `SetHeight`, `Alignment`, `AllowBreakAcrossPages`, `Borders`, `Cells`, `HeadingFormat`, `Height`, `HeightRule`, `Index`, `IsFirst`, `IsLast`, `Next`, `Parent`, `Previous`, `Range`, `Shading` |
| Planned, step 8 | `ConvertToText`, `SetLeftIndent`, `LeftIndent`, `NestingLevel`, `SpaceBetweenColumns` |
| Planned, step 19 | `Select` |
| Decided against | `ID` (an HTML id) |

<!-- learn: columns -->

**Columns.**

| Status | Members |
|---|---|
| Done | `Add`, `DistributeWidth`, `Item`, `Count`, `First`, `Last`, `Parent` |
| Planned, step 8 | `AutoFit`, `Delete`, `SetWidth`, `Borders`, `NestingLevel`, `PreferredWidth`, `PreferredWidthType`, `Shading`, `Width` |
| Planned, step 19 | `Select` |

<!-- learn: column -->

**Column.**

| Status | Members |
|---|---|
| Done | `Delete`, `Borders`, `Cells`, `Index`, `IsFirst`, `IsLast`, `Parent`, `PreferredWidth`, `PreferredWidthType`, `Shading`, `Width` |
| Planned, step 8 | `AutoFit`, `SetWidth`, `Sort`, `NestingLevel`, `Next`, `Previous` |
| Planned, step 19 | `Select` |

<!-- learn: cells -->

**Cells.**

| Status | Members |
|---|---|
| Done | `Item`, `Count`, `NestingLevel`, `Parent` |
| Planned, step 8 | `Add`, `AutoFit`, `Delete`, `DistributeHeight`, `DistributeWidth`, `Merge`, `SetHeight`, `SetWidth`, `Split`, `Borders`, `Height`, `HeightRule`, `PreferredWidth`, `PreferredWidthType`, `Shading`, `VerticalAlignment`, `Width` |

`Cell.Width` sets only the width that cell asks for. In Word it also changes the grid and `w:tblW`, as `Column.Width` does; that comes with these members in step 8.

<!-- learn: cell -->

**Cell.**

| Status | Members |
|---|---|
| Done | `Delete` (shifting left only; step 8 adds up, and vertically merged cells), `Merge`, `Split`, `Borders`, `BottomPadding`, `Column`, `ColumnIndex`, `FitText`, `Height`, `HeightRule`, `LeftPadding`, `Next`, `Parent`, `PreferredWidth`, `PreferredWidthType`, `Previous`, `Range`, `RightPadding`, `Row`, `RowIndex`, `Shading`, `Tables`, `TopPadding`, `VerticalAlignment`, `Width`, `WordWrap` |
| Planned, step 8 | `AutoSum`, `Formula`, `SetHeight`, `SetWidth`, `NestingLevel` |
| Planned, step 19 | `Select` |
| Decided against | `ID` (an HTML id) |

### 3.12 Bookmarks and Comments
<!-- learn: bookmark -->

**Bookmark.**

| Status | Members |
|---|---|
| Done | `Copy`, `Delete`, `Empty`, `End`, `Name`, `Parent`, `Range`, `Start` |
| Planned | `Column` (10), `StoryType` (11), `Select` (19) |

<!-- learn: bookmarks -->

**Bookmarks.** `DeleteAll` is extra.

| Status | Members |
|---|---|
| Done | `Add`, `Exists`, `Item`, `Count`, `DefaultSorting`, `Parent`, `ShowHidden` |

<!-- learn: comment -->

**Comment.** `Author`, `Initial`, `Text` and `Delete` are kept although the current reference drops them.

| Status | Members |
|---|---|
| Done | `Ancestor`, `Date`, `Done`, `Index`, `Parent`, `Range`, `Replies`, `Scope` |
| Planned, step 10 | `DeleteRecursively`, `Reference`, `Contact` (a small object over the author, initials and the people part) |
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
| Done | `Delete`, `Unlink`, `Update`, `Code`, `Index`, `Locked`, `Next`, `Parent`, `Previous`, `Result`, `Type` |
| Planned | `Data`, `Kind` (10); `InlineShape` (16); `Select` (19) |
| Out of scope | `Copy`, `Cut` (clipboard), `DoClick` (UI), `ShowCodes` (view state), `LinkFormat`, `OLEFormat` (OLE links) |
| Decided against | `UpdateSource` (writes a result back into a linked file) |

<!-- learn: fields -->

**Fields.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Update`, `Count`, `Locked` (with `Lock` and `Unlock`), `Parent` |
| Planned, step 10 | `Unlink` |
| Out of scope | `ToggleShowCodes` (view state) |
| Decided against | `UpdateSource` (as on `Field`) |

### 3.14 Sections, headers and footers, page setup
<!-- learn: section -->

**Section.**

| Status | Members |
|---|---|
| Done | `Borders`, `Footers`, `Headers`, `Index`, `PageSetup`, `Parent`, `Range` |
| Planned, step 10 | `ProtectedForForms` |

<!-- learn: sections -->

**Sections.**

| Status | Members |
|---|---|
| Done | `Add`, `Item`, `Count`, `First`, `Last`, `Parent` |
| Planned, step 10 | `PageSetup` |

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
| Planned, step 10 | `TogglePortrait`, `SectionDirection`, `CharsLine`, `LinesPage`, `LayoutMode`, `GutterStyle`, `Parent` |
| Out of scope | Printing: `TwoPagesOnOne`, `BookFoldPrinting`, `BookFoldRevPrinting`, `BookFoldPrintingSheets`, `FirstPageTray`, `OtherPagesTray`. View: `ShowGrid`. Templates: `SetAsTemplateDefault` |

<!-- learn: pagenumbers -->

**PageNumbers.**

| Status | Members |
|---|---|
| Done | `Add` (without Inside and Outside alignment), `Count`, `NumberStyle`, `RestartNumberingAtSection`, `StartingNumber` |
| Planned, step 10 | `Item`, `ChapterPageSeparator`, `HeadingLevelForChapter`, `IncludeChapterNumber`, `ShowFirstPageNumber`, `Parent` |
| Decided against | `DoubleQuote` (Hebrew numbering only) |

<!-- learn: pagenumber -->

**PageNumber.**

| Status | Members |
|---|---|
| Planned | `Delete`, `Alignment`, `Index`, `Parent` (10); `Select` (19) |
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
| Done | `SetCount`, `Count`, `EvenlySpaced`, `LineBetween`, `Spacing`, `Width` |
| Planned, step 10 | `Add`, `Item`, `FlowDirection`, `Parent` |

<!-- learn: textcolumn -->

**TextColumn.**

| Status | Members |
|---|---|
| Planned, step 10 | `SpaceAfter`, `Width`, `Parent` |

### 3.15 Borders, Border, Shading
<!-- learn: borders -->

`Borders` exists for tables, rows, columns, cells, sections and pages. Step 7 adds it to paragraphs, characters and ranges, step 9 to styles.

**Borders.** `SetOutside`, `SetInside` and `Clear` are extra.

| Status | Members |
|---|---|
| Done | `Item`, `Count`, `Enable`, `InsideColor`, `InsideLineStyle`, `InsideLineWidth`, `OutsideColor`, `OutsideLineStyle`, `OutsideLineWidth` |
| Planned, step 7 | `InsideColorIndex`, `OutsideColorIndex`, `Shadow`, `DistanceFromBottom`, `DistanceFromLeft`, `DistanceFromRight`, `DistanceFromTop`, `Parent` |
| Planned, step 8 | `HasHorizontal`, `HasVertical` |
| Planned, step 10 | `ApplyPageBordersToAllSections`, `AlwaysInFront`, `DistanceFrom`, `EnableFirstPageInSection`, `EnableOtherPagesInSection`, `JoinBorders`, `SurroundFooter`, `SurroundHeader` |

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
| Planned, step 11 | `Item`, `Count`, `Parent` |

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
| Planned, step 12 | `Delete`, `Address`, `SubAddress`, `EmailSubject`, `Name`, `Range`, `ScreenTip`, `Target`, `TextToDisplay`, `Type`, `Parent` |
| Planned, step 18 | `Shape` |
| Out of scope | `AddToFavorites`, `Follow`, `ExtraInfoRequired` (browser and UI) |
| Decided against | `CreateNewDocument` (creates a file as a side effect) |

<!-- learn: hyperlinks -->

**Hyperlinks.**

| Status | Members |
|---|---|
| Planned, step 12 | `Add`, `Item`, `Count`, `Parent` |

### 4.2 Lists (step 13)

`ApplyBulletDefault`, `ApplyNumberDefault` and `ApplyOutlineNumberDefault` use list definitions that match Word 365's defaults, not the user's galleries.

<!-- learn: listformat -->

**ListFormat.**

| Status | Members |
|---|---|
| Planned, step 13 | `ApplyBulletDefault`, `ApplyListTemplate`, `ApplyListTemplateWithLevel`, `ApplyNumberDefault`, `ApplyOutlineNumberDefault`, `CanContinuePreviousList`, `ConvertNumbersToText`, `CountNumberedItems`, `ListIndent`, `ListOutdent`, `RemoveNumbers`, `List`, `ListLevelNumber`, `ListString`, `ListTemplate`, `ListType`, `ListValue`, `Parent`, `SingleList`, `SingleListTemplate` |
| Decided against | `ListPictureBullet` (legacy picture bullets) |

<!-- learn: list -->

**List.**

| Status | Members |
|---|---|
| Planned, step 13 | `ApplyListTemplate`, `ApplyListTemplateWithLevel`, `CanContinuePreviousList`, `ConvertNumbersToText`, `CountNumberedItems`, `RemoveNumbers`, `ListParagraphs`, `Parent`, `Range`, `SingleListTemplate`, `StyleName` |

<!-- learn: lists,listparagraphs,listlevels -->

**Lists, ListParagraphs, ListLevels.**

| Status | Members |
|---|---|
| Planned, step 13 | `Item`, `Count`, `Parent` |

<!-- learn: listtemplate -->

**ListTemplate.**

| Status | Members |
|---|---|
| Planned, step 13 | `Convert`, `ListLevels`, `Name`, `OutlineNumbered`, `Parent` |

<!-- learn: listtemplates -->

**ListTemplates.**

| Status | Members |
|---|---|
| Planned, step 13 | `Add`, `Item`, `Count`, `Parent` |

<!-- learn: listlevel -->

**ListLevel.**

| Status | Members |
|---|---|
| Planned, step 13 | `Alignment`, `Font`, `Index`, `LinkedStyle`, `NumberFormat`, `NumberPosition`, `NumberStyle`, `Parent`, `ResetOnHigher`, `StartAt`, `TabPosition`, `TextPosition`, `TrailingCharacter` |
| Decided against | `ApplyPictureBullet`, `PictureBullet` (legacy picture bullets) |

### 4.3 Footnotes and endnotes (steps 11 and 14)

Step 11 lays the groundwork every new story needs: `StoryType`, footnote, endnote and text-box stories in the story map and the span registry, and `StoryRanges`.

<!-- learn: footnote,endnote -->

**Footnote, Endnote.**

| Status | Members |
|---|---|
| Planned, step 14 | `Delete`, `Index`, `Parent`, `Range`, `Reference` |

<!-- learn: footnotes -->

**Footnotes.**

| Status | Members |
|---|---|
| Planned, step 14 | `Add`, `Convert`, `Item`, `ResetContinuationNotice`, `ResetContinuationSeparator`, `ResetSeparator`, `SwapWithEndnotes`, `ContinuationNotice`, `ContinuationSeparator`, `Count`, `Location`, `NumberingRule`, `NumberStyle`, `Parent`, `Separator`, `StartingNumber` |

<!-- learn: endnotes -->

**Endnotes.**

| Status | Members |
|---|---|
| Planned, step 14 | `Add`, `Convert`, `Item`, `ResetContinuationNotice`, `ResetContinuationSeparator`, `ResetSeparator`, `SwapWithFootnotes`, `ContinuationNotice`, `ContinuationSeparator`, `Count`, `Location`, `NumberingRule`, `NumberStyle`, `Parent`, `Separator`, `StartingNumber` |

<!-- learn: footnoteoptions -->

**FootnoteOptions.**

| Status | Members |
|---|---|
| Planned, step 14 | `LayoutColumns`, `Location`, `NumberingRule`, `NumberStyle`, `Parent`, `StartingNumber` |

<!-- learn: endnoteoptions -->

**EndnoteOptions.**

| Status | Members |
|---|---|
| Planned, step 14 | `Location`, `NumberingRule`, `NumberStyle`, `Parent`, `StartingNumber` |

### 4.4 Content controls and custom XML parts (step 15)
<!-- learn: contentcontrol -->

**ContentControl.**

| Status | Members |
|---|---|
| Planned, step 15 | `Delete`, `SetCheckedSymbol`, `SetPlaceholderText`, `SetUncheckedSymbol`, `Ungroup`, `AllowInsertDeleteSection`, `Appearance`, `BuildingBlockCategory`, `BuildingBlockType`, `Checked`, `Color`, `DateCalendarType`, `DateDisplayFormat`, `DateDisplayLocale`, `DateStorageFormat`, `DefaultTextStyle`, `DropdownListEntries`, `ID`, `Level`, `LockContentControl`, `LockContents`, `MultiLine`, `Parent`, `ParentContentControl`, `PlaceholderText`, `Range`, `RepeatingSectionItems`, `RepeatingSectionItemTitle`, `ShowingPlaceholderText`, `Tag`, `Temporary`, `Title`, `Type`, `XMLMapping` |
| Out of scope | `Copy`, `Cut` (clipboard) |

<!-- learn: contentcontrols -->

**ContentControls.**

| Status | Members |
|---|---|
| Planned, step 15 | `Add`, `Item`, `Count`, `Parent` |

<!-- learn: contentcontrollistentries -->

**ContentControlListEntries.**

| Status | Members |
|---|---|
| Planned, step 15 | `Add`, `Clear`, `Item`, `Count`, `Parent` |

<!-- learn: contentcontrollistentry -->

**ContentControlListEntry.**

| Status | Members |
|---|---|
| Planned | `Delete`, `MoveDown`, `MoveUp`, `Index`, `Parent`, `Text`, `Value` (15); `Select` (19) |

<!-- learn: xmlmapping -->

**XMLMapping.**

| Status | Members |
|---|---|
| Planned, step 15 | `Delete`, `SetMapping`, `SetMappingByNode`, `CustomXMLNode`, `CustomXMLPart`, `IsMapped`, `Parent`, `PrefixMappings`, `XPath` |

<!-- learn: office:customxmlparts -->

**CustomXMLParts** (Office).

| Status | Members |
|---|---|
| Planned, step 15 | `Add`, `Item`, `SelectByID`, `SelectByNamespace`, `Count`, `Parent` |
| Out of scope | The events `PartAfterAdd`, `PartAfterLoad`, `PartBeforeDelete` |

<!-- learn: office:customxmlpart -->

**CustomXMLPart** (Office).

| Status | Members |
|---|---|
| Planned, step 15 | `AddNode`, `Delete`, `Load`, `LoadXML`, `SelectNodes`, `SelectSingleNode`, `BuiltIn`, `DocumentElement`, `Id`, `NamespaceManager`, `NamespaceURI`, `Parent`, `XML` |
| Out of scope | The events `NodeAfterDelete`, `NodeAfterInsert`, `NodeAfterReplace` |
| Decided against | `SchemaCollection`, `Errors` (schema validation of the part) |

### 4.5 Inline shapes (step 16)
<!-- learn: inlineshape -->

**InlineShape.**

| Status | Members |
|---|---|
| Planned, step 16 | `Delete`, `Reset`, `AlternativeText`, `Borders`, `Field`, `Height`, `Hyperlink`, `LockAspectRatio`, `Parent`, `PictureFormat`, `Range`, `ScaleHeight`, `ScaleWidth`, `Title`, `Type`, `Width` |
| Planned, steps 18 and 19 | `ConvertToShape` (18), `Select` (19) |
| Out of scope | Charts and SmartArt: `Chart`, `HasChart`, `SmartArt`, `HasSmartArt`. OLE: `OLEFormat`, `LinkFormat`. HTML: `Script` |
| Decided against | Picture effects, each an object of its own: `Fill`, `Glow`, `Line`, `Reflection`, `Shadow`, `SoftEdge`, `TextEffect`. `GroupItems`, `HorizontalLineFormat`, `IsPictureBullet` (legacy VML) |

<!-- learn: inlineshapes -->

**InlineShapes.**

| Status | Members |
|---|---|
| Planned, step 16 | `AddPicture`, `Item`, `Count`, `Parent` |
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
* `Find` cannot search or replace by formatting until step 9, and has no sounds-like or word-forms matching.
* Legacy `hMerge` tables are rejected by merge, split and column delete until step 8. `Cell.Delete` shifts left only and not on vertically merged cells.
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
