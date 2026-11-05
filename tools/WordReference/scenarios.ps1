<#
.SYNOPSIS
The scenarios New-WordReferences.ps1 runs in Word. Each is a scriptblock of Word object model calls on a new document (Documents.Add()).
.DESCRIPTION
The name is the folder the snapshot goes to under OpenXml.Wrappers.Test/TestFiles/WordReference, and MarkupConformanceTest has a case with
the same name that does the same thing through the library. Word's indexes are 1-based; the library's are 0-based. Keep the two in step.
#>

# Twenty three-letter words, one for each font property below: word i is at positions 4i to 4i + 3.
$script:FontWords = (1..20 | ForEach-Object { 'w{0:00}' -f $_ }) -join ' '

function Set-FontProperty($font, [int] $index, [bool] $on) {
    switch ($index) {
        0 { $font.Italic = $on }
        1 { $font.StrikeThrough = $on }
        2 { $font.DoubleStrikeThrough = $on }
        3 { $font.AllCaps = $on }
        4 { $font.SmallCaps = $on }
        5 { $font.Hidden = $on }
        6 { $font.Emboss = $on }
        7 { $font.Engrave = $on }
        8 { $font.Outline = $on }
        9 { $font.Shadow = $on }
        10 { $font.Superscript = $on }
        11 { $font.Subscript = $on }
        12 { $font.Underline = $(if ($on) { 1 } else { 0 }) }  # wdUnderlineSingle, wdUnderlineNone
        13 { $font.Size = $(if ($on) { 14 } else { 12 }) }
        14 { $font.Color = $(if ($on) { 255 } else { -16777216 }) }  # wdColorRed, wdColorAutomatic
        15 { $font.Spacing = $(if ($on) { 1.5 } else { 0.0 }) }
        16 { $font.Scaling = $(if ($on) { 150 } else { 100 }) }
        17 { $font.Position = $(if ($on) { 3 } else { 0 }) }
        18 { $font.Kerning = $(if ($on) { 14 } else { 1 }) }
        19 { $font.Bold = $on }
    }
}

function Add-Paragraphs($doc, [int] $count) {
    for ($i = 0; $i -lt $count; $i++) {
        $doc.Content.InsertAfter('p')
        $doc.Content.InsertParagraphAfter()
    }
}

function Set-ParagraphProperty($format, [int] $index, [bool] $on) {
    switch ($index) {
        0 { $format.Alignment = $(if ($on) { 1 } else { 0 }) }
        1 { $format.Alignment = $(if ($on) { 2 } else { 0 }) }
        2 { $format.Alignment = $(if ($on) { 3 } else { 0 }) }
        3 { $format.LeftIndent = $(if ($on) { 18 } else { 0 }) }
        4 { $format.RightIndent = $(if ($on) { 18 } else { 0 }) }
        5 { $format.FirstLineIndent = $(if ($on) { 18 } else { 0 }) }
        6 { $format.FirstLineIndent = $(if ($on) { -18 } else { 0 }) }
        7 { $format.SpaceBefore = $(if ($on) { 6 } else { 0 }) }
        8 { $format.SpaceAfter = $(if ($on) { 6 } else { 8 }) }
        9 { $format.LineSpacingRule = $(if ($on) { 1 } else { 0 }) }  # wdLineSpace1pt5, wdLineSpaceSingle
        10 { $format.KeepWithNext = $on }
        11 { $format.KeepTogether = $on }
        12 { $format.PageBreakBefore = $on }
        13 { $format.WidowControl = -not $on }
        14 { $format.OutlineLevel = $(if ($on) { 1 } else { 10 }) }  # wdOutlineLevel1, wdOutlineLevelBodyText
        15 { $format.NoLineNumber = $on }
        16 { $format.Hyphenation = -not $on }
    }
}

[ordered]@{
    'blank-document'                  = { param($doc) }

    'insert-after-tab-and-line-break' = {
        param($doc)
        # [char]11 is a manual line break in Word's text (vbVerticalTab).
        $doc.Content.InsertAfter("Before`tafter the tab" + [char]11 + 'after the break')
    }

    'bold-on'                         = {
        param($doc)
        $doc.Content.InsertAfter('Bold text')
        $doc.Range(0, 4).Font.Bold = $true
    }

    'bold-on-then-off'                = {
        param($doc)
        $doc.Content.InsertAfter('Bold text')
        $doc.Range(0, 4).Font.Bold = $true
        $doc.Range(0, 4).Font.Bold = $false
    }

    'font-name'                       = {
        param($doc)
        $doc.Content.InsertAfter('Arial text')
        $doc.Range(0, 5).Font.Name = 'Arial'
    }
    'font-complex-script'             = {
        param($doc)
        $doc.Content.InsertAfter('Complex text here')
        $font = $doc.Range(0, 7).Font
        $font.BoldBi = $true
        $font.ItalicBi = $true
        $font.SizeBi = 14
        $font.NameBi = 'Tahoma'
        $font.NameFarEast = 'MS Gothic'
        $font.NameOther = 'Verdana'
    }

    'font-emphasis-grid'              = {
        param($doc)
        $doc.Content.InsertAfter('Mark text here')
        $doc.Range(0, 4).Font.EmphasisMark = 1
        $doc.Range(5, 9).Font.EmphasisMark = 4
        $doc.Range(10, 14).Font.DisableCharacterSpaceGrid = $true
    }

    'font-name-after-complex'         = {
        param($doc)
        $doc.Content.InsertAfter('Name text here')
        $doc.Range(0, 4).Font.NameBi = 'Tahoma'
        $doc.Range(0, 4).Font.Name = 'Calibri'
    }

    'font-grow-shrink'                = {
        param($doc)
        $doc.Content.InsertAfter('Grow shrink text')
        $doc.Range(0, 4).Font.Grow()
        $doc.Range(5, 11).Font.Shrink()
        $doc.Range(12, 16).Font.Grow()
        $doc.Range(12, 16).Font.Grow()
    }

    'paragraph-spacing-shortcuts'     = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree`rFour")
        $call = { param($o, $n, $a) [void][System.__ComObject].InvokeMember($n, [Reflection.BindingFlags]::InvokeMethod, $null, $o, $a) }
        & $call $doc.Paragraphs.Item(1) 'OpenUp' @()
        & $call $doc.Paragraphs.Item(2) 'Space15' @()
        & $call $doc.Paragraphs.Item(3) 'Indent' @()
        & $call $doc.Paragraphs.Item(4).Format 'TabHangingIndent' @(1)
    }

    'paragraph-east-asian-switches'   = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo")
        $format = $doc.Paragraphs.Item(1).Format
        $format.AddSpaceBetweenFarEastAndAlpha = 0
        $format.BaseLineAlignment = 1
        $format.HangingPunctuation = 0
        $format.FarEastLineBreakControl = 0
        $format.HalfWidthPunctuationOnTopOfLine = -1
        $format.DisableLineHeightGrid = -1
        $format.TextboxTightWrap = 1
        $format.WordWrap = 0
        $format.AutoAdjustRightIndent = 0
        $format.AddSpaceBetweenFarEastAndDigit = 0
    }

    'paragraph-right-to-left'         = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo")
        $doc.Paragraphs.Item(1).Format.ReadingOrder = 0
    }
    'paragraph-borders-shading'       = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree`rFour text")
        $doc.Paragraphs.Item(1).Borders.Item(-1).LineStyle = 1
        $outside = $doc.Paragraphs.Item(2).Borders
        $outside.OutsideLineStyle = 1
        $outside.OutsideLineWidth = 12
        $outside.OutsideColor = 255
        $doc.Paragraphs.Item(3).Borders.Item(-1).LineStyle = 1
        $doc.Paragraphs.Item(3).Borders.Item(-2).LineStyle = 1
        $doc.Paragraphs.Item(4).Shading.BackgroundPatternColor = 255
        $doc.Paragraphs.Item(4).Shading.Texture = 100
    }

    'font-border-shading'             = {
        param($doc)
        $doc.Content.InsertAfter('One two three four')
        $doc.Range(0, 3).Font.Borders.Item(-1).LineStyle = 1
        $doc.Range(4, 7).Font.Shading.BackgroundPatternColor = 65280
        $doc.Range(8, 13).Font.Borders.OutsideLineStyle = 7
    }
    'paragraph-outline'               = {
        param($doc)
        $doc.Content.InsertAfter("A`rB`rC`rD`rE")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(2).Style = 'Heading 2'
        $doc.Paragraphs.Item(5).Style = 'Heading 3'
        $doc.Paragraphs.Item(4).OutlineLevel = 3
        $call = { param($o, $n, $a) [void][System.__ComObject].InvokeMember($n, [Reflection.BindingFlags]::InvokeMethod, $null, $o, $a) }
        & $call $doc.Paragraphs.Item(1) 'OutlineDemote' @()
        & $call $doc.Paragraphs.Item(2) 'OutlinePromote' @()
        & $call $doc.Paragraphs.Item(3) 'OutlineDemote' @()
        & $call $doc.Paragraphs.Item(4) 'OutlineDemote' @()
        & $call $doc.Paragraphs.Item(5) 'OutlineDemoteToBody' @()
    }

    'paragraphs-spacing-add'          = {
        param($doc)
        $doc.Content.InsertAfter("A`rB`rC")
        $call = { param($o, $n, $a) [void][System.__ComObject].InvokeMember($n, [Reflection.BindingFlags]::InvokeMethod, $null, $o, $a) }
        & $call $doc.Paragraphs 'IncreaseSpacing' @()
        & $call $doc.Paragraphs 'IncreaseSpacing' @()
        & $call $doc.Paragraphs 'DecreaseSpacing' @()
        & $call $doc.Paragraphs 'Indent' @()
        [void]$doc.Paragraphs.Add($doc.Paragraphs.Item(2).Range)
    }
    'cells-collection-format'         = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $cells = $doc.Tables.Item(1).Range.Cells
        $cells.Shading.BackgroundPatternColor = 255
        $cells.VerticalAlignment = 1
        $cells.Height = 30
        $cells.PreferredWidthType = 3
        $cells.PreferredWidth = 80
        $doc.Tables.Item(1).Rows.Item(2).Cells.SetHeight(40, 2)
    }

    'columns-collection-format'       = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $columns = $doc.Tables.Item(1).Columns
        $columns.Width = 100
        $columns.PreferredWidthType = 3
        $columns.PreferredWidth = 90
        $columns.Shading.BackgroundPatternColor = 255
    }

    'rows-set-height'                 = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Rows.SetHeight(40, 2)
        $doc.Tables.Item(1).Cell(3, 1).SetHeight(25, 1)
    }
    'cell-width-ragged'               = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Cell(2, 2).Width = 100
    }
    'cell-set-width-proportional'     = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Cell(2, 1).SetWidth(100, 1)
    }
    'cell-set-width-first-column'     = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Cell(2, 1).SetWidth(100, 2)
    }
    'cells-set-width-row'             = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Rows.Item(2).Cells.SetWidth(100, 0)
    }
    'cells-distribute-width'          = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Cell(2, 2).Width = 100
        $doc.Tables.Item(1).Range.Cells.DistributeWidth()
    }
    'columns-set-width'               = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Columns.Item(1).SetWidth(100, 2)
        $doc.Tables.Item(1).Columns.Item(2).SetWidth(120, 1)
    }
    'columns-distribute-width'        = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Columns.Item(1).Width = 100
        $doc.Tables.Item(1).Columns.DistributeWidth()
    }
    'table-3x3'                       = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
    }

    # The table the Insert Table command makes: Word 2000 behaviour, which gives it the Table Grid style.
    'table-3x3-word9-behavior'        = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3, 1, 0)  # wdWord9TableBehavior, wdAutoFitFixed
    }

    'cell-merge'                      = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $table.Cell(1, 1).Merge($table.Cell(1, 2))
    }

    'comment-with-reply'              = {
        param($doc)
        $doc.Content.InsertAfter('Commented text')
        $comment = $doc.Comments.Add($doc.Range(0, 9), 'A comment')
        [void]$comment.Replies.Add($comment.Scope, 'A reply')
    }

    'bookmark'                        = {
        param($doc)
        $doc.Content.InsertAfter('Bookmarked text')
        [void]$doc.Bookmarks.Add('Mark', $doc.Range(0, 10))
    }

    'date-field'                      = {
        param($doc)
        [void]$doc.Fields.Add($doc.Range(0, 0), 31)  # wdFieldDate
    }

    'header-footer-text'              = {
        param($doc)
        $section = $doc.Sections.Item(1)
        $section.Headers.Item(1).Range.Text = 'Header text'  # wdHeaderFooterPrimary
        $section.Footers.Item(1).Range.Text = 'Footer text'
    }

    'section-add'                     = {
        param($doc)
        $doc.Content.InsertAfter('First section')
        [void]$doc.Sections.Add()
    }

    'page-borders'                    = {
        param($doc)
        $doc.Sections.Item(1).Borders.Enable = $true
    }

    'protect-read-only'               = {
        param($doc)
        $doc.Protect(3, $false, 'password')  # wdAllowOnlyReading
    }

    'heading-insert-paragraph-after'  = {
        param($doc)
        $doc.Content.InsertAfter('Heading')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
    }

    # Each built-in style the library can add (BuiltInStyle, plus Table Grid), applied in turn: Word writes the definition of each, and of
    # the linked character style of a paragraph style. These are also the definitions the library adds (OpenXml.Wrappers/Word/Templates).
    # Keep the list in step with BuiltInStylesScenario in MarkupConformanceTest.
    'built-in-styles'                 = {
        param($doc)
        $styles = @(-2, -3, -4, -5, -6, -7, -8, -9, -10, -20, -21, -22, -23, -24, -25, -26, -27, -28, -30, -31, -32, -33, -35, -39, -40, -41,
            -42, -43, -44, -63, -64, -65, -67, -68, -75, -76, -77, -81, -82, -86, -87, -88, -89, -91, -180, -181, -182, -261, -262, -263, -264, -265,
            'Comment Subject')  # no WdBuiltinStyle constant; Word adds it with a comment
        foreach ($id in $styles) {
            $style = $doc.Styles.Item([object]$id)
            $doc.Content.InsertAfter('x')
            $paragraph = $doc.Paragraphs.Item($doc.Paragraphs.Count).Range
            if ($style.Type -eq 2) {  # wdStyleTypeCharacter
                $doc.Range($paragraph.Start, $paragraph.Start + 1).Style = $style
            }
            else {
                $paragraph.Style = $style
            }
            $doc.Content.InsertParagraphAfter()
        }
        $table = $doc.Tables.Add($doc.Paragraphs.Item($doc.Paragraphs.Count).Range, 1, 1)
        $table.Style = 'Table Grid'
    }

    'font-properties-on'              = {
        param($doc)
        $doc.Content.InsertAfter($script:FontWords)
        for ($i = 0; $i -lt 20; $i++) {
            Set-FontProperty $doc.Range(4 * $i, 4 * $i + 3).Font $i $true
        }
    }

    'font-properties-on-then-off'     = {
        param($doc)
        $doc.Content.InsertAfter($script:FontWords)
        for ($i = 0; $i -lt 20; $i++) {
            Set-FontProperty $doc.Range(4 * $i, 4 * $i + 3).Font $i $true
            Set-FontProperty $doc.Range(4 * $i, 4 * $i + 3).Font $i $false
        }
    }

    'bold-off-in-a-bold-style'        = {
        param($doc)
        $doc.Content.InsertAfter('Strong text')
        $doc.Range(0, 6).Style = -88  # wdStyleStrong
        $doc.Range(0, 6).Font.Bold = $false
    }

    'runs-merge-when-formatting-matches' = {
        param($doc)
        $doc.Content.InsertAfter('One two three')
        $doc.Range(0, 3).Font.Bold = $true
        $doc.Range(3, 7).Font.Bold = $true
    }

    'paragraph-properties-on'         = {
        param($doc)
        Add-Paragraphs $doc 17
        for ($i = 0; $i -lt 17; $i++) {
            Set-ParagraphProperty $doc.Paragraphs.Item($i + 1).Range.ParagraphFormat $i $true
        }
    }

    'paragraph-properties-on-then-off' = {
        param($doc)
        Add-Paragraphs $doc 17
        for ($i = 0; $i -lt 17; $i++) {
            Set-ParagraphProperty $doc.Paragraphs.Item($i + 1).Range.ParagraphFormat $i $true
            Set-ParagraphProperty $doc.Paragraphs.Item($i + 1).Range.ParagraphFormat $i $false
        }
    }

    'paragraph-style-over-most-of-a-format' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 6).Font.Italic = 1
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-over-less-of-a-format' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 5).Font.Italic = 1
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-counts-the-mark' = {
        param($doc)
        $doc.Content.InsertAfter('012345678')
        $doc.Range(0, 4).Font.Italic = 1
        $doc.Range(9, 10).Font.Italic = 1
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-on-a-new-paragraph' = {
        param($doc)
        $doc.Content.InsertAfter('First')
        $doc.Paragraphs.Item(1).Range.Font.Italic = 1
        $doc.Content.InsertParagraphAfter()
        $doc.Paragraphs.Item(2).Range.InsertBefore('Contents')
        $doc.Paragraphs.Item(2).Style = 'Heading 1'
    }

    'paragraph-style-by-property'     = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 7).Font.Bold = 1
        $doc.Range(0, 3).Font.Italic = 1
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-by-value'        = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 3).Font.Size = 30
        $doc.Range(3, 7).Font.Size = 22
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-first-of-equal-values' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 3).Font.Size = 30
        $doc.Range(3, 6).Font.Size = 22
        $doc.Range(6, 9).Font.Size = 18
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-value-of-the-style-counts' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 3).Font.Size = 20
        $doc.Range(3, 7).Font.Size = 22
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'insert-symbol-ansi' = {
        param($doc)
        $doc.Content.InsertAfter('ab')
        $doc.Range(1, 1).InsertSymbol(171, 'Symbol', $false)
    }

    'insert-symbol-unicode' = {
        param($doc)
        $doc.Content.InsertAfter('ab')
        $doc.Range(1, 1).InsertSymbol(8364, 'Calibri', $true)
    }

    'insert-symbol-unicode-in-bold-text' = {
        param($doc)
        $doc.Content.InsertAfter('ab')
        $doc.Range(0, 2).Font.Bold = $true
        $doc.Range(1, 1).InsertSymbol(169, 'Arial', $true)
    }

    'insert-symbol-replaces-range' = {
        param($doc)
        $doc.Content.InsertAfter('ab')
        $doc.Range(0, 1).InsertSymbol(169, 'Times New Roman', $true)
    }

    'insert-alignment-tab-right-margin' = {
        param($doc)
        $doc.Content.InsertAfter('ab')
        $doc.Range(1, 1).InsertAlignmentTab(2, 0)
    }

    'insert-alignment-tab-center-indent' = {
        param($doc)
        $doc.Content.InsertAfter('ab')
        $doc.Range(1, 1).InsertAlignmentTab(1, 1)
    }

    'case-upper-part-of-bold-run' = {
        param($doc)
        $doc.Content.InsertAfter('abc DEF')
        $doc.Range(0, 3).Font.Bold = $true
        $doc.Range(0, 3).Case = 1
    }

    'case-upper-inside-a-run' = {
        param($doc)
        $doc.Content.InsertAfter('abcdef')
        $doc.Range(2, 4).Case = 1
    }

    'case-title-sentence' = {
        param($doc)
        $doc.Content.InsertAfter('hello wORLD. second sentence here')
        $doc.Range(0, 33).Case = 4
    }

    'formatted-text-copy-bold-to-end' = {
        param($doc)
        $doc.Content.InsertAfter('abc def ghi')
        $doc.Range(4, 7).Font.Bold = $true
        $target = $doc.Range(11, 11)
        $target.FormattedText = $doc.Range(4, 7).FormattedText
    }

    'formatted-text-copy-across-paragraphs' = {
        param($doc)
        $doc.Content.InsertAfter("aaa`rbbb`rccc")
        $doc.Paragraphs.Item(2).Alignment = 1
        $target = $doc.Range(11, 11)
        $target.FormattedText = $doc.Range(2, 6).FormattedText
    }

    'formatted-text-copy-paragraph-with-mark' = {
        param($doc)
        $doc.Content.InsertAfter("aaa`rbbb`rccc")
        $doc.Paragraphs.Item(2).Style = 'Heading 1'
        $target = $doc.Range(0, 0)
        $target.FormattedText = $doc.Paragraphs.Item(2).Range.FormattedText
    }

    'paragraph-style-same-value'      = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 3).Font.Size = 28
        $doc.Range(3, 6).Font.Size = 30
        $doc.Paragraphs.Item(1).Style = 'Title'
    }

    'paragraph-style-clears-paragraph-format' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $paragraph = $doc.Paragraphs.Item(1)
        $paragraph.Alignment = 1
        $paragraph.SpaceAfter = 30
        $paragraph.LeftIndent = 20
        $paragraph.KeepWithNext = -1
        $paragraph.PageBreakBefore = -1
        $paragraph.Style = 'Heading 1'
    }

    'paragraph-style-keeps-some-format' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 7).Font.Superscript = 1
        $doc.Range(0, 7).Font.Scaling = 150
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-underline-color' = {
        param($doc)
        $doc.Content.InsertAfter('0123456789')
        $doc.Range(0, 7).Font.Underline = 1
        $doc.Range(0, 7).Font.UnderlineColor = 255
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'paragraph-style-keeps-a-section' = {
        param($doc)
        $doc.Content.InsertAfter('First')
        [void]$doc.Sections.Add()
        $doc.Paragraphs.Item(1).Alignment = 1
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
    }

    'normal-style-applied'            = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Paragraphs.Item(1).Style = -1  # wdStyleNormal
    }

    'insert-paragraph-after-formatted' = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        $doc.Range(0, 4).Font.Bold = $true
        $doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
    }

    'insert-paragraph-before-formatted' = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        $doc.Range(0, 4).Font.Bold = $true
        $doc.Paragraphs.Item(1).Range.InsertParagraphBefore()
    }

    'paragraphs-add'                  = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        [void]$doc.Paragraphs.Add()
    }

    'split-a-formatted-paragraph'     = {
        param($doc)
        $doc.Content.InsertAfter('One two')
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        $doc.Range(3, 3).InsertParagraphAfter()
    }

    'content-text-with-paragraphs'    = {
        param($doc)
        $doc.Content.Text = "One`rTwo"
    }

    'insert-before-with-a-paragraph'  = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Range(0, 0).InsertBefore("A`rB")
    }

    'join-formatted-paragraphs'       = {
        param($doc)
        $doc.Content.InsertAfter('First')
        $doc.Content.InsertParagraphAfter()
        $doc.Content.InsertAfter('Second')
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        $doc.Paragraphs.Item(2).Range.ParagraphFormat.Alignment = 2
        [void]$doc.Paragraphs.Item(1).Range.Characters.Last.Delete()
    }

    'delete-a-section-break'          = {
        param($doc)
        $doc.Content.InsertAfter('First')
        [void]$doc.Sections.Add()
        $doc.Content.InsertAfter('Second')
        $doc.Sections.Item(1).PageSetup.LeftMargin = 36
        [void]$doc.Sections.Item(1).Range.Characters.Last.Delete()
    }

    'breaks'                          = {
        param($doc)
        $doc.Content.InsertAfter('Line Page Column')
        $doc.Range(11, 11).InsertBreak(8)  # wdColumnBreak
        $doc.Range(5, 5).InsertBreak(7)  # wdPageBreak
        $doc.Range(0, 0).InsertBreak(6)  # wdLineBreak
    }

    'delete-text-within-a-run'        = {
        param($doc)
        $doc.Content.InsertAfter('One two three')
        [void]$doc.Range(3, 7).Delete()
    }

    'replace-text-within-a-run'       = {
        param($doc)
        $doc.Content.InsertAfter('One two three')
        $doc.Range(4, 7).Text = '2'
    }

    'field-from-code'                 = {
        param($doc)
        [void]$doc.Fields.Add($doc.Range(0, 0), -1, 'PAGE', $false)  # wdFieldEmpty
    }

    'field-within-text'               = {
        param($doc)
        $doc.Content.InsertAfter('Before after')
        [void]$doc.Fields.Add($doc.Range(7, 7), 31)  # wdFieldDate
    }

    'field-within-bold-text'          = {
        param($doc)
        $doc.Content.InsertAfter('Before after')
        $doc.Range(0, 12).Font.Bold = $true
        [void]$doc.Fields.Add($doc.Range(7, 7), 31)  # wdFieldDate
    }

    'insert-paragraph-after-middle'   = {
        param($doc)
        $doc.Content.InsertAfter('One')
        $doc.Content.InsertParagraphAfter()
        $doc.Content.InsertAfter('Two')
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        $doc.Paragraphs.Item(2).Range.ParagraphFormat.Alignment = 2
        $doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
    }

    'heading-with-format-insert-paragraph-after' = {
        param($doc)
        $doc.Content.InsertAfter('Heading')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
        $doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
    }

    'split-a-heading-at-its-end'      = {
        param($doc)
        $doc.Content.InsertAfter('Heading')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Range(7, 7).InsertParagraphAfter()
    }

    'split-a-heading-in-the-middle'   = {
        param($doc)
        $doc.Content.InsertAfter('Heading')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Range(3, 3).InsertParagraphAfter()
    }

    'page-break-at-paragraph-start'   = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Range(0, 0).InsertBreak(7)  # wdPageBreak
    }

    'page-break-at-paragraph-end'     = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Range(4, 4).InsertBreak(7)  # wdPageBreak
    }

    'column-break-at-paragraph-start' = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Range(0, 0).InsertBreak(8)  # wdColumnBreak
    }

    'column-break-at-paragraph-end'   = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Range(4, 4).InsertBreak(8)  # wdColumnBreak
    }

    'line-break-at-paragraph-end'     = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        $doc.Range(4, 4).InsertBreak(6)  # wdLineBreak
    }

    'field-date-without-merge-format' = {
        param($doc)
        [void]$doc.Fields.Add($doc.Range(0, 0), 31, '', $false)  # wdFieldDate
    }

    'field-page-with-merge-format'    = {
        param($doc)
        [void]$doc.Fields.Add($doc.Range(0, 0), 33)  # wdFieldPage
    }

    'field-date-from-code'            = {
        param($doc)
        [void]$doc.Fields.Add($doc.Range(0, 0), -1, 'DATE', $false)  # wdFieldEmpty
    }

    'table-2x4-word9-behavior'        = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 2, 4, 1, 0)  # wdWord9TableBehavior, wdAutoFitFixed
    }

    'table-autofit-content'           = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 1, 2, 0, 1)  # wdWord8TableBehavior, wdAutoFitContent
    }

    'table-autofit-window'            = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 1, 2, 0, 2)  # wdWord8TableBehavior, wdAutoFitWindow
    }

    'table-word9-autofit-content'     = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 1)
    }

    'table-word9-autofit-window'      = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 2)
    }

    'table-nested'                    = {
        param($doc)
        $outer = $doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 0)
        [void]$doc.Tables.Add($outer.Cell(1, 1).Range, 1, 2, 1, 0)
    }

    'table-in-text'                   = {
        param($doc)
        $doc.Content.InsertAfter('Before after')
        [void]$doc.Tables.Add($doc.Range(7, 7), 1, 1)
    }

    'cell-vertical-merge'             = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 3, 2, 1, 0)
        $table.Cell(1, 1).Merge($table.Cell(2, 1))
    }

    'cell-merge-block'                = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 3, 3, 1, 0)
        $table.Cell(1, 1).Merge($table.Cell(2, 2))
    }

    'cell-split'                      = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 0)
        $table.Cell(1, 1).Split(1, 2)
    }

    'row-properties'                  = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 4, 1, 1, 0)
        $table.Rows.Item(1).HeightRule = 1  # wdRowHeightAtLeast
        $table.Rows.Item(1).Height = 20
        $table.Rows.Item(2).HeightRule = 2  # wdRowHeightExactly
        $table.Rows.Item(2).Height = 20
        $table.Rows.Item(3).HeadingFormat = $true
        $table.Rows.Item(4).AllowBreakAcrossPages = $false
    }

    'rows-add'                        = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 0)
        $table.Cell(1, 1).Range.Text = 'a'
        [void]$table.Rows.Add()
    }

    'columns-add'                     = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 0)
        [void]$table.Columns.Add()
    }

    'column-delete'                   = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 1, 3, 1, 0)
        $table.Columns.Item(2).Delete()
    }

    'column-width-narrower'           = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.Columns.Item(1).Width = 100
    }

    'column-width-wider'              = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.Columns.Item(1).Width = 200
    }

    'column-width-preferred-points'   = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.PreferredWidthType = 3  # wdPreferredWidthPoints
        $table.PreferredWidth = 300
        $table.Columns.Item(1).Width = 200
    }

    'column-width-after-add'          = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        [void]$table.Columns.Add()
        $table.Columns.Item(1).Width = 120
        $table.Columns.Item(4).Width = 70
    }

    'columns-add-within-space'        = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.Columns.Item(1).Width = 40
        $table.Columns.Item(2).Width = 40
        [void]$table.Columns.Add()
    }

    'columns-add-preferred-points'    = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.PreferredWidthType = 3  # wdPreferredWidthPoints
        $table.PreferredWidth = 300
        [void]$table.Columns.Add()
    }

    'column-delete-preferred-points'  = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.PreferredWidthType = 3  # wdPreferredWidthPoints
        $table.PreferredWidth = 300
        $table.Columns.Item(2).Delete()
    }

    'table-preferred-percent'         = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 2, 3, 1, 0)
        $table.PreferredWidthType = 2  # wdPreferredWidthPercent
        $table.PreferredWidth = 80
    }

    'table-borders-enable'            = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 1, 1)
        $table.Borders.Enable = $true
    }

    'table-properties'                = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 1, 2, 1, 0)
        $table.PreferredWidthType = 3  # wdPreferredWidthPoints
        $table.PreferredWidth = 200
        $table.Rows.Alignment = 1  # wdAlignRowCenter
        $table.Rows.LeftIndent = 0
        $table.Cell(1, 1).Shading.BackgroundPatternColor = 255
        $table.Cell(1, 2).VerticalAlignment = 1  # wdCellAlignVerticalCenter
    }

    'section-properties'              = {
        param($doc)
        $setup = $doc.Sections.Item(1).PageSetup
        $setup.Orientation = 1  # wdOrientLandscape
        $setup.Orientation = 0  # wdOrientPortrait
        $setup.SectionStart = 0  # wdSectionContinuous
        $setup.SectionStart = 2  # wdSectionNewPage
    }

    'section-landscape'               = {
        param($doc)
        $doc.Sections.Item(1).PageSetup.Orientation = 1
    }

    'section-continuous'              = {
        param($doc)
        $doc.Sections.Item(1).PageSetup.SectionStart = 0
    }

    'line-numbers'                    = {
        param($doc)
        $doc.Sections.Item(1).PageSetup.LineNumbering.Active = $true
    }

    'text-columns'                    = {
        param($doc)
        $doc.Sections.Item(1).PageSetup.TextColumns.SetCount(2)
    }

    'different-first-page'            = {
        param($doc)
        $doc.Sections.Item(1).PageSetup.DifferentFirstPageHeaderFooter = $true
    }

    'page-numbers-add'                = {
        param($doc)
        [void]$doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add()
    }

    'two-sections-with-headers'       = {
        param($doc)
        $doc.Sections.Item(1).Headers.Item(1).Range.Text = 'Header'
        $doc.Content.InsertAfter('First')
        [void]$doc.Sections.Add()
    }

    'header-unlink'                   = {
        param($doc)
        $doc.Sections.Item(1).Headers.Item(1).Range.Text = 'Header'
        [void]$doc.Sections.Add()
        $doc.Sections.Item(2).Headers.Item(1).LinkToPrevious = $false
    }

    'comment-only'                    = {
        param($doc)
        $doc.Content.InsertAfter('Commented text')
        [void]$doc.Comments.Add($doc.Range(0, 9), 'A comment')
    }

    'comment-on-insertion-point'      = {
        param($doc)
        $doc.Content.InsertAfter('Text')
        [void]$doc.Comments.Add($doc.Range(2, 2), 'A comment')
    }

    'comment-deleted'                 = {
        param($doc)
        $doc.Content.InsertAfter('Commented text')
        $comment = $doc.Comments.Add($doc.Range(0, 9), 'A comment')
        [void]$doc.Comments.Add($doc.Range(10, 14), 'Another')
        $comment.DeleteRecursively()
    }

    'page-numbers-add-left'           = {
        param($doc)
        [void]$doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add(0)  # wdAlignPageNumberLeft
    }

    'page-numbers-add-center'         = {
        param($doc)
        [void]$doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add(1)  # wdAlignPageNumberCenter
    }

    'page-numbers-add-to-text'        = {
        param($doc)
        $doc.Sections.Item(1).Footers.Item(1).Range.Text = 'Footer'
        [void]$doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add()
    }

    'comment-in-empty-paragraph'      = {
        param($doc)
        [void]$doc.Comments.Add($doc.Range(0, 0), 'A comment')
    }

    'heading-rows-off'                = {
        param($doc)
        $table = $doc.Tables.Add($doc.Range(0, 0), 3, 1, 1, 0)
        $table.Rows.Item(2).HeadingFormat = $true
        $table.Rows.Item(1).HeadingFormat = $false
    }

    'join-paragraphs'                 = {
        param($doc)
        $doc.Content.InsertAfter("First paragraph`rSecond paragraph")
        $paragraph = $doc.Paragraphs.Item(1).Range
        [void]$paragraph.Characters.Last.Delete()
    }
}
