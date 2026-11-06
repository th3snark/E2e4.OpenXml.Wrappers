<#
.SYNOPSIS
The scenarios New-WordReferences.ps1 runs in Word. Each is a scriptblock of Word object model calls on a new document (Documents.Add()).
.DESCRIPTION
The name is the folder the snapshot goes to under OpenXml.Wrappers.Test/TestFiles/WordReference, and MarkupConformanceTest has a case with
the same name that does the same thing through the library. Word's indexes are 1-based; the library's are 0-based. Keep the two in step.
#>

# Twenty three-letter words, one for each font property below: word i is at positions 4i to 4i + 3.
$script:FontWords = (1..20 | ForEach-Object { 'w{0:00}' -f $_ }) -join ' '

# ContentControl.SetPlaceholderText takes a building block, a range and text, of which Word needs the first two to be null objects (an omitted
# or $null argument is a type mismatch): wrapped in a DispatchWrapper they are.
function Set-PlaceholderText($control, $range, [string]$text) {
    $none = [Activator]::CreateInstance([System.Runtime.InteropServices.DispatchWrapper], @($null))
    $first = if ($null -ne $range) { $none } else { $none }
    $second = if ($null -ne $range) { $range } else { $none }
    [void][System.__ComObject].InvokeMember('SetPlaceholderText', [Reflection.BindingFlags]::InvokeMethod, $null, $control, @($first, $second, $text))
}
# The picture the inline picture scenarios insert (40 x 20 pixels at 96 dpi).
$script:Picture = (Resolve-Path (Join-Path $PSScriptRoot '..\..\OpenXml.Wrappers.Test\TestFiles\Pictures\blue-40x20.png')).Path

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
    'cells-borders-table'             = {
        param($doc)
        [void]$doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $doc.Tables.Item(1).Range.Cells.Borders.Enable = $true
    }
    'cells-borders-row'               = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $cells = $t.Rows.Item(2).Range.Cells
        $cells.Borders.OutsideLineStyle = 1
        $cells.Borders.InsideLineStyle = 1
    }

    'rows-borders'                    = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Borders.OutsideLineStyle = 1
        $t.Rows.Borders.InsideLineStyle = 1
    }
    'columns-borders'                 = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Columns.Borders.OutsideLineStyle = 1
        $t.Columns.Borders.InsideLineStyle = 1
    }
    'cells-merge-row'                 = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        $t.Rows.Item(2).Cells.Merge()
    }
    'cells-merge-block'               = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        $t.Cell(1, 1).Merge($t.Cell(2, 2))
    }
    'table-convert-to-text-tabs'      = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        [void]$t.ConvertToText(1)
    }
    'table-convert-to-text-paragraphs' = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        [void]$t.ConvertToText(0)
    }
    'table-convert-to-text-commas'    = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        [void]$t.ConvertToText(';')
    }
    'table-convert-to-text-commas-enum' = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        [void]$t.ConvertToText(2)
    }
    'table-convert-to-text-default'   = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        [void]$t.ConvertToText(3)
    }
    'row-convert-to-text'             = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        [void]$t.Rows.Item(2).ConvertToText(1)
    }
    'range-convert-to-table-tabs'     = {
        param($doc)
        $doc.Content.InsertAfter("a`tb`tc`rd`te`tf")
        [void]$doc.Content.ConvertToTable(1)
    }
    'range-convert-to-table-paragraphs' = {
        param($doc)
        $doc.Content.InsertAfter("a`rb`rc`rd")
        [void]$doc.Content.ConvertToTable(0, [Type]::Missing, 2)
    }
    'table-sort-text'                 = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 4, 2)
        $v = @('b', '10', 'c', '2', 'a', '1', 'd', '20')
        $i = 0
        foreach ($c in $t.Range.Cells) { $c.Range.Text = $v[$i]; $i++ }
        $t.Sort($false, 1)
    }
    'table-sort-numeric-header'       = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 4, 2)
        $v = @('h', 'n', 'c', '10', 'b', '2', 'a', '1')
        $i = 0
        foreach ($c in $t.Range.Cells) { $c.Range.Text = $v[$i]; $i++ }
        $t.Sort($true, 2, 1, 1)
    }
    'cell-formula-sum-above'          = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 2)
        $t.Cell(1, 1).Range.Text = '1'
        $t.Cell(2, 1).Range.Text = '2'
        $t.Cell(3, 1).Formula('=SUM(ABOVE)', '0.00')
    }
    'cell-autosum'                    = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 2)
        $t.Cell(1, 1).Range.Text = '1'
        $t.Cell(2, 1).Range.Text = '2'
        $t.Cell(3, 1).Formula()
    }
    'row-left-indent'                 = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).LeftIndent = 20
    }
    'row-set-left-indent'             = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).SetLeftIndent(30, 0)
    }
    'row-set-left-indent-proportional' = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).SetLeftIndent(30, 1)
    }
    'row-set-left-indent-first-column' = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).SetLeftIndent(30, 2)
    }
    'row-set-left-indent-same-width'  = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).SetLeftIndent(30, 3)
    }
    'cell-delete-up'                  = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        $t.Cell(2, 2).Delete(1)
    }
    'cells-delete-row-up'             = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $i = 1
        foreach ($c in $t.Range.Cells) { $c.Range.Text = "c$i"; $i++ }
        $t.Rows.Item(2).Cells.Delete(1)
    }
    'cells-add-end'                   = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).Cells.Add()
    }
    'cells-add-before'                = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).Cells.Add($t.Cell(2, 2))
    }
    'cells-split-columns'             = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).Cells.Split(1, 2)
    }
    'cells-split-rows'                = {
        param($doc)
        $t = $doc.Tables.Add($doc.Range(0, 0), 3, 3)
        $t.Rows.Item(2).Cells.Split(2, 1, $false)
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

    'page-border-options'             = {
        param($doc)
        $borders = $doc.Sections.Item(1).Borders
        $borders.Enable = $true
        $borders.AlwaysInFront = $false
        $borders.DistanceFrom = 0
        $borders.EnableFirstPageInSection = $false
        $borders.SurroundHeader = $false
        $borders.SurroundFooter = $false
    }

    'page-setup-direction'            = {
        param($doc)
        $doc.PageSetup.TogglePortrait()
        $doc.PageSetup.SectionDirection = 0  # wdSectionDirectionRtl
    }

    'page-numbers-add-inside'         = {
        param($doc)
        [void]$doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add(3)  # wdAlignPageNumberInside
    }

    'page-numbers-add-outside'        = {
        param($doc)
        [void]$doc.Sections.Item(1).Footers.Item(1).PageNumbers.Add(4)  # wdAlignPageNumberOutside
    }

    'page-numbers-chapter'            = {
        param($doc)
        $numbers = $doc.Sections.Item(1).Footers.Item(1).PageNumbers
        $numbers.IncludeChapterNumber = $true
        $numbers.LevelForChapter = 2
        $numbers.ChapterPageSeparator = 2  # wdSeparatorColon
        $numbers.ShowFirstPageNumber = $false
    }

    'page-numbers-chapter-default'    = {
        param($doc)
        $numbers = $doc.Sections.Item(1).Footers.Item(1).PageNumbers
        $numbers.IncludeChapterNumber = $true
        $numbers.ChapterPageSeparator = 1  # wdSeparatorPeriod
        $numbers.ChapterPageSeparator = 0  # wdSeparatorHyphen
    }

    'text-columns-add'                = {
        param($doc)
        $columns = $doc.PageSetup.TextColumns
        [void]$columns.Add(72, 10)
        [void]$columns.Add(100, 20)
        [void]$columns.Add(90)
    }


    'text-columns-add-no-width'       = {
        param($doc)
        $columns = $doc.PageSetup.TextColumns
        [void]$columns.Add()
    }

    'text-columns-items'              = {
        param($doc)
        $columns = $doc.PageSetup.TextColumns
        $columns.SetCount(3)
        $columns.Item(1).Width = 100
        $columns.Item(2).SpaceAfter = 10
    }

    'text-columns-flow'               = {
        param($doc)
        $doc.PageSetup.TextColumns.SetCount(2)
        $doc.PageSetup.TextColumns.FlowDirection = 1  # wdFlowRtl
    }


    'section-protected-for-forms'     = {
        param($doc)
        $doc.Content.InsertAfter('a')
        [void]$doc.Sections.Add()
        $doc.Sections.Item(1).ProtectedForForms = $false
        $doc.Sections.Item(2).ProtectedForForms = $false
        $doc.Sections.Item(2).ProtectedForForms = $true
    }

    'fields-unlink'                   = {
        param($doc)
        $doc.Content.InsertAfter('text')
        [void]$doc.Fields.Add($doc.Range(0, 0), -1, 'SEQ a', $false)
        [void]$doc.Fields.Add($doc.Range(0, 0), -1, 'SEQ b', $false)
        $doc.Fields.Unlink()
    }

    'document-convert'                = {
        param($doc)
        $doc.SetCompatibilityMode(14)
        $doc.Convert()
    }

    'hyperlink-add'                   = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.com/')
    }

    'hyperlink-add-place-only'        = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Bookmarks.Add('Mark', $doc.Range(6, 11))
        [void]$doc.Hyperlinks.Add($doc.Range(0, 5), '', 'Mark', 'Go to mark', 'Display')
    }

    'hyperlink-add-all-arguments'     = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.org/page', 'sec1', 'Tip', 'Shown', '_blank')
    }

    'hyperlink-add-formatted'         = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Range(0, 5).Font.Bold = $true
        $doc.Range(3, 5).Font.Italic = $true
        [void]$doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.com/')
    }

    'hyperlink-add-collapsed-text'    = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Hyperlinks.Add($doc.Range(5, 5), 'https://example.com/', '', '', 'Inserted')
    }

    'hyperlink-add-collapsed'         = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Hyperlinks.Add($doc.Range(5, 5), 'https://example.com/')
    }

    'hyperlink-set-text'              = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $link = $doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.com/')
        $link.TextToDisplay = 'Longer replacement'
    }

    'hyperlink-edit'                  = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $link = $doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.com/')
        $link.Address = 'https://changed.example/'
        $link.ScreenTip = 'A tip'
        $link.SubAddress = 'Mark2'
        $link.Target = '_top'
    }

    'hyperlink-email-subject'         = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $link = $doc.Hyperlinks.Add($doc.Range(0, 5), 'mailto:a@b.com')
        $link.EmailSubject = 'Hi there'
    }

    'hyperlink-delete'                = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.com/')
        $doc.Hyperlinks.Item(1).Delete()
    }

    'hyperlink-two-same-address'      = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Hyperlinks.Add($doc.Range(12, 17), 'https://example.com/')
        [void]$doc.Hyperlinks.Add($doc.Range(0, 5), 'https://example.com/')
    }
    'list-defaults'                   = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Paragraphs.Item(1).Range.ListFormat.ApplyBulletDefault()
        $doc.Paragraphs.Item(2).Range.ListFormat.ApplyNumberDefault()
        $doc.Paragraphs.Item(3).Range.ListFormat.ApplyOutlineNumberDefault()
    }

    'list-apply-bullet'               = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyBulletDefault()
    }

    'list-apply-number'               = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyNumberDefault()
    }

    'list-apply-outline'              = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyOutlineNumberDefault()
    }

    'list-apply-on-heading'           = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Paragraphs.Item(1).Range.Style = 'Heading 1'
        $doc.Paragraphs.Item(1).Range.ListFormat.ApplyBulletDefault()
    }

    'list-apply-on-indented'          = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Paragraphs.Item(1).Range.ParagraphFormat.LeftIndent = 72
        $doc.Paragraphs.Item(1).Range.ListFormat.ApplyNumberDefault()
    }

    'list-indent-outdent'             = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyOutlineNumberDefault()
        $doc.Paragraphs.Item(2).Range.ListFormat.ListIndent()
        $doc.Paragraphs.Item(3).Range.ListFormat.ListIndent()
        $doc.Paragraphs.Item(3).Range.ListFormat.ListOutdent()
    }

    'list-outdent-first-level'        = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyOutlineNumberDefault()
        $doc.Paragraphs.Item(1).Range.ListFormat.ListOutdent()
    }

    'list-level-number'               = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyOutlineNumberDefault()
        $doc.Paragraphs.Item(3).Range.ListFormat.ListLevelNumber = 4
    }

    'list-remove-numbers'             = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyOutlineNumberDefault()
        $doc.Paragraphs.Item(2).Range.ListFormat.ListLevelNumber = 2
        $doc.Paragraphs.Item(2).Range.ListFormat.RemoveNumbers()
    }

    'list-convert-bullet-to-text'     = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyBulletDefault()
        $doc.Paragraphs.Item(2).Range.ListFormat.ConvertNumbersToText()
    }

    'list-convert-number-to-text'     = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyNumberDefault()
        $doc.Paragraphs.Item(1).Range.ListFormat.ConvertNumbersToText()
    }

    'list-apply-template-continue'    = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Paragraphs.Item(1).Range.ListFormat.ApplyNumberDefault()
        $doc.Paragraphs.Item(2).Range.ListFormat.ApplyListTemplate($doc.ListTemplates.Item(1), $true)
    }
    'list-template-add'               = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $t = $doc.ListTemplates.Add($false, 'Mine')
        Write-Host ('template: name "{0}" outline {1} levels {2}' -f $t.Name, $t.OutlineNumbered, $t.ListLevels.Count)
        $l = $t.ListLevels.Item(1)
        Write-Host ('level1 default: fmt "{0}" style {1} numpos {2} textpos {3} tab {4} trail {5} start {6}' -f $l.NumberFormat, $l.NumberStyle, $l.NumberPosition, $l.TextPosition, $l.TabPosition, $l.TrailingCharacter, $l.StartAt)
        $doc.Content.ListFormat.ApplyListTemplate($t)
        Write-Host ('count templates {0}' -f $doc.ListTemplates.Count)
    }

    'list-template-add-outline'       = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo")
        $t = $doc.ListTemplates.Add($true, 'Outline')
        Write-Host ('template: name "{0}" outline {1} levels {2}' -f $t.Name, $t.OutlineNumbered, $t.ListLevels.Count)
        $l = $t.ListLevels.Item(2)
        Write-Host ('level2 default: fmt "{0}" style {1} numpos {2} textpos {3} tab {4}' -f $l.NumberFormat, $l.NumberStyle, $l.NumberPosition, $l.TextPosition, $l.TabPosition)
        $doc.Content.ListFormat.ApplyListTemplate($t)
    }

    'list-template-add-unnamed'       = {
        param($doc)
        $doc.Content.InsertAfter('One')
        $t = $doc.ListTemplates.Add()
        Write-Host ('unnamed: name "{0}" outline {1}' -f $t.Name, $t.OutlineNumbered)
    }

    'list-level-font-tab'             = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo")
        $t = $doc.ListTemplates.Add($false, 'Styled')
        $l = $t.ListLevels.Item(1)
        $l.NumberFormat = '%1.'
        $l.NumberStyle = 0
        $l.Font.Bold = 1
        $l.Font.Size = 14
        $l.Font.Name = 'Arial'
        $l.Font.Color = 255
        $l.TabPosition = 100
        Write-Host ('font: bold {0} size {1} name {2} color {3}; tab {4}' -f $l.Font.Bold, $l.Font.Size, $l.Font.Name, $l.Font.Color, $l.TabPosition)
        $doc.Content.ListFormat.ApplyListTemplate($t)
    }

    'list-apply-with-level'           = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Paragraphs.Item(1).Range.ListFormat.ApplyNumberDefault()
        $t = $doc.ListTemplates.Item(1)
        $r = $doc.Range($doc.Paragraphs.Item(2).Range.Start, $doc.Paragraphs.Item(3).Range.End)
        Write-Host ('can continue: {0}' -f $r.ListFormat.CanContinuePreviousList($t))
        $r.ListFormat.ApplyListTemplateWithLevel($t, $true, 0, [Type]::Missing, 2)
        Write-Host ('levels: {0} {1} {2}' -f $doc.Paragraphs.Item(1).Range.ListFormat.ListLevelNumber, $doc.Paragraphs.Item(2).Range.ListFormat.ListLevelNumber, $doc.Paragraphs.Item(3).Range.ListFormat.ListLevelNumber)
    }

    #list-can-continue prints what Word does and makes no snapshot: CanContinuePreviousList makes Word add a list for the template it is asked about.
    'list-can-continue'               = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree`rFour`rFive")
        $p = { param($i) $doc.Paragraphs.Item($i).Range }
        $t = $doc.ListTemplates.Add($false, 'A')
        Write-Host ('no list before: {0}' -f (& $p 1).ListFormat.CanContinuePreviousList($t))
        (& $p 1).ListFormat.ApplyListTemplate($t)
        Write-Host ('after para1 numbered by t: para2 {0}' -f (& $p 2).ListFormat.CanContinuePreviousList($t))
        $u = $doc.ListTemplates.Add($false, 'B')
        Write-Host ('other template: para2 {0}' -f (& $p 2).ListFormat.CanContinuePreviousList($u))
        Write-Host ('gap: para3 {0}' -f (& $p 3).ListFormat.CanContinuePreviousList($t))
        Write-Host ('para1 itself: {0}' -f (& $p 1).ListFormat.CanContinuePreviousList($t))
        (& $p 4).ListFormat.ApplyListTemplate($t, $false)
        Write-Host ('para5 after para4 reset: {0}' -f (& $p 5).ListFormat.CanContinuePreviousList($t))
    }

    #The three list-on-list scenarios print what Word does and make no snapshot: touching Doc.Lists makes Word name a template (A2), which the library
    #does not do. ListWrapperTest replays the printed values.
    'list-apply-template-in-use'      = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $t = $doc.ListTemplates.Add($false, 'A')
        $doc.Paragraphs.Item(1).Range.ListFormat.ApplyListTemplate($t)
        $doc.Paragraphs.Item(2).Range.ListFormat.ApplyListTemplate($t, $false)
        $doc.Paragraphs.Item(3).Range.ListFormat.ApplyListTemplate($t, $false)
    }

    'list-on-list'                    = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyNumberDefault()
        $l = $doc.Lists.Item(1)
        $t = $doc.ListTemplates.Add($false, 'Z')
        $t.ListLevels.Item(1).NumberFormat = '(%1)'
        $l.ApplyListTemplate($t)
        Write-Host ('after apply: lists {0} strings {1} {2} {3}' -f $doc.Lists.Count, $doc.Paragraphs.Item(1).Range.ListFormat.ListString, $doc.Paragraphs.Item(2).Range.ListFormat.ListString, $doc.Paragraphs.Item(3).Range.ListFormat.ListString)
    }

    'list-on-list-remove'             = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyNumberDefault()
        $doc.Lists.Item(1).RemoveNumbers()
        Write-Host ('after remove: lists {0}' -f $doc.Lists.Count)
    }

    'list-on-list-convert'            = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        $doc.Content.ListFormat.ApplyNumberDefault()
        $doc.Lists.Item(1).ConvertNumbersToText()
        Write-Host ('after convert: lists {0}; text "{1}"' -f $doc.Lists.Count, $doc.Content.Text.Replace("`r", '|').Replace("`t", '>'))
    }

    'selection-insert-rows-below'     = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 2)
        $t.Cell(1, 1).Range.Text = 'a'
        $t.Cell(2, 1).Range.Text = 'c'
        $t.Cell(1, 1).Select()
        $doc.Application.Selection.InsertRowsBelow(2)
    }

    'selection-insert-rows-above'     = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 2)
        $t.Cell(1, 1).Range.Text = 'a'
        $t.Cell(2, 1).Range.Text = 'c'
        $t.Cell(2, 1).Select()
        $doc.Application.Selection.InsertRowsAbove(1)
    }

    # Printing-only (no snapshot): what the selection is after inserting columns, extending it, and moving to fields. The unit tests replay the values.
    'selection-insert-columns'        = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 3)
        for ($r = 1; $r -le 2; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
        $t.Cell(1, 2).Select()
        $doc.Application.Selection.InsertColumns()
    }

    'selection-insert-columns-right'  = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 3)
        for ($r = 1; $r -le 2; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
        $t.Cell(1, 2).Select()
        $doc.Application.Selection.InsertColumnsRight()
    }

    'selection-insert-columns-two'    = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 3)
        for ($r = 1; $r -le 2; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
        $doc.Application.Selection.SetRange($t.Cell(1, 2).Range.Start, $t.Cell(1, 3).Range.End)
        $doc.Application.Selection.InsertColumns()
    }

    'selection-insert-rows'           = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 3, 2)
        for ($r = 1; $r -le 3; $r++) { for ($c = 1; $c -le 2; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
        $t.Cell(2, 1).Select()
        $doc.Application.Selection.InsertRows(2)
    }

    'selection-extend-case'           = {
        param($doc)
        $doc.Content.InsertAfter("one Two three Four")
        $sel = $doc.Application.Selection
        function Ext($s, $u) { Add-Type -AssemblyName Microsoft.VisualBasic; [void][Microsoft.VisualBasic.Interaction]::CallByName($s, 'Extend', [Microsoft.VisualBasic.CallType]::Method, $u) }
        $sel.SetRange(0, 0); Ext $sel 't'; Write-Host ("extend to t from 0: {0}-{1}" -f $sel.Start, $sel.End)
        $sel.SetRange(0, 0); Ext $sel 'T'; Write-Host ("extend to T from 0: {0}-{1}" -f $sel.Start, $sel.End)
        $sel.SetRange(0, 0); Ext $sel 'tw'; Write-Host ("extend to 'tw' from 0: {0}-{1}" -f $sel.Start, $sel.End)
        $sel.SetRange(0, 3); Ext $sel 'o'; Write-Host ("extend to o from 0-3: {0}-{1}" -f $sel.Start, $sel.End)
    }

    'selection-insert-columns-behaviour' = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 3)
        for ($r = 1; $r -le 2; $r++) { for ($c = 1; $c -le 3; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
        $sel = $doc.Application.Selection
        function Show($label) {
            $t = $doc.Tables.Item(1)
            Write-Host ('{0}: cols {1}; sel {2}-{3} type {4}; row1 "{5}"; widths {6}' -f $label, $t.Columns.Count, $sel.Start, $sel.End, $sel.Type, (($t.Rows.Item(1).Cells | ForEach-Object { $_.Range.Text.TrimEnd([char]13, [char]7) }) -join '|'), (($t.Rows.Item(1).Cells | ForEach-Object { $_.Width }) -join ','))
        }
        Show 'start'
        $t.Cell(1, 2).Select()
        $sel.InsertColumns()
        Show 'insert columns (cell 1,2)'
        $doc.Tables.Item(1).Cell(1, 2).Select()
        $sel.InsertColumnsRight()
        Show 'insert columns right (cell 1,2)'
        $doc.Tables.Item(1).Cell(1, 1).Range.Select()
        $sel.SetRange($doc.Tables.Item(1).Cell(1, 1).Range.Start, $doc.Tables.Item(1).Cell(1, 2).Range.End)
        Show 'two cells selected'
        $sel.InsertColumns()
        Show 'insert columns (two cells)'
        $doc.Tables.Item(1).Cell(1, 3).Select()
        $sel.SelectColumn()
        Show 'column selected'
        $sel.InsertColumnsRight()
        Show 'insert columns right (column)'
    }

    'selection-insert-rows-behaviour' = {
        param($doc)
        $doc.Content.InsertAfter("Before`r")
        $t = $doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 3, 2)
        for ($r = 1; $r -le 3; $r++) { for ($c = 1; $c -le 2; $c++) { $t.Cell($r, $c).Range.Text = "r$r" + "c$c" } }
        $sel = $doc.Application.Selection
        $t.Cell(2, 1).Select()
        $sel.InsertRows(2)
        $t = $doc.Tables.Item(1)
        Write-Host ('rows after InsertRows(2) on row 2: {0}; first column "{1}"' -f $t.Rows.Count, (($t.Columns.Item(1).Cells | ForEach-Object { $_.Range.Text.TrimEnd([char]13, [char]7) }) -join '|'))
        Write-Host ('selection {0}-{1} type {2}' -f $sel.Start, $sel.End, $sel.Type)
        $t.Cell(1, 1).Select(); $sel.InsertRowsAbove(1)
        Write-Host ('after InsertRowsAbove(1): selection {0}-{1} type {2}' -f $sel.Start, $sel.End, $sel.Type)
        $t.Cell(1, 1).Select(); $sel.InsertRowsBelow(1)
        Write-Host ('after InsertRowsBelow(1): selection {0}-{1} type {2}' -f $sel.Start, $sel.End, $sel.Type)
        $t = $doc.Tables.Item(1)
        $sel.SetRange($t.Cell(1, 1).Range.Start, $t.Cell(2, 2).Range.End)
        Write-Host ('two rows selected: {0}-{1} type {2}' -f $sel.Start, $sel.End, $sel.Type)
        $sel.InsertRows(1)
        $t = $doc.Tables.Item(1)
        Write-Host ('after InsertRows(1) on two rows: rows {0}; selection {1}-{2} type {3}' -f $t.Rows.Count, $sel.Start, $sel.End, $sel.Type)
        $sel.SetRange(2, 2)
        try { $sel.InsertRows(2); Write-Host ('InsertRows outside a table: ok; tables {0}; rows {1}' -f $doc.Tables.Count, $doc.Tables.Item(1).Rows.Count) } catch { Write-Host ('InsertRows outside a table: ' + $_.Exception.Message) }
        $doc.Content.InsertParagraphAfter(); $sel.SetRange(2, 2)
        try { $sel.InsertColumns(); Write-Host 'InsertColumns outside a table: ok' } catch { Write-Host ('InsertColumns outside a table: ' + $_.Exception.Message) }
    }

    'selection-extend-units'          = {
        param($doc)
        $doc.Content.InsertAfter("One two three. Four five six.`rSecond paragraph here.`rThird.")
        $sel = $doc.Application.Selection
        function Ext($s, $u) { Add-Type -AssemblyName Microsoft.VisualBasic; [void][Microsoft.VisualBasic.Interaction]::CallByName($s, 'Extend', [Microsoft.VisualBasic.CallType]::Method, $u) }
        function Show($label) { Write-Host ('{0}: {1}-{2} mode {3} "{4}"' -f $label, $sel.Start, $sel.End, $sel.ExtendMode, ($sel.Text -replace "`r", '\r')) }
        $sel.SetRange(5, 5)
        Show 'start'
        for ($i = 0; $i -lt 6; $i++) { $sel.Extend(); Show "extend $i" }
        for ($i = 0; $i -lt 6; $i++) { $sel.Shrink(); Show "shrink $i" }
        $sel.ExtendMode = $false
        Show 'mode off'
        $sel.SetRange(5, 5)
        Ext $sel 'e'; Show "extend to 'e'"
        Ext $sel 'e'; Show "extend to 'e' again"
        Ext $sel '.'; Show "extend to '.'"
        Ext $sel "`r"; Show 'extend to paragraph mark'
        Ext $sel 'Q'; Show "extend to 'Q' (absent)"
        $sel.Shrink(); Show 'shrink after character extends'
        $sel.Collapse(1); Show 'collapse'
        $sel.SetRange(5, 8); $sel.Shrink(); Show 'shrink on 3 chars'
        $sel.SetRange(5, 5); $sel.ExtendMode = $true; Ext $sel 'z'; Show "extend mode on, extend to 'z'"; Ext $sel 'f'; Show "extend to 'f'"
        $sel.ExtendMode = $true; Show 'mode set on'; $sel.MoveRight(1, 1, 0) | Out-Null; Show 'moveright in extend mode'
        $sel.ExtendMode = $false
        $sel.SetRange(5, 5); $sel.Shrink(); Show 'shrink insertion point'
    }

    'selection-extend-mode'           = {
        param($doc)
        $doc.Content.InsertAfter("One two three. Four five six.`rSecond paragraph here.`rThird.")
        $sel = $doc.Application.Selection
        function Show($label) { Write-Host ('{0}: {1}-{2} mode {3} "{4}"' -f $label, $sel.Start, $sel.End, $sel.ExtendMode, ($sel.Text -replace "`r", '\r')) }
        $sel.SetRange(5, 5); $sel.ExtendMode = $true; $sel.MoveRight(1, 3, 0) | Out-Null; Show 'moveright 3 in mode'
        $sel.MoveLeft(1, 1, 0) | Out-Null; Show 'moveleft 1 in mode'
        $sel.MoveRight(2, 1, 0) | Out-Null; Show 'moveright word in mode'
        $sel.MoveDown(4, 1, 0) | Out-Null; Show 'movedown paragraph in mode'
        $sel.Collapse(1); Show 'collapse'
        $sel.SetRange(2, 6); Show 'setrange'
        $sel.Select(); Show 'range select'
        $sel.Range.Select(); Show 'range.select'
        $sel.TypeText('X'); Show 'typetext'
        $sel.ExtendMode = $true; $sel.Extend(); $sel.Extend(); Show 'extend twice'
        $sel.SetRange(8, 8); Show 'after setrange to ip'
        $sel.Extend(); Show 'extend'
        $sel.ExtendMode = $false; $sel.Extend(); Show 'extend from mode off (turns on)'
        $sel.ExtendMode = $false
        $sel.SetRange(6, 12); $sel.Extend(); Show 'extend from 6-12 (mode on)'
        $sel.ExtendMode = $false
        $sel.SetRange(16, 16); $sel.Extend(); $sel.Extend(); Show 'extend ip at 16 twice (sentence start)'
    }
    'selection-next-field'            = {
        param($doc)
        $doc.Content.InsertAfter("A ")
        $r = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        [void]$doc.Fields.Add($r, -1, 'DATE', $false)
        $doc.Content.InsertAfter(" B ")
        $r = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
        [void]$doc.Fields.Add($r, -1, 'PAGE', $false)
        $doc.Content.InsertAfter(" C")
        $sel = $doc.Application.Selection
        function Show($label, $f) { Write-Host ('{0}: {1}-{2} type {3} found {4}' -f $label, $sel.Start, $sel.End, $sel.Type, $(if ($f -eq $null) { 'nothing' } else { $f.Code.Text.Trim() })) }
        $sel.SetRange(0, 0)
        Show 'next from start' $sel.NextField()
        Show 'next again' $sel.NextField()
        Show 'next at last' $sel.NextField()
        Show 'previous' $sel.PreviousField()
        Show 'previous again' $sel.PreviousField()
        Show 'previous at first' $sel.PreviousField()
        $sel.SetRange(3, 3)
        Show 'next from inside first' $sel.NextField()
        $sel.SetRange($doc.Content.End - 1, $doc.Content.End - 1)
        Show 'previous from end' $sel.PreviousField()
        $sel.SetRange(2, 2)
        Show 'next from ip at field start' $sel.NextField()
        $sel.SetRange(3, 3)
        Show 'previous from ip inside first' $sel.PreviousField()
        $sel.SetRange(2, 2)
        Show 'previous from ip at field start' $sel.PreviousField()
        $sel.SetRange(21, 21)
        Show 'next from ip at field end' $sel.NextField()
        $sel.SetRange(0, 5)
        Show 'next from 0-5 spanning field start' $sel.NextField()
    }
    'selection-clear-character'       = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Range(0, 5).Style = 'Strong'
        $doc.Range(0, 5).Font.Size = 20
        $doc.Range(6, 11).Style = 'Strong'
        $doc.Range(6, 11).Font.Size = 20
        $doc.Range(0, 5).Select()
        $doc.Application.Selection.ClearCharacterStyle()
        $doc.Range(6, 11).Select()
        $doc.Application.Selection.ClearCharacterDirectFormatting()
    }

    'selection-clear-character-all'   = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Range(0, 5).Style = 'Strong'
        $doc.Range(0, 5).Font.Size = 20
        $doc.Range(0, 5).Select()
        $doc.Application.Selection.ClearCharacterAllFormatting()
    }

    'selection-clear-paragraph'       = {
        param($doc)
        $doc.Content.InsertAfter("One`rTwo`rThree")
        foreach ($i in 1, 2, 3) {
            $doc.Paragraphs.Item($i).Style = 'Heading 1'
            $doc.Paragraphs.Item($i).Range.ParagraphFormat.Alignment = 1
        }
        $doc.Paragraphs.Item(1).Range.Select()
        $doc.Application.Selection.ClearParagraphStyle()
        $doc.Paragraphs.Item(2).Range.Select()
        $doc.Application.Selection.ClearParagraphDirectFormatting()
        $doc.Paragraphs.Item(3).Range.Select()
        $doc.Application.Selection.ClearParagraphAllFormatting()
    }
    'selection-bold-run'              = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Range(0, 5).Select()
        $doc.Application.Selection.BoldRun()
        $doc.Range(6, 11).Select()
        $doc.Application.Selection.BoldRun()
        $doc.Application.Selection.BoldRun()
        $doc.Range(12, 17).Select()
        $doc.Application.Selection.ItalicRun()
    }

    'footnote-add'                    = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'The note text')
    }

    'footnote-add-to-range'           = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 10))
    }

    'footnote-add-custom-mark'        = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), 'ref*', 'custom mark')
    }

    'footnote-add-before-another'     = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(11, 11), [Type]::Missing, 'second in the file')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'first in the text')
    }

    'footnote-delete'                 = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Footnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.Item(1).Delete()
    }

    'endnote-add'                     = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Endnotes.Add($doc.Range(5, 5), [Type]::Missing, 'The end note')
    }

    'footnote-and-endnote'            = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'The end note')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'The footnote')
    }
    'note-option-footnote-style'      = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.NumberStyle = 2
    }

    'note-option-footnote-start'      = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.StartingNumber = 5
    }

    'note-option-footnote-restart-section' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.NumberingRule = 1
    }

    'note-option-footnote-restart-page' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.NumberingRule = 2
    }

    'note-option-footnote-location'   = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.Location = 1
    }

    'note-option-footnote-symbol'     = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.NumberStyle = 9
    }

    'note-option-endnote-section-end' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Endnotes.Location = 0
    }

    'note-option-endnote-decimal-start' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Endnotes.NumberStyle = 0
        $doc.Endnotes.StartingNumber = 3
    }

    'note-convert-footnotes'          = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.Convert()
    }

    'note-convert-endnotes'           = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Endnotes.Convert()
    }

    'note-swap'                       = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.Footnotes.Add($doc.Range(5, 5), [Type]::Missing, 'one')
        [void]$doc.Endnotes.Add($doc.Range(11, 11), [Type]::Missing, 'two')
        $doc.Footnotes.SwapWithEndnotes()
    }

    'content-control-placeholders'    = {
        param($doc)
        $doc.Content.InsertAfter("A`rB`rC")
        $types = 1, 4, 6
        for ($i = 1; $i -le 3; $i++) {
            $place = $doc.Paragraphs.Item($i).Range
            $place.Collapse(1)
            $doc.Application.Selection.EndKey(6)  # wdStory: Word adds to the selection too, which must not be in a control
            [void]$doc.ContentControls.Add($types[$i - 1], $place)
        }
    }
    'content-control-rich-text'       = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(0, $doc.Range(6, 11))
    }

    'content-control-text'            = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(1, $doc.Range(6, 11))
    }

    'content-control-combo-box'       = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(3, $doc.Range(6, 11))
    }

    'content-control-dropdown-list'   = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(4, $doc.Range(6, 11))
    }

    'content-control-date'            = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(6, $doc.Range(6, 11))
    }

    'content-control-check-box'       = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(8, $doc.Range(6, 6))
    }

    'content-control-collapsed-text'  = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(1, $doc.Range(5, 5))
    }

    'content-control-collapsed-date'  = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(6, $doc.Range(5, 5))
    }

    'content-control-collapsed-dropdown' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        [void]$doc.ContentControls.Add(4, $doc.Range(5, 5))
    }

    'content-control-properties'      = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(0, $doc.Range(6, 11))
        $cc.Title = 'My title'
        $cc.Tag = 'my-tag'
        $cc.LockContentControl = $true
        $cc.LockContents = $true
        $cc.Appearance = 2
    }

    'content-control-lock-control-only' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(0, $doc.Range(6, 11))
        $cc.LockContentControl = $true
    }

    'content-control-lock-contents-only' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(0, $doc.Range(6, 11))
        $cc.LockContents = $true
    }

    'content-control-delete-keep'     = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(0, $doc.Range(6, 11))
        $cc.Delete($false)
    }

    'content-control-delete-contents' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(0, $doc.Range(6, 11))
        $cc.Delete($true)
    }

    'content-control-check-box-checked' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(8, $doc.Range(6, 6))
        $cc.Checked = $true
    }

    'content-control-check-box-symbols' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(8, $doc.Range(6, 6))
        $cc.SetCheckedSymbol(10004, 'Segoe UI Symbol')
        $cc.SetUncheckedSymbol(10006, 'Segoe UI Symbol')
        $cc.Checked = $true
    }

    'content-control-dropdown-entries' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(4, $doc.Range(5, 5))
        [void]$cc.DropdownListEntries.Add('One', '1')
        [void]$cc.DropdownListEntries.Add('Two')
        [void]$cc.DropdownListEntries.Add('Zero', '0', 1)
    }

    'content-control-dropdown-select' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(4, $doc.Range(5, 5))
        [void]$cc.DropdownListEntries.Add('One', '1')
        [void]$cc.DropdownListEntries.Add('Two', '2')
        $cc.DropdownListEntries.Item(2).Select()
    }

    'content-control-dropdown-clear' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(4, $doc.Range(5, 5))
        [void]$cc.DropdownListEntries.Add('One', '1')
        [void]$cc.DropdownListEntries.Add('Two', '2')
        $cc.DropdownListEntries.Item(1).Delete()
        [void]$cc.DropdownListEntries.Add('Three', '3')
        $cc.DropdownListEntries.Item(1).MoveDown()
    }

    'content-control-combo-entries' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(3, $doc.Range(5, 5))
        [void]$cc.DropdownListEntries.Add('One', '1')
        [void]$cc.DropdownListEntries.Add('Two', '2')
        $cc.DropdownListEntries.Clear()
        [void]$cc.DropdownListEntries.Add('Only', 'only')
    }

    'content-control-date-format' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(6, $doc.Range(5, 5))
        $cc.DateDisplayFormat = 'dddd, d MMMM yyyy'
        $cc.DateStorageFormat = 1
        $cc.DateDisplayLocale = 1041
        $cc.DateCalendarType = 4
    }

    'content-control-multiline' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(1, $doc.Range(5, 5))
        $cc.MultiLine = $true
    }

    'content-control-color-temporary' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(0, $doc.Range(6, 11))
        $cc.Color = 255
        $cc.Temporary = $true
    }

    'content-control-default-text-style' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $cc = $doc.ContentControls.Add(1, $doc.Range(6, 11))
        $cc.DefaultTextStyle = 'Strong'
    }

    'custom-xml-part-add'             = {
        param($doc)
        $doc.Content.InsertAfter('Hello')
        [void]$doc.CustomXMLParts.Add('<root xmlns="urn:test"><name>Bob</name></root>')
    }

    'content-control-xml-mapping'     = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $part = $doc.CustomXMLParts.Add('<root xmlns="urn:test"><name>Bob</name></root>')
        $cc = $doc.ContentControls.Add(1, $doc.Range(5, 5))
        [void]$cc.XMLMapping.SetMapping('/ns:root/ns:name', "xmlns:ns='urn:test'", $part)
    }

    'content-control-xml-mapping-delete' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $part = $doc.CustomXMLParts.Add('<root xmlns="urn:test"><name>Bob</name></root>')
        $cc = $doc.ContentControls.Add(1, $doc.Range(5, 5))
        [void]$cc.XMLMapping.SetMapping('/ns:root/ns:name', "xmlns:ns='urn:test'", $part)
        $cc.XMLMapping.Delete()
    }

    'content-control-xml-node'        = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $doc.Application.Selection.EndKey(6)
        $part = $doc.CustomXMLParts.Add('<root xmlns="urn:test" id="7"><name>Bob</name><age>42</age></root>')
        $part.NamespaceManager.AddNamespace('ns', 'urn:test')
        $node = $part.SelectSingleNode('/ns:root/ns:name')
        Write-Host ('node: base "{0}" ns "{1}" text "{2}" xml "{3}" type {4} xpath "{5}" children {6} hasChildren {7} parent "{8}"' -f $node.BaseName, $node.NamespaceURI, $node.Text, $node.XML, $node.NodeType, $node.XPath, $node.ChildNodes.Count, $node.HasChildNodes, $node.ParentNode.BaseName)
        $root = $part.DocumentElement
        Write-Host ('root: base "{0}" children {1} attributes {2} first "{3}" last "{4}" attr "{5}"={6}' -f $root.BaseName, $root.ChildNodes.Count, $root.Attributes.Count, $root.FirstChild.BaseName, $root.LastChild.BaseName, $root.Attributes.Item(1).BaseName, $root.Attributes.Item(1).Text)
        Write-Host ('select: names {0}; all {1}; none {2}' -f $part.SelectNodes('//ns:name').Count, $part.SelectNodes('/ns:root/*').Count, $part.SelectNodes('//ns:zzz').Count)
        $cc = $doc.ContentControls.Add(1, $doc.Range(5, 5))
        $cc.XMLMapping.SetMappingByNode($node)
        Write-Host ('mapping: mapped {0} xpath "{1}" prefixes "{2}" node "{3}" text now "{4}"' -f $cc.XMLMapping.IsMapped, $cc.XMLMapping.XPath, $cc.XMLMapping.PrefixMappings, $cc.XMLMapping.CustomXMLNode.BaseName, $cc.Range.Text)
        $node.Text = 'Alice'
        Write-Host ('after set: cc text "{0}" part xml "{1}"' -f $cc.Range.Text, $part.XML)
        Write-Host ('linked: {0}; unlinked: {1}' -f $doc.SelectLinkedControls($node).Count, $doc.SelectUnlinkedControls($part).Count)
        $node.Delete()
        Write-Host ('after delete: part xml "{0}"' -f $part.XML)
    }

    'content-control-placeholder-text' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $tc = $doc.ContentControls.Add(1, $doc.Range(12, 17))
        $cc = $doc.ContentControls.Add(1, $doc.Range(5, 5))
        Set-PlaceholderText $cc $null 'Type here'
        Write-Host ('point: showing {0} text "{1}" placeholder name "{2}" value "{3}"' -f $cc.ShowingPlaceholderText, $cc.Range.Text, $cc.PlaceholderText.Name, $cc.PlaceholderText.Value)
        Set-PlaceholderText $tc $null 'Other'
        Write-Host ('around text: showing {0} text "{1}" placeholder "{2}"' -f $tc.ShowingPlaceholderText, $tc.Range.Text, $tc.PlaceholderText.Value)
        $tc.Range.Text = ''
        Write-Host ('emptied: showing {0} text "{1}"' -f $tc.ShowingPlaceholderText, $tc.Range.Text)
        Set-PlaceholderText $cc $null 'Changed'
        Write-Host ('changed: showing {0} text "{1}" value "{2}"' -f $cc.ShowingPlaceholderText, $cc.Range.Text, $cc.PlaceholderText.Value)
        Write-Host ('parent: {0}' -f $cc.Parent.GetType().FullName)
    }
    #content-control-placeholder-range prints what Word does and makes no snapshot: Word gives the entry a paragraph style of its own name.
    'content-control-placeholder-range' = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $cc = $doc.ContentControls.Add(1, $doc.Range(5, 5))
        Set-PlaceholderText $cc $doc.Range(6, 11) ''
        Write-Host ('from range: showing {0} text "{1}" value "{2}"' -f $cc.ShowingPlaceholderText, $cc.Range.Text, $cc.PlaceholderText.Value)
    }
    'custom-xml-node-edit'            = {
        param($doc)
        $doc.Content.InsertAfter('Hello')
        $part = $doc.CustomXMLParts.Add('<root xmlns="urn:test"><name>Bob</name></root>')
        $part.NamespaceManager.AddNamespace('ns', 'urn:test')
        $root = $part.DocumentElement
        [void]$root.AppendChildNode('age', 'urn:test', 1, '42')
        Write-Host ('append: {0}' -f $part.XML)
        [void]$root.AppendChildNode('note', '', 1, 'plain')
        Write-Host ('append no ns: {0}' -f $part.XML)
        [void]$root.AppendChildNode('id', '', 2, '7')
        Write-Host ('attribute: {0}' -f $part.XML)
        $root.AppendChildSubtree('<city xmlns="urn:test">Paris</city>')
        Write-Host ('subtree: {0}' -f $part.XML)
        $first = $part.SelectSingleNode('/ns:root/ns:name')
        [void]$root.InsertNodeBefore('first', 'urn:test', 1, 'x', $first)
        Write-Host ('before: {0}' -f $part.XML)
        [void]$root.ReplaceChildNode($first, 'surname', 'urn:test', 1, 'Smith')
        Write-Host ('replace: {0}' -f $part.XML)
        [void]$part.AddNode($root, 'extra', 'urn:test', 1, 'e')
        Write-Host ('addnode: {0}' -f $part.XML)
        $part.SelectSingleNode('/ns:root/ns:age').Text = '43'
        Write-Host ('final: {0}' -f $part.XML)
    }

    'custom-xml-part-delete'          = {
        param($doc)
        $doc.Content.InsertAfter('Hello')
        $part = $doc.CustomXMLParts.Add('<root xmlns="urn:test"><name>Bob</name></root>')
        $part.Delete()
    }

    'inline-picture-add'              = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        Write-Host ('size {0} x {1}, type {2}, lock {3}, alt "{4}", title "{5}"' -f $s.Width, $s.Height, $s.Type, $s.LockAspectRatio, $s.AlternativeText, $s.Title)
    }

    'inline-picture-twice'            = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        [void]$doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(0, 0))
    }

    'inline-picture-properties'       = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        $s.AlternativeText = 'A blue box'
        $s.Title = 'Box'
        $s.Width = 60
        Write-Host ('after width: {0} x {1}' -f $s.Width, $s.Height)
        $s.LockAspectRatio = $false
        $s.Height = 30
        Write-Host ('after unlocked height: {0} x {1}, scale {2} x {3}' -f $s.Width, $s.Height, $s.ScaleWidth, $s.ScaleHeight)
    }

    'inline-picture-scale'            = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        $s.ScaleWidth = 200
        Write-Host ('scale: {0} x {1}, scale {2} x {3}' -f $s.Width, $s.Height, $s.ScaleWidth, $s.ScaleHeight)
    }

    'inline-picture-lock'             = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        $s.LockAspectRatio = $true
        $s.Width = 60
        Write-Host ('locked: {0} x {1}, lock {2}' -f $s.Width, $s.Height, $s.LockAspectRatio)
    }

    'inline-picture-linked'           = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.InlineShapes.AddPicture($script:Picture, $true, $false, $doc.Range(5, 5))
        Write-Host ('linked type {0}' -f $s.Type)
    }

    'inline-picture-delete'           = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        $s.Delete()
    }

    'inline-picture-over-text'        = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        [void]$doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(6, 11))
    }

    'toc-add'                         = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rSecond`rMore text`rSub heading`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Heading 1'
        $doc.Paragraphs.Item(5).Style = 'Heading 2'
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0))
        Write-Host ('toc: count {0}, upper {1}, lower {2}, pages {3}, hyper {4}, right {5}, headings {6}, fields {7}, leader {8}, id "{9}"' -f $doc.TablesOfContents.Count, $toc.UpperHeadingLevel, $toc.LowerHeadingLevel, $toc.IncludePageNumbers, $toc.UseHyperlinks, $toc.RightAlignPageNumbers, $toc.UseHeadingStyles, $toc.UseFields, $toc.TabLeader, $toc.TableID)
    }

    'toc-added-styles'                = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rTitled`rMore text`rSub`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Title'
        $doc.Paragraphs.Item(5).Style = 'Subtitle'
        $m = [Type]::Missing
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0), $true, 1, 3, $m, $m, $m, $m, 'Title,1,Subtitle,2')
        Write-Host ('toc: headings {0}' -f $toc.HeadingStyles.Count)
        foreach ($h in $toc.HeadingStyles) { Write-Host ('  style "{0}" level {1}' -f $h.Style, $h.Level) }
    }

    #toc-heading-styles prints what Word does and makes no snapshot: after the last change Word wrote a bare o in the code.
    'toc-heading-styles'              = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rTitled`rMore text`rSub`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Title'
        $doc.Paragraphs.Item(5).Style = 'Subtitle'
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0))
        Write-Host ('before: headings {0}' -f $toc.HeadingStyles.Count)
        $h = $toc.HeadingStyles.Add('Title', 1)
        $toc.HeadingStyles.Add('Subtitle', 2) | Out-Null
        Write-Host ('after: headings {0}; first "{1}" level {2}; second "{3}" level {4}' -f $toc.HeadingStyles.Count, $toc.HeadingStyles.Item(1).Style, $toc.HeadingStyles.Item(1).Level, $toc.HeadingStyles.Item(2).Style, $toc.HeadingStyles.Item(2).Level)
        $toc.HeadingStyles.Item(2).Level = 3
        Write-Host ('changed level: {0}' -f $toc.HeadingStyles.Item(2).Level)
        $toc.HeadingStyles.Item(1).Delete()
        Write-Host ('after delete: headings {0}' -f $toc.HeadingStyles.Count)
        $toc.Update()
    }

    #toc-heading-styles-range makes no snapshot either: Word rebuilds the entries with formatting of its own on the end of the field.
    'toc-heading-styles-range'        = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rTitled`rMore text`rSub`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 2'
        $doc.Paragraphs.Item(3).Style = 'Title'
        $m = [Type]::Missing
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0), $true, 2, 4)
        Write-Host ('start: upper {0} lower {1}' -f $toc.UpperHeadingLevel, $toc.LowerHeadingLevel)
        $toc.HeadingStyles.Add('Title', 1) | Out-Null
        Write-Host ('after add: upper {0} lower {1} headings {2}' -f $toc.UpperHeadingLevel, $toc.LowerHeadingLevel, $toc.HeadingStyles.Count)
    }

    'toc-added-styles-nopages'        = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rTitled`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Title'
        $m = [Type]::Missing
        [void]$doc.TablesOfContents.Add($doc.Range(0, 0), $true, 1, 3, $false, $m, $false, $false, 'Title,2')
    }

    'toc-added-styles-outline'        = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rTitled`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Title'
        $m = [Type]::Missing
        [void]$doc.TablesOfContents.Add($doc.Range(0, 0), $true, 1, 3, $false, $m, $true, $true, 'Title,2', $m, $m, $true)
    }

    'tof-add'                         = {
        param($doc)
        $doc.Content.InsertAfter("Intro`rPicture one`rMiddle`rPicture two`rEnd")
        $doc.Paragraphs.Item(4).Range.InsertCaption('Figure', ': Second')
        $doc.Paragraphs.Item(2).Range.InsertCaption('Figure', ': First')
        $t = $doc.TablesOfFigures.Add($doc.Range(0, 0), 'Figure')
        Write-Host ('tof: count {0}, caption "{1}", include label {2}, pages {3}, hyper {4}, right {5}, id "{6}", leader {7}' -f $doc.TablesOfFigures.Count, $t.Caption, $t.IncludeLabel, $t.IncludePageNumbers, $t.UseHyperlinks, $t.RightAlignPageNumbers, $t.TableID, $t.TabLeader)
    }

    'tof-options'                     = {
        param($doc)
        $doc.Content.InsertAfter("Intro`rPicture one`rEnd")
        $doc.Paragraphs.Item(2).Range.InsertCaption('Figure', ': First')
        $m = [Type]::Missing
        [void]$doc.TablesOfFigures.Add($doc.Range(0, 0), 'Figure', $false, $m, $m, $m, $m, $m, $false, $false, $m, $true)
    }

    'toc-mark-entry'                  = {
        param($doc)
        $doc.Content.InsertAfter("Intro text here`rSecond bit")
        $m = [Type]::Missing
        [void]$doc.TablesOfContents.MarkEntry($doc.Range(16, 22), 'Second', $m, $m, 1)
        [void]$doc.TablesOfContents.MarkEntry($doc.Range(0, 5), 'Marked one', $m, 'A', 2)
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0), $false, 1, 3, $true, 'A')
        Write-Host ('toc: fields {0}, id "{1}"' -f $toc.UseFields, $toc.TableID)
    }

    'caption-add'                     = {
        param($doc)
        $doc.Content.InsertAfter("Intro`rPicture one`rEnd")
        $doc.Paragraphs.Item(2).Range.InsertCaption('Figure', ': First')
    }

    'caption-above'                   = {
        param($doc)
        $doc.Content.InsertAfter("Intro`rPicture one`rEnd")
        $doc.Paragraphs.Item(2).Range.InsertCaption('Figure', ': First', [Type]::Missing, 0)
    }

    'caption-exclude-label'           = {
        param($doc)
        $doc.Content.InsertAfter("Intro`rPicture one`rEnd")
        $doc.Paragraphs.Item(2).Range.InsertCaption('Figure', ': First', [Type]::Missing, 0, $true)
    }

    'toc-options'                     = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rSecond`rMore text`rSub heading`rEnd")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Heading 1'
        $doc.Paragraphs.Item(5).Style = 'Heading 2'
        [void]$doc.TablesOfContents.Add($doc.Range(0, 0), $true, 1, 1, $false, [Type]::Missing, $false, $false, [Type]::Missing, $false, $true)
    }

    'toc-hyperlinks'                  = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rSub heading")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Heading 2'
        [void]$doc.TablesOfContents.Add($doc.Range(0, 0), $true, 1, 3, $false, [Type]::Missing, $true, $true, [Type]::Missing, $true)
    }

    'toc-flags'                       = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rSub heading")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Heading 2'
        [void]$doc.TablesOfContents.Add($doc.Range(0, 0), $true, 2, 3, $true, 'A', $true, $true, [Type]::Missing, $true, $false, $true)
    }

    'toc-margins'                     = {
        param($doc)
        $doc.PageSetup.LeftMargin = 100
        $doc.Content.InsertAfter("First`rSome text")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        [void]$doc.TablesOfContents.Add($doc.Range(0, 0))
    }

    'toc-empty'                       = {
        param($doc)
        $doc.Content.InsertAfter("Plain`rtext")
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0))
        Write-Host ('empty toc range text: "{0}"' -f $toc.Range.Text)
    }



    'toc-delete'                      = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rSecond")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $doc.Paragraphs.Item(3).Style = 'Heading 1'
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0))
        $toc.Delete()
    }

    'toc-update'                      = {
        param($doc)
        $doc.Content.InsertAfter("First`rSome text`rSecond")
        $doc.Paragraphs.Item(1).Style = 'Heading 1'
        $toc = $doc.TablesOfContents.Add($doc.Range(0, 0))
        $doc.Paragraphs.Item($doc.Paragraphs.Count).Style = 'Heading 1'
        $toc.Update()
    }

    'shape-add-picture'               = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        Write-Host ('picture: type {0}, name "{1}", id {2}, left {3} top {4} w {5} h {6}, relH {7} relV {8}, wrap {9} side {10}, z {11}, lockanchor {12}, lockaspect {13}, visible {14}, rot {15}' -f $s.Type, $s.Name, $s.ID, $s.Left, $s.Top, $s.Width, $s.Height, $s.RelativeHorizontalPosition, $s.RelativeVerticalPosition, $s.WrapFormat.Type, $s.WrapFormat.Side, $s.ZOrderPosition, $s.LockAnchor, $s.LockAspectRatio, $s.Visible, $s.Rotation)
        Write-Host ('wrap distances: {0} {1} {2} {3} overlap {4}; anchor start {5}' -f $s.WrapFormat.DistanceTop, $s.WrapFormat.DistanceBottom, $s.WrapFormat.DistanceLeft, $s.WrapFormat.DistanceRight, $s.WrapFormat.AllowOverlap, $s.Anchor.Start)
    }




    'shape-two-pictures'              = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $a = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        $b = $doc.Shapes.AddPicture($script:Picture, $false, $true, 10, 10, 60, 30, $doc.Range(5, 5))
        Write-Host ('z {0} {1}; names {2} {3}; count {4}' -f $a.ZOrderPosition, $b.ZOrderPosition, $a.Name, $b.Name, $doc.Shapes.Count)
        $b.ZOrder(1)
        Write-Host ('after send to back: z {0} {1}' -f $a.ZOrderPosition, $b.ZOrderPosition)
    }


    'shape-to-inline'                 = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        $inline = $s.ConvertToInlineShape()
        Write-Host ('inline type {0}, w {1}, shapes {2}, inline shapes {3}' -f $inline.Type, $inline.Width, $doc.Shapes.Count, $doc.InlineShapes.Count)
    }

    'shape-picture-properties'        = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        $s.Name = 'Box'
        $s.AlternativeText = 'A box'
        $s.Title = 'Title'
        $s.Rotation = 30
        $s.RelativeHorizontalPosition = 1
        $s.RelativeVerticalPosition = 0
        $s.Left = 40
        $s.Top = 50
        $s.WrapFormat.DistanceLeft = 10
        $s.WrapFormat.AllowOverlap = $false
        $s.LockAnchor = $true
        $s.LockAspectRatio = $true
        $s.Visible = $false
        Write-Host ('after: left {0} top {1} relH {2} relV {3} rot {4} w {5} h {6}' -f $s.Left, $s.Top, $s.RelativeHorizontalPosition, $s.RelativeVerticalPosition, $s.Rotation, $s.Width, $s.Height)
    }

    'shape-picture-wraps'             = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        foreach ($type in 0, 1, 2, 4, 5) {
            $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
            $s.WrapFormat.Type = $type
        }
        $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        $s.WrapFormat.Type = 0
        $s.WrapFormat.Side = 1
        $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        $s.WrapFormat.Type = 0
        $s.WrapFormat.Side = 3
        Write-Host ('count {0}' -f $doc.Shapes.Count)
    }

    'shape-add-textbox'               = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddTextbox(1, 20, 30, 120, 60, $doc.Range(5, 5))
        $s.TextFrame.TextRange.Text = 'Boxed'
        Write-Host ('textbox: type {0}, name "{1}", id {2}, w {3} h {4}, wrap {5}, autoshape {6}, fill {7} line {8}' -f $s.Type, $s.Name, $s.ID, $s.Width, $s.Height, $s.WrapFormat.Type, $s.AutoShapeType, $s.Fill.Visible, $s.Line.Visible)
    }

    'shape-add-rectangle'             = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddShape(1, 20, 30, 80, 40, $doc.Range(5, 5))
        Write-Host ('rect: type {0}, name "{1}", id {2}, w {3} h {4}, wrap {5}, autoshape {6}, fill {7} {8}, line {9} {10} w {11}' -f $s.Type, $s.Name, $s.ID, $s.Width, $s.Height, $s.WrapFormat.Type, $s.AutoShapeType, $s.Fill.Visible, $s.Fill.ForeColor.RGB, $s.Line.Visible, $s.Line.ForeColor.RGB, $s.Line.Weight)
    }

    'shape-add-line'                  = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddLine(20, 30, 100, 70, $doc.Range(5, 5))
        Write-Host ('line: type {0}, name "{1}", w {2} h {3} left {4} top {5}' -f $s.Type, $s.Name, $s.Width, $s.Height, $s.Left, $s.Top)
    }

    'shape-autoshapes'                = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        foreach ($t in 1, 4, 5, 7, 9, 10, 33) {
            $s = $doc.Shapes.AddShape($t, 20, 30, 80, 40, $doc.Range(5, 5))
            Write-Host ('autoshape {0}: name "{1}" type {2} autoshapetype {3} left {4} top {5} relH {6} relV {7}' -f $t, $s.Name, $s.Type, $s.AutoShapeType, $s.Left, $s.Top, $s.RelativeHorizontalPosition, $s.RelativeVerticalPosition)
        }
    }

    'shape-fill-line'                 = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $a = $doc.Shapes.AddShape(1, 20, 30, 80, 40, $doc.Range(5, 5))
        $a.Fill.ForeColor.RGB = 255
        $a.Line.ForeColor.RGB = 65280
        $a.Line.Weight = 3
        $c = $doc.Shapes.AddTextbox(1, 20, 150, 80, 40, $doc.Range(5, 5))
        $c.Fill.ForeColor.RGB = 16711680
        $c.Fill.Transparency = 0.5
        Write-Host ('a fill {0} {1} line {2} {3} {4}; c fill {5} {6} tr {7} line {8} {9}' -f $a.Fill.Visible, $a.Fill.ForeColor.RGB, $a.Line.Visible, $a.Line.ForeColor.RGB, $a.Line.Weight, $c.Fill.Visible, $c.Fill.ForeColor.RGB, $c.Fill.Transparency, $c.Line.Visible, $c.Line.Weight)
    }
    'shape-delete'                    = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $s = $doc.Shapes.AddPicture($script:Picture, $false, $true, 20, 30, 60, 30, $doc.Range(5, 5))
        $s.Delete()
    }

    'shape-from-inline'               = {
        param($doc)
        $doc.Content.InsertAfter('Hello world again')
        $i = $doc.InlineShapes.AddPicture($script:Picture, $false, $true, $doc.Range(5, 5))
        $s = $i.ConvertToShape()
        Write-Host ('shape: type {0}, wrap {1}, left {2} top {3} w {4} h {5} relH {6} relV {7}' -f $s.Type, $s.WrapFormat.Type, $s.Left, $s.Top, $s.Width, $s.Height, $s.RelativeHorizontalPosition, $s.RelativeVerticalPosition)
    }
    # The three selection scenarios do not make a document to compare: they print what Word does, and SelectionWrapperTest checks the
    # library against those values. Run them with New-WordReferences.ps1 and delete the folders they write.
    'selection-behaviour'             = {
        param($doc)
        $sel = $doc.Application.Selection
        function Show($label) { Write-Host ('{0}: start {1} end {2} type {3} text "{4}" | doc "{5}"' -f $label, $sel.Start, $sel.End, $sel.Type, ($sel.Text -replace "`r", '\r'), (($doc.Content.Text) -replace "`r", '\r')) }
        Show 'new'
        $sel.TypeText('Hello world again')
        Show 'typed'
        $sel.TypeParagraph()
        Show 'paragraph'
        $sel.TypeText('Second')
        Show 'second'
        $sel.HomeKey(6) | Out-Null
        Show 'home story'
        $sel.EndKey(5) | Out-Null
        Show 'end line'
        $sel.MoveRight(1, 5) | Out-Null
        Show 'right 5'
        $sel.MoveRight(1, 5, 1) | Out-Null
        Show 'extend right 5'
        $sel.TypeText('X')
        Show 'type over'
        $sel.TypeBackspace()
        Show 'backspace'
        $sel.MoveLeft(1, 3, 1) | Out-Null
        Show 'extend left 3'
        $sel.TypeBackspace()
        Show 'backspace selection'
        $sel.EndKey(6) | Out-Null
        Show 'end story'
        $sel.HomeKey(5, 1) | Out-Null
        Show 'extend home line'
        $sel.MoveUp(4, 1) | Out-Null
        Show 'up paragraph'
        $sel.MoveDown(4, 1) | Out-Null
        Show 'down paragraph'
        $sel.Collapse(1)
        Show 'collapse start'
        $sel.Expand(2) | Out-Null
        Show 'expand word'
        $sel.Expand(4) | Out-Null
        Show 'expand paragraph'
        $sel.Collapse(0)
        Show 'collapse end'
        $sel.MoveLeft(1, 100) | Out-Null
        Show 'left 100'
        $sel.MoveRight(1, 3) | Out-Null
        $sel.Delete() | Out-Null
        Show 'delete char'
        $sel.InsertAfter('Z')
        Show 'insert after'
        $sel.InsertBefore('Y')
        Show 'insert before'
        $doc.Range(0, 3).Select()
        Show 'range select'
        $doc.Range(2, 2).InsertAfter('QQ')
        Show 'after edit before selection'
        $sel.SetRange(1, 4)
        Show 'setrange'
        $sel.TypeText('')
        Show 'type empty'
    }

    'selection-moves'                 = {
        param($doc)
        $sel = $doc.Application.Selection
        function Show($label) { Write-Host ('{0}: start {1} end {2} type {3} text "{4}"' -f $label, $sel.Start, $sel.End, $sel.Type, ($sel.Text -replace "`r", '\r')) }
        $sel.TypeText("Hello world again`rSecond line here`rThird")
        $doc.Range(6, 11).Select()
        Show 'select world'
        [void]$sel.MoveRight(1, 1)
        Show 'right 1 from selection'
        $doc.Range(6, 11).Select()
        [void]$sel.MoveLeft(1, 1)
        Show 'left 1 from selection'
        $doc.Range(6, 11).Select()
        [void]$sel.MoveRight(1, 3)
        Show 'right 3 from selection'
        $doc.Range(6, 11).Select()
        [void]$sel.Move(1, 2)
        Show 'move 2 from selection'
        $doc.Range(6, 11).Select()
        [void]$sel.MoveEnd(1, 2)
        Show 'moveend 2'
        $doc.Range(6, 11).Select()
        [void]$sel.MoveStart(1, -2)
        Show 'movestart -2'
        $doc.Range(8, 8).Select()
        [void]$sel.MoveRight(1, 3, 1)
        Show 'extend right 3'
        [void]$sel.MoveLeft(1, 5, 1)
        Show 'extend left 5 (crossing)'
        [void]$sel.MoveRight(1, 1, 1)
        Show 'extend right 1'
        $doc.Range(8, 8).Select()
        [void]$sel.MoveRight(2, 1)
        Show 'right word'
        [void]$sel.MoveRight(2, 1, 1)
        Show 'extend right word'
        $doc.Range(25, 25).Select()
        [void]$sel.MoveUp(4, 1)
        Show 'up paragraph from 25'
        [void]$sel.MoveUp(4, 1)
        Show 'up again'
        [void]$sel.MoveDown(4, 2)
        Show 'down 2'
        $doc.Range(2, 2).Select()
        [void]$sel.MoveDown(4, 1, 1)
        Show 'extend down paragraph'
        $doc.Range(6, 11).Select()
        $sel.TypeParagraph()
        Show 'typeparagraph over selection'
        $doc.Range(6, 11).Select()
        Write-Host ('delete returns {0}' -f $sel.Delete())
        Show 'delete selection'
        Write-Host ('delete word returns {0}' -f $sel.Delete(2, 1))
        Show 'delete word'
        $sel.HomeKey(6) | Out-Null
        Write-Host ('homekey returns {0}' -f $sel.EndKey(6))
        Show 'end story'
        Write-Host ('mover left at start returns {0}' -f $sel.MoveLeft(1, 3))
        $sel.EndKey(6) | Out-Null
        Write-Host ('moveright at end returns {0}' -f $sel.MoveRight(1, 3))
    }

    'selection-tables'                = {
        param($doc)
        $sel = $doc.Application.Selection
        function Show($label) { Write-Host ('{0}: start {1} end {2} type {3} text "{4}"' -f $label, $sel.Start, $sel.End, $sel.Type, ($sel.Text -replace "`r", '\r' -replace "`a", '\a')) }
        $doc.Content.InsertAfter("Before`r")
        [void]$doc.Tables.Add($doc.Range($doc.Content.End - 1, $doc.Content.End - 1), 2, 2)
        $doc.Tables.Item(1).Cell(1, 1).Range.Text = 'a'
        $doc.Tables.Item(1).Cell(1, 2).Range.Text = 'b'
        $doc.Tables.Item(1).Cell(2, 1).Range.Text = 'c'
        $doc.Tables.Item(1).Cell(2, 2).Range.Text = 'd'
        $doc.Tables.Item(1).Cell(1, 2).Range.Select()
        Show 'cell'
        $sel.SelectRow()
        Show 'row'
        $sel.SelectColumn()
        Show 'column'
        $sel.SelectCell()
        Show 'selectcell'
        $sel.HomeKey(10) | Out-Null
        Show 'home row'
        $sel.EndKey(10) | Out-Null
        Show 'end row'
        $sel.HomeKey(9) | Out-Null
        Show 'home column'
        $sel.EndKey(9) | Out-Null
        Show 'end column'
    }

    'document-settings'               = {
        param($doc)
        $doc.DefaultTabStop = 54
        $doc.AutoHyphenation = $true
        $doc.HyphenateCaps = $false
        $doc.ConsecutiveHyphensLimit = 2
        $doc.UpdateStylesOnOpen = $true
        $doc.EmbedTrueTypeFonts = $true
        $doc.SaveSubsetFonts = $true
        $doc.DoNotEmbedSystemFonts = $false
        $doc.RemovePersonalInformation = $true
        $doc.RemoveDateAndTime = $true
    }

    'default-table-style'             = {
        param($doc)
        $doc.SetDefaultTableStyle('Table Grid', $false)
    }

    'heading-insert-paragraph-after'  = {
        param($doc)
        $doc.Content.InsertAfter('Heading')
        $doc.Paragraphs.Item(1).Style = -2  # wdStyleHeading1
        $doc.Paragraphs.Item(1).Range.InsertParagraphAfter()
    }

    # Each built-in style the library can add (BuiltInStyle, plus Table Grid), applied in turn: Word writes the definition of each, and of
    # the linked character style of a paragraph style. These are also the definitions the library adds (OpenXml.Wrappers/Word/Templates).
    # The list styles that carry numbering (List Bullet and List Number, 1 to 5) are left out until lists are done: Word writes a numbering part for them.
    # Keep the list in step with BuiltInStylesScenario in MarkupConformanceTest.
    'built-in-styles'                 = {
        param($doc)
        $styles = @(-2, -3, -4, -5, -6, -7, -8, -9, -10, -20, -21, -22, -23, -24, -25, -26, -27, -28, -30, -31, -32, -33, -35, -39, -40, -41,
            -42, -43, -44, -63, -64, -65, -67, -68, -75, -76, -77, -81, -82, -86, -87, -88, -89, -91, -180, -181, -182, -261, -262, -263, -264, -265,
            -11, -12, -13, -14, -15, -16, -17, -18, -19, -29, -34, -36, -37, -38, -45, -46, -47, -48, -51, -52, -53, -54,
            -69, -70, -71, -72, -73, -74, -78, -79, -80, -83, -84, -85, -90, -95, -96, -97, -98, -99, -100, -101, -102,
            -103, -104, -105, -158,
            'Comment Subject', 'Placeholder Text')  # no WdBuiltinStyle constant; Word adds them with a comment and a content control
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
        foreach ($id in -159..-178) {  # the table styles
            $table.Style = $doc.Styles.Item([object]$id)
        }
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

    'style-flags'                     = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.QuickStyle = $true
        $s.Priority = 5
        $s.Locked = $true
        $s.UnhideWhenUsed = $true
        $s.AutomaticallyUpdate = $true
        $s.NoSpaceBetweenParagraphsOfSameStyle = $true
        $s.Visibility = $false
    }
    'style-based-next'                = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.BaseStyle = 'Heading 1'
        $s.NextParagraphStyle = 'Heading 2'
    }
    'style-language'                  = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.LanguageID = 1033
        $s.LanguageIDFarEast = 1041
        $s.NoProofing = $true
    }
    'style-add-paragraph'             = {
        param($doc)
        [void]$doc.Styles.Add('RefStyle', 1)
    }
    'style-add-character'             = {
        param($doc)
        [void]$doc.Styles.Add('RefStyle', 2)
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

    'style-borders-paragraph'         = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.Borders.OutsideLineStyle = 1
        $s.Borders.OutsideColor = 255
    }
    'style-borders-top'               = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.Borders.Item(-1).LineStyle = 1
    }
    'style-shading-paragraph'         = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.Shading.BackgroundPatternColor = 255
        $s.Shading.Texture = 100
    }
    'style-borders-shading-character' = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 2)
        $s.Font.Borders.Item(-1).LineStyle = 1
        $s.Font.Shading.BackgroundPatternColor = 65280
    }
    'style-font-borders-shading'      = {
        param($doc)
        $s = $doc.Styles.Add('RefStyle', 1)
        $s.Font.Borders.Item(-1).LineStyle = 1
        $s.Font.Shading.BackgroundPatternColor = 65280
    }
    'table-style-add'                 = {
        param($doc)
        [void]$doc.Styles.Add('RefTable', 3)
    }
    'table-style-members'             = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Alignment = 1
        $t.AllowBreakAcrossPage = 0
        $t.TopPadding = 5
        $t.BottomPadding = 6
        $t.LeftPadding = 7
        $t.RightPadding = 8
        $t.RowStripe = 2
        $t.ColumnStripe = 3
        $t.Spacing = 2
        $t.LeftIndent = 10
    }
    'table-style-borders-shading'     = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Borders.OutsideLineStyle = 1
        $t.Borders.InsideLineStyle = 1
        $t.Shading.BackgroundPatternColor = 255
    }
    'table-style-conditional-mixed'   = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $c = $t.Condition(0)
        $c.Font.Bold = $true
        $c.ParagraphFormat.Alignment = 1
        $c.Shading.BackgroundPatternColor = 65280
    }
    'table-style-padding-top'         = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.TopPadding = 5
    }
    'table-style-conditional-banding' = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Condition(2).Shading.BackgroundPatternColor = 255
        $t.Condition(5).Font.Italic = $true
        $t.Condition(11).Font.Bold = $true
    }
    'table-style-cond-shading-first'  = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Condition(0).Shading.BackgroundPatternColor = 65280
    }
    'table-style-cond-shading-band2'  = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Condition(3).Shading.BackgroundPatternColor = 65280
    }
    'table-style-cond-shading-vband'  = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Condition(6).Shading.BackgroundPatternColor = 65280
    }
    'table-style-cond-borders-band'   = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Condition(2).Borders.Item(-3).LineStyle = 1
    }
    'table-style-cond-padding-only'   = {
        param($doc)
        $t = $doc.Styles.Add('RefTable', 3).Table
        $t.Condition(0).TopPadding = 3
    }
    'find-match-prefix'               = {
        param($doc)
        $doc.Content.InsertAfter('pre prefix preach unpre repre')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.MatchPrefix = $true
        [void]$f.Execute('pre', $m, $m, $m, $m, $m, $m, $m, $m, 'X', 2)
    }
    'find-match-suffix'               = {
        param($doc)
        $doc.Content.InsertAfter('ing sing singing ingot')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.MatchSuffix = $true
        [void]$f.Execute('ing', $m, $m, $m, $m, $m, $m, $m, $m, 'X', 2)
    }
    'find-match-prefix-suffix'        = {
        param($doc)
        $doc.Content.InsertAfter('cat concat cats cat')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.MatchPrefix = $true
        $f.MatchSuffix = $true
        [void]$f.Execute('cat', $m, $m, $m, $m, $m, $m, $m, $m, 'X', 2)
    }
    'find-ignore-punct'               = {
        param($doc)
        $doc.Content.InsertAfter('co-op coop co,op co op')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.IgnorePunct = $true
        [void]$f.Execute('coop', $m, $m, $m, $m, $m, $m, $m, $m, 'X', 2)
    }
    'find-ignore-space'               = {
        param($doc)
        $doc.Content.InsertAfter('a b ab a  b ab')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.IgnoreSpace = $true
        [void]$f.Execute('ab', $m, $m, $m, $m, $m, $m, $m, $m, 'X', 2)
    }
    'find-match-phrase'               = {
        param($doc)
        $doc.Content.InsertAfter("a  b`tc`vd e and two  three")
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.MatchPhrase = $true
        [void]$f.Execute('a b c d e', $m, $m, $m, $m, $m, $m, $m, $m, 'X', 2)
        [void]$f.Execute('two three', $m, $m, $m, $m, $m, $m, $m, $m, 'Y', 2)
    }
    'find-replace-language'           = {
        param($doc)
        $doc.Content.InsertAfter('One two three')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.LanguageID = 1036
        $f.Replacement.NoProofing = $true
        [void]$f.Execute('two', $m, $m, $m, $m, $m, $m, $m, $true, '^&', 2)
    }
    'find-replace-font'               = {
        param($doc)
        $doc.Content.InsertAfter('One two three two')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.Font.Bold = $true
        $f.Replacement.Font.Size = 14
        [void]$f.Execute('two', $m, $m, $m, $m, $m, $m, $m, $true, '2', 2)
    }
    'find-replace-font-text-unchanged' = {
        param($doc)
        $doc.Content.InsertAfter('One two three two')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.Font.Italic = $true
        [void]$f.Execute('two', $m, $m, $m, $m, $m, $m, $m, $true, '^&', 2)
    }
    'find-format-only'                = {
        param($doc)
        $doc.Content.InsertAfter('plain bold plain bold')
        $doc.Range(6, 10).Font.Bold = $true
        $doc.Range(17, 21).Font.Bold = $true
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Font.Bold = $true
        $f.Replacement.Font.Italic = $true
        $f.Replacement.Font.Bold = $false
        [void]$f.Execute('', $m, $m, $m, $m, $m, $m, $m, $true, '', 2)
    }
    'find-replace-style'              = {
        param($doc)
        $doc.Content.InsertAfter('One')
        $doc.Content.InsertParagraphAfter()
        $doc.Content.InsertAfter('Two')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.Style = 'Heading 1'
        [void]$f.Execute('One', $m, $m, $m, $m, $m, $m, $m, $true, 'Uno', 2)
    }
    'find-replace-paragraphformat'    = {
        param($doc)
        $doc.Content.InsertAfter('One')
        $doc.Content.InsertParagraphAfter()
        $doc.Content.InsertAfter('Two')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.ParagraphFormat.Alignment = 1
        [void]$f.Execute('Two', $m, $m, $m, $m, $m, $m, $m, $true, 'Two', 2)
    }
    'find-replace-highlight'          = {
        param($doc)
        $doc.Content.InsertAfter('One two three')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.Highlight = $true
        [void]$f.Execute('two', $m, $m, $m, $m, $m, $m, $m, $true, '^&', 2)
    }

    'find-replace-empty-with-format'  = {
        param($doc)
        $doc.Content.InsertAfter('One two three')
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Replacement.Font.Italic = $true
        [void]$f.Execute('two', $m, $m, $m, $m, $m, $m, $m, $true, '', 2)
    }
    'find-style-criteria'             = {
        param($doc)
        $doc.Content.InsertAfter('One')
        $doc.Content.InsertParagraphAfter()
        $doc.Content.InsertAfter('Two')
        $doc.Paragraphs(1).Style = 'Heading 1'
        $m = [Type]::Missing
        $f = $doc.Content.Find
        $f.ClearFormatting()
        $f.Replacement.ClearFormatting()
        $f.Style = 'Heading 1'
        $f.Replacement.Font.Italic = $true
        [void]$f.Execute('', $m, $m, $m, $m, $m, $m, $m, $true, '', 2)
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
