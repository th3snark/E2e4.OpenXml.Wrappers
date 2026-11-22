#Requires -Version 7
<#
.SYNOPSIS
Measures what Word's REF field shows with the list-number switches (\n, \r, \w, \t, \p) and writes the results for FieldReferenceNumberTest.
.DESCRIPTION
For each of five list templates linked to Heading 1 to 3, builds a document of numbered headings and body paragraphs, bookmarks the text of
every paragraph, and from a number of paragraphs adds REF fields to the bookmarks with each switch, updates them and records the result.
The results go to OpenXml.Wrappers.Test/TestFiles/RefNumbers.txt (tab-separated: template, layout, host, target, switches, result), which the
test builds the same documents for through the library and compares with.

Word is needed, so this runs on a developer's machine, not in CI.
.PARAMETER WordPath
WINWORD.EXE. By default it is found through the App Paths registry key.
#>
[CmdletBinding()]
param(
    [string] $WordPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WordAutomation.ps1')
if (-not $WordPath) {
    $WordPath = Get-WordPath
}

# The levels of each template: number format and style (0 arabic, 2 lowercase roman, 4 lowercase letter). Keep in step with the test.
$templates = [ordered]@{
    1 = @(@('Chapter %1', 0), @('(%2)', 4), @('%3.', 2))
    2 = @(@('%1.', 0), @('%1.%2.', 0), @('%1.%2.%3.', 0))
    3 = @(@('%1)', 0), @('%2)', 4), @('%3)', 2))
    4 = @(@('%1.', 0), @('%2.', 4), @('%3.', 2))
    5 = @(@('Part %1:', 0), @('Section %1-%2', 0), @('Clause %3 of', 0))
}

# The paragraphs of each layout (style and text), the paragraphs the fields go at the end of, the paragraphs they refer to and the switches.
$layouts = [ordered]@{
    A = @{
        Paragraphs = @(@('Normal', 'Start'), @('Heading 1', 'One'), @('Heading 2', 'OneA'), @('Heading 3', 'OneAi'), @('Heading 3', 'OneAii'),
            @('Heading 2', 'OneB'), @('Heading 1', 'Two'), @('Heading 2', 'TwoA'), @('Heading 3', 'TwoAi'))
        Hosts = @(1, 5, 9)
        Targets = @(2, 3, 4, 5, 6, 7, 8, 9)
        Switches = @('\n', '\r', '\w', '\n \t', '\r \t', '\w \t', '\p', '\r \p', '\n \p', '\w \p')
    }
    B = @{
        Paragraphs = @(@('Normal', 'Start'), @('Heading 1', 'One'), @('Normal', 'BodyOne'), @('Heading 2', 'OneA'), @('Heading 3', 'OneAi'),
            @('Normal', 'BodyOneAi'), @('Normal', 'Plain'), @('Heading 2', 'OneB'), @('Heading 1', 'Two'), @('Normal', 'BodyTwo'))
        Hosts = @(3, 6, 7, 10)
        Targets = @(2, 4, 5, 7, 8, 9)
        Switches = @('\n', '\r', '\w', '\n \t', '\r \t', '\w \t')
    }
}

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# What Word 16 shows for REF fields with list-number switches. Written by tools/WordReference/Measure-RefNumbers.ps1; do not edit.')
$lines.Add("# template`tlayout`thost`ttarget`tswitches`tresult")

$template = Backup-NormalTemplate
$word = Start-Word $WordPath
try {
    foreach ($t in $templates.Keys) {
        foreach ($name in $layouts.Keys) {
            $layout = $layouts[$name]
            $d = $word.Documents.Add()
            try {
                $listTemplate = $d.ListTemplates.Add($true)
                for ($i = 1; $i -le 3; $i++) {
                    $level = $listTemplate.ListLevels.Item($i)
                    $level.NumberFormat = $templates[$t][$i - 1][0]
                    $level.NumberStyle = $templates[$t][$i - 1][1]
                    $level.LinkedStyle = "Heading $i"
                }
                $paragraphs = $layout.Paragraphs
                $d.Content.Text = $paragraphs[0][1]
                for ($k = 1; $k -lt $paragraphs.Count; $k++) {
                    $d.Content.InsertParagraphAfter()
                    $d.Paragraphs.Item($d.Paragraphs.Count).Range.InsertBefore($paragraphs[$k][1])
                }
                $d.Paragraphs.Item(2).Range.ListFormat.ApplyListTemplateWithLevel($listTemplate, $false, 0, 2) | Out-Null
                for ($k = 0; $k -lt $paragraphs.Count; $k++) {
                    $d.Paragraphs.Item($k + 1).Style = $paragraphs[$k][0]
                }
                for ($k = 1; $k -le $paragraphs.Count; $k++) {
                    $range = $d.Paragraphs.Item($k).Range
                    $range.MoveEnd(1, -1) | Out-Null
                    $d.Bookmarks.Add("B$k", $range) | Out-Null
                }
                foreach ($h in $layout.Hosts) {
                    foreach ($target in $layout.Targets) {
                        foreach ($switches in $layout.Switches) {
                            $at = $d.Paragraphs.Item($h).Range
                            $at.MoveEnd(1, -1) | Out-Null
                            $at.Collapse(0)
                            $field = $d.Fields.Add($at, -1, "REF B$target $switches", $false)
                            $field.Update() | Out-Null
                            $lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f $t, $name, $paragraphs[$h - 1][1], $paragraphs[$target - 1][1], $switches, $field.Result.Text))
                            $field.Delete()
                        }
                    }
                }
            }
            finally {
                $d.Close(0)
            }
        }
    }
}
finally {
    Stop-Word $word
}
Assert-NormalTemplateUnchanged $template

$out = Join-Path $PSScriptRoot '..\..\OpenXml.Wrappers.Test\TestFiles\RefNumbers.txt'
[System.IO.File]::WriteAllLines([System.IO.Path]::GetFullPath($out), $lines, [System.Text.UTF8Encoding]::new($false))
Write-Host ('{0} results written to {1}' -f ($lines.Count - 2), $out)
