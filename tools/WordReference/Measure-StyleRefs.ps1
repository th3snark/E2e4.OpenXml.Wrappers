#Requires -Version 7
<#
.SYNOPSIS
Measures what Word's STYLEREF field shows in the main text and writes the results for FieldStyleReferenceTest.
.DESCRIPTION
Builds a document of headings numbered by a list template linked to Heading 1 to 3 ("Chapter %1", "(%2)", "%3."), body paragraphs, an
empty Heading 2 and a word in the Strong character style. From several paragraphs it adds STYLEREF fields with each code, updates them and
records the result in OpenXml.Wrappers.Test/TestFiles/StyleRefs.txt (tab-separated: host, code, result). The test builds the same document
through the library and compares with it.

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

# Keep in step with FieldStyleReferenceTest.
$levels = @(@('Chapter %1', 0), @('(%2)', 4), @('%3.', 2))
$paragraphs = @(@('Normal', 'Start'), @('Heading 1', 'Intro'), @('Heading 2', 'Alpha'), @('Normal', 'BodyAlpha'), @('Heading 1', 'Second'),
    @('Normal', 'BodySecond'), @('Heading 2', 'Beta'), @('Heading 2', ''), @('Normal', 'BodyEmpty'), @('Normal', 'Strong words here'))
$hosts = @(1, 2, 3, 4, 6, 9, 10)
$codes = @(
    'STYLEREF "Heading 1"', 'STYLEREF "Heading 2"', 'STYLEREF "heading 1"', 'STYLEREF Heading1', 'STYLEREF 1', 'STYLEREF "Normal"',
    'STYLEREF "Heading 1" \n', 'STYLEREF "Heading 2" \n', 'STYLEREF "Heading 2" \r', 'STYLEREF "Heading 2" \w', 'STYLEREF "Heading 1" \r',
    'STYLEREF "Heading 2" \w \t', 'STYLEREF "Heading 2" \r \t', 'STYLEREF "Heading 1" \n \t',
    'STYLEREF "Heading 2" \p', 'STYLEREF "Heading 1" \p', 'STYLEREF "Heading 2" \n \p', 'STYLEREF "Heading 1" \w \p',
    'STYLEREF "Normal" \n', 'STYLEREF "Normal" \w', 'STYLEREF "Normal" \p', 'STYLEREF "Normal" \n \p',
    'STYLEREF "Heading 1" \l', 'STYLEREF "Heading 1" \* Upper', 'STYLEREF "Heading 9"', 'STYLEREF "Title"', 'STYLEREF "No such style"',
    'STYLEREF Strong', 'STYLEREF "Heading 1" \s')

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# What Word 16 shows for STYLEREF fields in the main text. Written by tools/WordReference/Measure-StyleRefs.ps1; do not edit.')
$lines.Add("# host`tcode`tresult")

$template = Backup-NormalTemplate
$word = Start-Word $WordPath
try {
    $d = $word.Application.Documents.Add()
    try {
        $listTemplate = $d.ListTemplates.Add($true)
        for ($i = 1; $i -le 3; $i++) {
            $level = $listTemplate.ListLevels.Item($i)
            $level.NumberFormat = $levels[$i - 1][0]
            $level.NumberStyle = $levels[$i - 1][1]
            $level.LinkedStyle = "Heading $i"
        }
        $d.Content.Text = $paragraphs[0][1]
        for ($k = 1; $k -lt $paragraphs.Count; $k++) {
            $d.Content.InsertParagraphAfter()
            $d.Paragraphs.Item($d.Paragraphs.Count).Range.InsertBefore($paragraphs[$k][1])
        }
        $d.Paragraphs.Item(2).Range.ListFormat.ApplyListTemplateWithLevel($listTemplate, $false, 0, 2) | Out-Null
        for ($k = 0; $k -lt $paragraphs.Count; $k++) {
            $d.Paragraphs.Item($k + 1).Style = $paragraphs[$k][0]
        }
        $strong = $d.Paragraphs.Item($paragraphs.Count).Range
        $strong.SetRange($strong.Start, $strong.Start + 6)
        $strong.Style = 'Strong'

        foreach ($h in $hosts) {
            foreach ($code in $codes) {
                $at = $d.Paragraphs.Item($h).Range
                $at.MoveEnd(1, -1) | Out-Null
                $at.Collapse(0)
                $field = $d.Fields.Add($at, -1, $code, $false)
                $field.Update() | Out-Null
                $lines.Add(("{0}`t{1}`t{2}" -f $paragraphs[$h - 1][1], $code, $field.Result.Text))
                $field.Delete()
            }
        }
    }
    finally {
        $d.Close(0)
    }
}
finally {
    Stop-Word $word
}
Assert-NormalTemplateUnchanged $template

$out = Join-Path $PSScriptRoot '..\..\OpenXml.Wrappers.Test\TestFiles\StyleRefs.txt'
[System.IO.File]::WriteAllLines([System.IO.Path]::GetFullPath($out), $lines, [System.Text.UTF8Encoding]::new($false))
Write-Host ('{0} results written to {1}' -f ($lines.Count - 2), $out)
