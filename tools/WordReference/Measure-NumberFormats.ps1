#Requires -Version 7
<#
.SYNOPSIS
Measures what Word shows for the number formats of the \* switch and writes the results for FieldNumberFormatTest.
.DESCRIPTION
Adds a formula field (= value \* format) for every value and format below, updates it and records the result, then does the same for a few
SEQ and QUOTE fields. The results go to OpenXml.Wrappers.Test/TestFiles/NumberFormats.txt (tab-separated: code, result). The test works out
the same codes through the library and compares with them.

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

# Keep in step with FieldNumberFormatTest.
$values = @(0..31) + @(40, 42, 50, 57, 60, 68, 70, 79, 80, 90, 99, 100, 101, 105, 110, 111, 115, 120, 199, 200, 212, 300, 999, 1000, 1001,
    1010, 1100, 1234, 2000, 3999, 4000, 4999, 5000, 9999, 10000, 12345, 32767, 32768, 65535, 99999, 100000, 999999, 1000000, 1234567) +
    @('-1', '-5', '-21', '0.5', '1.5', '2.4', '2.5', '3.5', '21.75', '-1.5')
$formats = @('ARABIC', 'Arabic', 'roman', 'ROMAN', 'Roman', 'alphabetic', 'ALPHABETIC', 'Alphabetic', 'Ordinal', 'ORDINAL', 'CardText',
    'cardtext', 'CARDTEXT', 'OrdText', 'ordtext', 'ORDTEXT', 'Hex', 'hex', 'HEX', 'ArabicDash', 'DollarText', 'dollartext', 'DOLLARTEXT',
    'CardText \* Upper', 'CardText \* FirstCap', 'CardText \* Caps', 'OrdText \* Caps', 'DollarText \* Upper', 'Upper \* CardText',
    'Ordinal \* Upper', 'hex \* Upper', 'HEX \* Lower', 'roman \* Upper', 'CardText \* MERGEFORMAT', 'ROMAN \* CardText', 'CardText \* ROMAN')
$codes = [System.Collections.Generic.List[string]]::new()
foreach ($value in $values) {
    foreach ($format in $formats) {
        $codes.Add(('= {0} \* {1}' -f $value, $format))
    }
}
foreach ($code in @('SEQ A \r 21 \* OrdText', 'SEQ B \r 3 \* CardText', 'SEQ C \r 255 \* Hex', 'SEQ D \r 7 \* ArabicDash',
        'SEQ E \r 12 \* Ordinal', 'SEQ F \r 42 \* DollarText', 'SEQ G \r 0 \* CardText', 'SEQ H \r 0 \* OrdText', 'SEQ I \r 0 \* Ordinal',
        'SEQ J \r 0 \* Hex', 'SEQ K \r 0 \* ArabicDash', 'SEQ L \r 0 \* roman', 'SEQ M \r 0 \* alphabetic', 'SEQ N \r 30 \* OrdText \* Upper',
        'QUOTE "21" \* CardText', 'QUOTE "21" \* OrdText', 'QUOTE "abc" \* CardText', 'QUOTE "21 apples" \* CardText', 'QUOTE "2.5" \* CardText',
        'QUOTE "21" \* Hex', 'QUOTE "21" \* ROMAN', 'QUOTE "21" \* Ordinal',
        'SEQ Q \r 5 \h \* roman', 'SEQ R \r 5 \h \* MERGEFORMAT', 'QUOTE "" \* CardText', 'QUOTE " 7" \* CardText', 'QUOTE "x7" \* CardText',
        'QUOTE "7x" \* Ordinal', 'QUOTE "-7" \* CardText', 'QUOTE "1,234" \* CardText', 'QUOTE "Hello" \* Upper \* Ordinal',
        '= 21 \# "0.00" \* CardText')) {
    $codes.Add($code)
}
foreach ($value in 500, 779, 780, 781, 800, 998) {
    $codes.Add(('= {0} \* alphabetic' -f $value))
}
foreach ($value in '0.4', '0.6', '0.9', '0.99', '-0.4', '-0.6', '-0.5', '1.0', '0.0') {
    $codes.Add(('= {0} \* ARABIC' -f $value))
    $codes.Add(('= {0} \* CardText' -f $value))
}
foreach ($value in '1.005', '2.675', '999999.99', '999999.995', '0.004', '0.996') {
    $codes.Add(('= {0} \* DollarText' -f $value))
}

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# What Word 16 shows for the number formats of \*. Written by tools/WordReference/Measure-NumberFormats.ps1; do not edit.')
$lines.Add("# code`tresult")

$template = Backup-NormalTemplate
$word = Start-Word $WordPath
try {
    $word.Application.ScreenUpdating = $false
    $d = $word.Application.Documents.Add()
    try {
        $done = 0
        foreach ($code in $codes) {
            $at = $d.Content
            $at.Collapse(1)
            $field = $d.Fields.Add($at, -1, $code, $false)
            $field.Update() | Out-Null
            $lines.Add(("{0}`t{1}" -f $code, $field.Result.Text))
            $d.Content.Delete() | Out-Null
            # Every edit goes on the undo stack, which makes each one slower than the last; clearing it keeps the run short.
            $d.UndoClear()
            $done++
            if ($done % 250 -eq 0) {
                Write-Host ('{0} of {1}' -f $done, $codes.Count)
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

$out = Join-Path $PSScriptRoot '..\..\OpenXml.Wrappers.Test\TestFiles\NumberFormats.txt'
[System.IO.File]::WriteAllLines([System.IO.Path]::GetFullPath($out), $lines, [System.Text.UTF8Encoding]::new($false))
Write-Host ('{0} results written to {1}' -f ($lines.Count - 2), $out)
