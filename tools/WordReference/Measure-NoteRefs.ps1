#Requires -Version 7
<#
.SYNOPSIS
Measures what Word's SECTION and NOTEREF fields show and writes the results for FieldNoteReferenceTest.
.DESCRIPTION
Builds a document of three sections with footnotes, endnotes (one footnote with a custom mark) and bookmarks on them, sets the options of
the notes (number style, starting number, numbering rule) in a few ways, and from a number of paragraphs (and from the text of a footnote)
adds SECTION and NOTEREF fields with their switches, updates them and records the result and the character style of its first character.
The results go to OpenXml.Wrappers.Test/TestFiles/NoteRefs.txt (tab-separated: options, host, code, result, style), which the test builds
the same document for through the library and compares with.

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

# Keep in step with FieldNoteReferenceTest. The paragraphs, and the note each has at the end of its text: F a footnote, E an endnote, * a
# footnote with the custom mark "*". A section break (next page) follows HostB, and a continuous one follows HostD.
$paragraphs = @(@('HostA', ''), @('Alpha', 'F'), @('Beta', 'F'), @('Gamma', 'E'), @('HostB', ''), @('Delta', 'F'), @('HostC', ''),
    @('Epsilon', 'E'), @('Plain', ''), @('HostD', ''), @('Zeta', '*'), @('Eta', 'F'), @('HostE', ''))
# Bookmarks: on the mark of the note of a paragraph (mark), on the text of a paragraph with its mark (text), or over paragraphs (from, to).
$bookmarks = [ordered]@{
    FnA = @('mark', 'Alpha'); FnB = @('text', 'Beta'); EnA = @('mark', 'Gamma'); FnC = @('mark', 'Delta'); EnB = @('mark', 'Epsilon')
    NoNote = @('text', 'Plain'); FnStar = @('mark', 'Zeta'); FnD = @('mark', 'Eta'); Two = @('span', 'Zeta', 'Eta')
}
# The options of the notes: footnotes then endnotes, as number style (WdNoteNumberStyle), starting number, numbering rule (WdNumberingRule)
# and, for endnotes, location (WdEndnoteLocation).
$options = [ordered]@{
    default = @(@(0, 1, 0), @(2, 1, 0, 1))
    roman = @(@(1, 3, 0), @(0, 1, 0, 1))
    symbol = @(@(9, 1, 0), @(4, 1, 1, 0))
    restart = @(@(0, 1, 1), @(2, 5, 0, 1))
    page = @(@(0, 1, 2), @(2, 1, 0, 1))
}
$allHosts = @('HostA', 'HostB', 'HostC', 'HostD', 'HostE', 'Note')
$someHosts = @('HostA', 'HostE', 'Note')

$codes = [System.Collections.Generic.List[string]]::new()
foreach ($code in 'SECTION', 'SECTION \* roman', 'SECTION \* CardText') {
    $codes.Add($code)
}
foreach ($name in @($bookmarks.Keys) + 'Missing') {
    foreach ($switches in '', ' \f', ' \p') {
        $codes.Add('NOTEREF ' + $name + $switches)
    }
}
foreach ($name in 'FnA', 'EnA', 'FnStar') {
    foreach ($switches in ' \h', ' \f \h', ' \p \f', ' \* roman', ' \* Upper', ' \* CardText', ' \* ALPHABETIC') {
        $codes.Add('NOTEREF ' + $name + $switches)
    }
}

function Find-Paragraph($Document, [string] $Text) {
    foreach ($p in $Document.Paragraphs) {
        $r = $p.Range
        $r.MoveEnd(1, -1) | Out-Null
        if ($r.Text -and $r.Text.StartsWith($Text)) {
            return $p
        }
    }
    throw "No paragraph $Text"
}

function Get-TextRange($Document, [string] $Text) {
    $r = (Find-Paragraph $Document $Text).Range
    $r.MoveEnd(1, -1) | Out-Null
    return $r
}

$lines = [System.Collections.Generic.List[string]]::new()
$lines.Add('# What Word 16 shows for SECTION and NOTEREF fields. Written by tools/WordReference/Measure-NoteRefs.ps1; do not edit.')
$lines.Add("# options`thost`tcode`tresult`tstyle")

$template = Backup-NormalTemplate
$word = Start-Word $WordPath
try {
    $word.Application.ScreenUpdating = $false
    foreach ($optionName in $options.Keys) {
        $d = $word.Application.Documents.Add()
        try {
            $d.Content.Text = $paragraphs[0][0]
            for ($k = 1; $k -lt $paragraphs.Count; $k++) {
                $d.Content.InsertParagraphAfter()
                $d.Paragraphs.Item($d.Paragraphs.Count).Range.InsertBefore($paragraphs[$k][0])
            }
            $notes = @{}
            foreach ($p in $paragraphs) {
                if (-not $p[1]) {
                    continue
                }
                $at = Get-TextRange $d $p[0]
                $at.Collapse(0)
                if ($p[1] -eq 'E') {
                    $note = $d.Endnotes.Add($at)
                }
                elseif ($p[1] -eq '*') {
                    $note = $d.Footnotes.Add($at, '*')
                }
                else {
                    $note = $d.Footnotes.Add($at)
                }
                $note.Range.Text = 'Note ' + $p[0]
                $notes[$p[0]] = $note
            }
            foreach ($h in 'HostB', 'HostD') {
                $at = Get-TextRange $d $h
                $at.Collapse(0)
                $at.InsertBreak($(if ($h -eq 'HostB') { 2 } else { 3 }))
            }
            foreach ($name in $bookmarks.Keys) {
                $b = $bookmarks[$name]
                if ($b[0] -eq 'mark') {
                    $d.Bookmarks.Add($name, $notes[$b[1]].Reference) | Out-Null
                }
                elseif ($b[0] -eq 'text') {
                    $d.Bookmarks.Add($name, (Get-TextRange $d $b[1])) | Out-Null
                }
                else {
                    $r = Get-TextRange $d $b[1]
                    $r.End = (Get-TextRange $d $b[2]).End
                    $d.Bookmarks.Add($name, $r) | Out-Null
                }
            }
            $o = $options[$optionName]
            $d.Footnotes.NumberingRule = $o[0][2]
            $d.Footnotes.NumberStyle = $o[0][0]
            $d.Footnotes.StartingNumber = $o[0][1]
            $d.Endnotes.Location = $o[1][3]
            $d.Endnotes.NumberingRule = $o[1][2]
            $d.Endnotes.NumberStyle = $o[1][0]
            $d.Endnotes.StartingNumber = $o[1][1]

            $hosts = if ($optionName -eq 'default') { $allHosts } else { $someHosts }
            foreach ($h in $hosts) {
                foreach ($code in $codes) {
                    if ($h -eq 'Note') {
                        $at = $notes['Delta'].Range
                    }
                    else {
                        $at = Get-TextRange $d $h
                    }
                    $at.Collapse(0)
                    $field = $d.Fields.Add($at, -1, $code, $false)
                    $field.Update() | Out-Null
                    $result = $field.Result
                    $style = ''
                    if ($result.Text) {
                        $first = $result.Characters.Item(1)
                        try {
                            $style = $first.CharacterStyle.NameLocal
                        }
                        catch {
                            $style = ''
                        }
                    }
                    $lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}" -f $optionName, $h, $code, $result.Text, $style))
                    $field.Delete()
                    $d.UndoClear()
                }
            }
            Write-Host ('{0}: {1} results' -f $optionName, ($lines.Count - 2))
        }
        finally {
            $d.Close(0)
        }
    }
}
finally {
    Stop-Word $word
}
Assert-NormalTemplateUnchanged $template

$out = Join-Path $PSScriptRoot '..\..\OpenXml.Wrappers.Test\TestFiles\NoteRefs.txt'
[System.IO.File]::WriteAllLines([System.IO.Path]::GetFullPath($out), $lines, [System.Text.UTF8Encoding]::new($false))
Write-Host ('{0} results written to {1}' -f ($lines.Count - 2), $out)
