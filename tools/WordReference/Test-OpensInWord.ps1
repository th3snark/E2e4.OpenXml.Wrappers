#Requires -Version 7
<#
.SYNOPSIS
Opens every sample document in the installed Word and fails if Word cannot open one, has to repair it or cannot export it to PDF.
.DESCRIPTION
Writes the sample documents (dotnet run --project samples/OpenXml.Wrappers.Samples) to a temporary folder, or takes the .docx files in
-Folder, then opens each read-only in a Word of its own (WINWORD.EXE /a /automation, DisplayAlerts off). A COM error while opening, a
document that Word reports as repaired and one that Word cannot export are failures: exporting needs the whole document laid out, which
opening does not, and a table far wider than the page can fail it. Differences between Word's Content.Text and the library's (through
tools/WordReference/Tool) are reported but do not fail the run. Normal.dotm is backed up first and checked afterwards.

Word is needed, so this runs on a developer's machine, not in CI.
.PARAMETER Folder
A folder of .docx files to open instead of the samples.
.PARAMETER WordPath
WINWORD.EXE. By default it is found through the App Paths registry key.
#>
[CmdletBinding()]
param(
    [string] $Folder,
    [string] $WordPath
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WordAutomation.ps1')

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$tool = Join-Path $PSScriptRoot 'Tool\WordReference.Tool.csproj'
if (-not $WordPath) {
    $WordPath = Get-WordPath
}

$ownFolder = -not $Folder
if ($ownFolder) {
    $Folder = Join-Path ([IO.Path]::GetTempPath()) ('WordReference-samples-' + [IO.Path]::GetRandomFileName())
    dotnet run --project (Join-Path $repositoryRoot 'samples\OpenXml.Wrappers.Samples') --framework net10.0 --configuration Release -- $Folder | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw 'The samples did not run.'
    }
}
$Folder = (Resolve-Path -LiteralPath $Folder).Path
$files = @(Get-ChildItem -LiteralPath $Folder -Filter '*.docx' | Sort-Object Name)
if ($files.Count -eq 0) {
    throw ('There are no .docx files in {0}.' -f $Folder)
}

dotnet run --project $tool --configuration Release -- text $Folder
if ($LASTEXITCODE -ne 0) {
    throw 'Reading the documents with the library failed.'
}

# Shows control characters, so that a difference in paragraph marks, cell marks or breaks can be seen.
function Format-Text([string] $Text) {
    return $Text.Replace("`r", '\r').Replace("`n", '\n').Replace("`t", '\t').Replace([string][char]7, '\a').Replace([string][char]11, '\v').Replace([string][char]12, '\f')
}

function Get-Difference([string] $Word, [string] $Library) {
    $length = [Math]::Min($Word.Length, $Library.Length)
    $index = 0
    while ($index -lt $length -and $Word[$index] -ceq $Library[$index]) {
        $index++
    }
    $start = [Math]::Max(0, $index - 20)
    return ('first difference at {0} (Word {1} characters, library {2}):{3}    Word:    "{4}"{3}    library: "{5}"' -f $index, $Word.Length,
        $Library.Length, [Environment]::NewLine, (Format-Text $Word.Substring($start, [Math]::Min(60, $Word.Length - $start))),
        (Format-Text $Library.Substring($start, [Math]::Min(60, $Library.Length - $start))))
}

$exportFolder = Join-Path ([IO.Path]::GetTempPath()) ('WordReference-export-' + [IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $exportFolder | Out-Null

$failures = [Collections.Generic.List[string]]::new()
$differences = [Collections.Generic.List[string]]::new()
$template = Backup-NormalTemplate
$word = Start-Word $WordPath
$application = $word.Application
try {
    foreach ($file in $files) {
        $count = $application.Documents.Count
        try {
            # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
            $document = $application.Documents.Open($file.FullName, $false, $true, $false)
        }
        catch {
            $failures.Add(('{0}: Word could not open it: {1}' -f $file.Name, $_.Exception.Message))
            continue
        }
        try {
            $caption = $document.ActiveWindow.Caption
            if ($caption -match 'Repaired' -or $application.Documents.Count -ne $count + 1) {
                $failures.Add(('{0}: Word repaired it ({1}).' -f $file.Name, $caption))
            }
            $wordText = $document.Content.Text
            $libraryTextPath = [IO.Path]::ChangeExtension($file.FullName, '.library.txt')
            if (-not (Test-Path -LiteralPath $libraryTextPath)) {
                $differences.Add(('{0}: the library could not read it.' -f $file.Name))
            }
            elseif ($wordText -cne ($libraryText = [IO.File]::ReadAllText($libraryTextPath))) {
                $differences.Add(('{0}: {1}' -f $file.Name, (Get-Difference $wordText $libraryText)))
            }
            try {
                $document.ExportAsFixedFormat((Join-Path $exportFolder ([IO.Path]::ChangeExtension($file.Name, '.pdf'))), 17)  # wdExportFormatPDF
            }
            catch {
                $failures.Add(('{0}: Word could not export it: {1}' -f $file.Name, $_.Exception.Message))
            }
            Write-Host ('Opened {0}' -f $file.Name)
        }
        finally {
            $document.Close(0)  # wdDoNotSaveChanges
        }
    }
}
finally {
    Stop-Word $word
}
Assert-NormalTemplateUnchanged $template
Remove-Item -LiteralPath $exportFolder -Recurse -Force

if ($ownFolder) {
    Remove-Item -LiteralPath $Folder -Recurse -Force
}
else {
    Get-ChildItem -LiteralPath $Folder -Filter '*.library.txt' | Remove-Item -Force
}

if ($differences.Count -gt 0) {
    Write-Host ''
    Write-Host ('Range.Text differs from the library''s for {0} of {1} document(s):' -f $differences.Count, $files.Count)
    $differences | ForEach-Object { Write-Host ('  ' + $_) }
}
if ($failures.Count -gt 0) {
    Write-Host ''
    $failures | ForEach-Object { Write-Host ('  ' + $_) -ForegroundColor Red }
    throw ('Word had {0} problem(s) with {1} document(s).' -f $failures.Count, $files.Count)
}
Write-Host ('Word opened and exported all {0} document(s) without repairing them.' -f $files.Count)
