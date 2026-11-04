#Requires -Version 7
<#
.SYNOPSIS
Regenerates the Word reference snapshots in OpenXml.Wrappers.Test/TestFiles/WordReference by running each scenario in the installed Word.
.DESCRIPTION
Starts a Word of its own (WINWORD.EXE /a /automation, so add-ins and the user's Normal.dotm are not loaded) and attaches to it through COM.
For the run it sets Application.UserName and UserInitials to 'Reference' and 'R' (and Options.UseLocalUserInfo, so that comments do not take
the signed-in account's name) and restores them afterwards. Normal.dotm is backed up first and checked afterwards; the Office user name in
the registry is checked too.

Each scenario in scenarios.ps1 runs on a new document with the current Office theme (Aptos) applied, which is saved as .docx to a temporary folder. Every XML part of each document is
then normalised by MarkupNormalizer (through tools/WordReference/Tool) and written to
OpenXml.Wrappers.Test/TestFiles/WordReference/<scenario>/<part>. README.md there records the Word build.

Word is needed only to regenerate the snapshots; the tests and CI read the committed files.
.PARAMETER Scenario
Only these scenarios (by name). By default every scenario runs and folders of scenarios that no longer exist are removed.
.PARAMETER WordPath
WINWORD.EXE. By default it is found through the App Paths registry key.
.PARAMETER KeepDocuments
Leaves the saved .docx files in the temporary folder and prints where it is.
#>
[CmdletBinding()]
param(
    [string[]] $Scenario,
    [string] $WordPath,
    [switch] $KeepDocuments
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'WordAutomation.ps1')

$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$outputRoot = Join-Path $repositoryRoot 'OpenXml.Wrappers.Test\TestFiles\WordReference'
$tool = Join-Path $PSScriptRoot 'Tool\WordReference.Tool.csproj'

$scenarios = & (Join-Path $PSScriptRoot 'scenarios.ps1')
$names = @(if ($Scenario) { $Scenario } else { $scenarios.Keys })
foreach ($name in $names) {
    if (-not $scenarios.Contains($name)) {
        throw ('There is no scenario called {0}.' -f $name)
    }
}
if (-not $WordPath) {
    $WordPath = Get-WordPath
}
$theme = Get-OfficeThemePath $WordPath

Write-Host (('This starts Word ({0}) on its own with /a, runs {1} scenario(s), and sets Word''s user name to ''Reference'' for the run ' +
    '(restored afterwards). Normal.dotm is backed up first and checked afterwards.') -f $WordPath, $names.Count)

# Build the normaliser first, so that a build failure does not leave Word running.
dotnet build $tool --configuration Release --nologo --verbosity quiet
if ($LASTEXITCODE -ne 0) {
    throw 'tools/WordReference/Tool did not build.'
}

$documents = Join-Path ([IO.Path]::GetTempPath()) ('WordReference-documents-' + [IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $documents | Out-Null

$template = Backup-NormalTemplate
$userInfo = Get-OfficeUserInfo
$word = Start-Word $WordPath
$application = $word.Application
try {
    $saved = [pscustomobject]@{
        UserName         = $application.UserName
        UserInitials     = $application.UserInitials
        UseLocalUserInfo = $application.Options.UseLocalUserInfo
    }
    try {
        $application.Options.UseLocalUserInfo = $true
        $application.UserName = 'Reference'
        $application.UserInitials = 'R'

        foreach ($name in $names) {
            Write-Host ('Running {0}' -f $name)
            $document = $application.Documents.Add()
            try {
                # A new document as Word 365 makes it: the current Office theme (Aptos), not the /a fallback (see Get-OfficeThemePath).
                $document.ApplyDocumentTheme($theme)
                & $scenarios[$name] $document
                $document.SaveAs2((Join-Path $documents ($name + '.docx')), 16)  # wdFormatDocumentDefault
            }
            finally {
                $document.Close(0)  # wdDoNotSaveChanges
            }
        }
        $version = $application.Version
        $build = $application.Build
    }
    finally {
        $application.UserName = $saved.UserName
        $application.UserInitials = $saved.UserInitials
        $application.Options.UseLocalUserInfo = $saved.UseLocalUserInfo
        if ($application.UserName -ne $saved.UserName -or $application.UserInitials -ne $saved.UserInitials) {
            Write-Warning 'Word''s user name or initials could not be restored.'
        }
    }
}
finally {
    Stop-Word $word
}

Assert-NormalTemplateUnchanged $template
$userInfoAfter = Get-OfficeUserInfo
if ($userInfoAfter.UserName -ne $userInfo.UserName -or $userInfoAfter.UserInitials -ne $userInfo.UserInitials) {
    throw ('The Office user name in the registry changed from ''{0}'' ({1}) to ''{2}'' ({3}).' -f $userInfo.UserName, $userInfo.UserInitials,
        $userInfoAfter.UserName, $userInfoAfter.UserInitials)
}
Write-Host 'The Office user name and initials are unchanged.'

if (-not $Scenario -and (Test-Path -LiteralPath $outputRoot)) {
    Get-ChildItem -LiteralPath $outputRoot -Directory | Remove-Item -Recurse -Force
}
New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
dotnet run --project $tool --configuration Release --no-build -- normalize $documents $outputRoot
if ($LASTEXITCODE -ne 0) {
    throw 'Normalising the documents failed.'
}
# The parts of a new document and the built-in style definitions that the library embeds, from the blank-document and built-in-styles runs.
dotnet run --project $tool --configuration Release --no-build -- templates $documents (Join-Path $repositoryRoot 'OpenXml.Wrappers\Word\Templates')
if ($LASTEXITCODE -ne 0) {
    throw 'Writing the library''s templates failed.'
}

$readme = @(
    '# Word reference snapshots'
    ''
    'Generated by `tools/WordReference/New-WordReferences.ps1`; do not edit by hand. Each folder is one scenario of'
    '`tools/WordReference/scenarios.ps1`, run in Word on a new document and saved as .docx. Every XML part of the result is normalised by'
    '`OpenXml.Wrappers.Test/Utils/MarkupNormalizer.cs`. `MarkupConformanceTest` compares what the library writes for the same operations.'
    ''
    ('Word version: {0}, build {1}.' -f $version, $build)
)
Set-Content -LiteralPath (Join-Path $outputRoot 'README.md') -Value $readme -Encoding utf8NoBOM

if ($KeepDocuments) {
    Write-Host ('The documents Word saved are in {0}' -f $documents)
}
else {
    Remove-Item -LiteralPath $documents -Recurse -Force
}
Write-Host ('Wrote {0} scenario(s) to {1} (Word {2}).' -f $names.Count, $outputRoot, $build)
