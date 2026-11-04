<#
.SYNOPSIS
    Checks that everything that carries the package version agrees with <Version> in the project file.

.DESCRIPTION
    <Version> in the packed project's csproj is the one source of the version. This script reads it (and what the SDK derives from it) with
    MSBuild and fails when anything disagrees:

    - Version must be a SemVer 2 version (1.2.3 or 1.2.3-preview.1).
    - AssemblyVersion and FileVersion must be the derived <major>.<minor>.<patch>.0, so nobody pins them by hand again.
    - CHANGELOG.md must have a "## [<Version>]" heading, or "## [Unreleased]" while the version is not released yet (with -Tag the version's
      own heading is required).
    - The version must not be lower than the newest release tag.
    - With -Tag: the tag must be <TagPrefix><Version> exactly.
    - With -Package: the .nupkg file name and its nuspec id and version must match.

    Used by .github/workflows/ci.yml and publish.yml; it can be run locally the same way.

.EXAMPLE
    pwsh eng/Check-Version.ps1
.EXAMPLE
    pwsh eng/Check-Version.ps1 -Package artifacts
.EXAMPLE
    pwsh eng/Check-Version.ps1 -Tag v1.1.0
#>
[CmdletBinding()]
param(
    # The packed project whose <Version> is checked.
    [string] $Project = 'OpenXml.Wrappers/OpenXml.Wrappers.csproj',
    # The prefix of this package's release tags.
    [string] $TagPrefix = 'v',
    # The release tag being published, e.g. v1.1.0.
    [string] $Tag,
    # A .nupkg file, or a folder that holds exactly one.
    [string] $Package,
    [string] $Changelog = 'CHANGELOG.md'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Push-Location $repoRoot
$failures = [System.Collections.Generic.List[string]]::new()

function Add-Failure([string] $message) {
    $failures.Add($message)
    if ($env:GITHUB_ACTIONS -eq 'true') { Write-Host "::error::$message" } else { Write-Host "FAIL: $message" -ForegroundColor Red }
}

$semVer = '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-((?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)(?:\.(?:0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$'

# Compares two SemVer 2 versions (build metadata ignored): -1, 0 or 1.
function Compare-SemVer([string] $a, [string] $b) {
    $ma = [regex]::Match($a, $semVer); $mb = [regex]::Match($b, $semVer)
    for ($i = 1; $i -le 3; $i++) {
        $c = ([long]$ma.Groups[$i].Value).CompareTo([long]$mb.Groups[$i].Value)
        if ($c -ne 0) { return $c }
    }
    $pa = $ma.Groups[4].Value; $pb = $mb.Groups[4].Value
    if ($pa -eq $pb) { return 0 }
    if ($pa -eq '') { return 1 }   # a release is higher than its prereleases
    if ($pb -eq '') { return -1 }
    $ia = $pa.Split('.'); $ib = $pb.Split('.')
    for ($i = 0; $i -lt [Math]::Min($ia.Length, $ib.Length); $i++) {
        $na = $ia[$i] -match '^\d+$'; $nb = $ib[$i] -match '^\d+$'
        if ($na -and $nb) { $c = ([long]$ia[$i]).CompareTo([long]$ib[$i]) }
        elseif ($na) { $c = -1 }
        elseif ($nb) { $c = 1 }
        else { $c = [string]::CompareOrdinal($ia[$i], $ib[$i]); $c = [Math]::Sign($c) }
        if ($c -ne 0) { return $c }
    }
    return $ia.Length.CompareTo($ib.Length)
}

try {
    # AssemblyVersion and FileVersion are only set by the SDK's GetAssemblyVersion target, which runs per target framework.
    $json = & dotnet msbuild $Project -nologo -p:TargetFramework=netstandard2.0 -target:GetAssemblyVersion `
        -getProperty:Version -getProperty:AssemblyVersion -getProperty:FileVersion -getProperty:PackageId
    if ($LASTEXITCODE -ne 0) { throw "dotnet msbuild failed for $Project`n$json" }
    $properties = ($json | Out-String | ConvertFrom-Json).Properties
    $version = $properties.Version
    $packageId = $properties.PackageId
    Write-Host "$packageId $version (AssemblyVersion $($properties.AssemblyVersion), FileVersion $($properties.FileVersion)) from $Project"

    if ($version -notmatch $semVer) {
        Add-Failure "Version '$version' in $Project is not a SemVer 2 version such as 1.2.3 or 1.2.3-preview.1."
    }
    else {
        $core = '{0}.{1}.{2}' -f $Matches[1], $Matches[2], $Matches[3]
        foreach ($name in 'AssemblyVersion', 'FileVersion') {
            if ($properties.$name -ne "$core.0") {
                Add-Failure "$name is '$($properties.$name)' but should be '$core.0', derived from Version. Remove any $name set by hand."
            }
        }

        # CHANGELOG
        $changelogText = Get-Content -Raw $Changelog
        $hasVersionHeading = $changelogText -match ('(?m)^## \[' + [regex]::Escape($version) + '\]')
        $hasUnreleased = $changelogText -match '(?m)^## \[Unreleased\]'
        if ($Tag -and -not $hasVersionHeading) {
            Add-Failure "$Changelog has no '## [$version]' heading for the release."
        }
        elseif (-not $hasVersionHeading -and -not $hasUnreleased) {
            Add-Failure "$Changelog has neither a '## [$version]' heading nor an '## [Unreleased]' one."
        }

        # Not lower than the newest release tag.
        $tags = @(& git tag --list "$TagPrefix*" 2>$null)
        if ($LASTEXITCODE -ne 0) { Write-Host 'git tags are not available; the newest-tag check was skipped.' }
        $released = @($tags | ForEach-Object { $_.Substring($TagPrefix.Length) } | Where-Object { $_ -match $semVer })
        if ($released.Count -eq 0) {
            Write-Host "No $TagPrefix* tags found; the newest-tag check was skipped (CI needs a checkout with tags)."
        }
        else {
            $newest = $released[0]
            foreach ($r in $released) { if ((Compare-SemVer $r $newest) -gt 0) { $newest = $r } }
            if ((Compare-SemVer $version $newest) -lt 0) {
                Add-Failure "Version $version is lower than the newest release tag $TagPrefix$newest."
            }
            else { Write-Host "Newest release tag: $TagPrefix$newest" }
        }

        # The tag being published.
        if ($Tag) {
            if ($Tag -ne "$TagPrefix$version") {
                Add-Failure "Tag '$Tag' does not match Version $version in $Project (expected '$TagPrefix$version')."
            }
        }

        # The package that was built.
        if ($Package) {
            $nupkgs = @(if (Test-Path $Package -PathType Container) { Get-ChildItem $Package -Filter *.nupkg } else { Get-Item $Package })
            if ($nupkgs.Count -ne 1) {
                Add-Failure "Expected one .nupkg at '$Package' but found $($nupkgs.Count)."
            }
            else {
                $nupkg = $nupkgs[0]
                if ($nupkg.Name -ne "$packageId.$version.nupkg") {
                    Add-Failure "Package file '$($nupkg.Name)' should be '$packageId.$version.nupkg'."
                }
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                $zip = [System.IO.Compression.ZipFile]::OpenRead($nupkg.FullName)
                try {
                    $entry = $zip.Entries | Where-Object { $_.FullName -notlike '*/*' -and $_.Name -like '*.nuspec' } | Select-Object -First 1
                    $reader = [System.IO.StreamReader]::new($entry.Open())
                    try { [xml] $nuspec = $reader.ReadToEnd() } finally { $reader.Dispose() }
                }
                finally { $zip.Dispose() }
                $metadata = $nuspec.package.metadata
                if ($metadata.id -ne $packageId -or $metadata.version -ne $version) {
                    Add-Failure "The nuspec in '$($nupkg.Name)' says $($metadata.id) $($metadata.version), not $packageId $version."
                }
            }
        }
    }
}
finally {
    Pop-Location
}

if ($failures.Count -gt 0) {
    Write-Host "$($failures.Count) version check(s) failed."
    exit 1
}
Write-Host 'Version checks passed.'
