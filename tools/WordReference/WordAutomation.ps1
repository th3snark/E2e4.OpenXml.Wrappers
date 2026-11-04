<#
.SYNOPSIS
Helpers shared by the Word reference scripts: finding, starting and quitting Word, and guarding Normal.dotm. Dot-source it.
#>

Set-StrictMode -Version Latest

if (-not ('WordReference.Com' -as [type])) {
    Add-Type -Namespace WordReference -Name Com -MemberDefinition @'
[DllImport("ole32.dll", CharSet = CharSet.Unicode)]
private static extern int CLSIDFromProgID(string progId, out Guid clsid);

[DllImport("oleaut32.dll")]
private static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object instance);

// Marshal.GetActiveObject is not available in PowerShell 7 (.NET), so this does the same.
public static object TryGetActiveObject(string progId)
{
    Guid clsid;
    if (CLSIDFromProgID(progId, out clsid) != 0) return null;
    object instance;
    return GetActiveObject(ref clsid, IntPtr.Zero, out instance) == 0 ? instance : null;
}
'@
}

function Get-WordPath {
    $key = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe'
    $path = (Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue).'(default)'
    if (-not $path -or -not (Test-Path -LiteralPath $path)) {
        throw 'Word (WINWORD.EXE) was not found. Pass -WordPath.'
    }
    return $path
}

# Starts a Word of its own with /a (no add-ins, and the built-in Normal template rather than the user's Normal.dotm) and /automation, and
# attaches to it through COM. Refuses when Word is already running, because COM would then hand back that instance instead.
function Start-Word([string] $WordPath) {
    if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        throw 'Word is already running. Close every Word window first, so that the script attaches to the Word it starts and not to yours.'
    }
    $process = Start-Process -FilePath $WordPath -ArgumentList '/a', '/automation' -PassThru
    $deadline = (Get-Date).AddSeconds(90)
    $application = $null
    while (-not $application) {
        if ($process.HasExited) {
            throw ('Word exited with code {0} before it could be reached.' -f $process.ExitCode)
        }
        if ((Get-Date) -gt $deadline) {
            Stop-Process -Id $process.Id -Force
            throw 'Word did not register for automation within 90 seconds.'
        }
        Start-Sleep -Milliseconds 500
        $application = [WordReference.Com]::TryGetActiveObject('Word.Application')
    }
    $application.DisplayAlerts = 0  # wdAlertsNone
    # Word can start with an empty Document1, which the first document opened then replaces; close it so that counts are predictable.
    foreach ($document in @($application.Documents)) {
        $document.Close(0)  # wdDoNotSaveChanges
    }
    return [pscustomobject]@{ Application = $application; Process = $process }
}

function Stop-Word($Word) {
    try {
        foreach ($document in @($Word.Application.Documents)) {
            $document.Close(0)  # wdDoNotSaveChanges
        }
        $Word.Application.Quit(0)
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Word.Application)
        if (-not $Word.Process.WaitForExit(30000)) {
            Stop-Process -Id $Word.Process.Id -Force
        }
    }
}

# Copies Normal.dotm aside and remembers its hash, so that Assert-NormalTemplateUnchanged can prove the run left it alone.
function Backup-NormalTemplate {
    $path = Join-Path $env:APPDATA 'Microsoft\Templates\Normal.dotm'
    if (-not (Test-Path -LiteralPath $path)) {
        Write-Host 'There is no Normal.dotm to back up.'
        return [pscustomobject]@{ Path = $path; Backup = $null; Hash = $null }
    }
    $folder = Join-Path ([IO.Path]::GetTempPath()) ('WordReference-' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType Directory -Path $folder -Force | Out-Null
    $backup = Join-Path $folder 'Normal.dotm'
    Copy-Item -LiteralPath $path -Destination $backup
    Write-Host ('Backed up Normal.dotm to {0}' -f $backup)
    return [pscustomobject]@{ Path = $path; Backup = $backup; Hash = (Get-FileHash -LiteralPath $path).Hash }
}

function Assert-NormalTemplateUnchanged($Template) {
    $exists = Test-Path -LiteralPath $Template.Path
    if (-not $Template.Hash) {
        if ($exists) {
            throw ('Word created {0} during the run.' -f $Template.Path)
        }
        return
    }
    if ($exists -and (Get-FileHash -LiteralPath $Template.Path).Hash -eq $Template.Hash) {
        Write-Host 'Normal.dotm is unchanged.'
        return
    }
    Copy-Item -LiteralPath $Template.Backup -Destination $Template.Path -Force
    throw ('Normal.dotm changed during the run; it has been restored from {0}.' -f $Template.Backup)
}

# The user information Office keeps in the registry, which Application.UserName and UserInitials read and write.
function Get-OfficeUserInfo {
    $values = Get-ItemProperty -LiteralPath 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo' -ErrorAction SilentlyContinue
    return [pscustomobject]@{
        UserName     = if ($values -and $values.PSObject.Properties['UserName']) { $values.UserName } else { $null }
        UserInitials = if ($values -and $values.PSObject.Properties['UserInitials']) { $values.UserInitials } else { $null }
    }
}
