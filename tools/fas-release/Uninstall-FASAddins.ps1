[CmdletBinding()]
param(
    [switch]$RestoreLatestBackup,
    [switch]$Force,
    [string]$StateRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($StateRoot)) {
    $StateRoot = Join-Path $env:APPDATA "FAS_Addins"
}

function Get-Sha256([string]$Path) {
    (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Test-FileLocked([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        return $false
    } catch {
        return $true
    }
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch { }
    }
}

function Get-OfficeAddInByPath {
    param(
        [Parameter(Mandatory = $true)]$AddIns,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    for ($index = 1; $index -le [int]$AddIns.Count; $index++) {
        $candidate = $null
        try {
            $candidate = $AddIns.Item($index)
            $candidatePath = [string]$candidate.FullName
            if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and
                ([System.IO.Path]::GetFullPath($candidatePath) -ieq $fullPath)) {
                return $candidate
            }
        } catch {
            if ($null -ne $candidate) { Release-ComObject $candidate }
        }
    }

    return $null
}

function Find-LatestBackup([string]$BackupRoot, [string]$Pattern) {
    if (-not (Test-Path -LiteralPath $BackupRoot)) { return $null }
    Get-ChildItem -LiteralPath $BackupRoot -Recurse -File -Filter $Pattern |
        Sort-Object LastWriteTimeUtc -Descending |
        Select-Object -First 1 |
        ForEach-Object { $_.FullName }
}

function Restore-OfficeRegistration {
    param(
        [Parameter(Mandatory = $true)]$Registration,
        [Parameter(Mandatory = $true)][bool]$RestoreEnabled
    )

    if (-not [bool]$Registration.registrationAttempted) { return }

    $application = $null
    $addIns = $null
    $addIn = $null
    try {
        if ([string]$Registration.office -eq "Excel") {
            $application = New-Object -ComObject Excel.Application
            $application.Visible = $false
            $application.DisplayAlerts = $false
            $addIns = $application.AddIns
            $addIn = Get-OfficeAddInByPath -AddIns $addIns -Path ([string]$Registration.path)
            if (($null -eq $addIn) -and [bool]$Registration.wasAvailableBefore -and (Test-Path -LiteralPath ([string]$Registration.path))) {
                $addIn = $addIns.Add(([string]$Registration.path), $false)
            }
            if ($null -ne $addIn) {
                $addIn.Installed = $RestoreEnabled
            }
        } else {
            $application = New-Object -ComObject PowerPoint.Application
            $application.Visible = $false
            $addIns = $application.AddIns
            $addIn = Get-OfficeAddInByPath -AddIns $addIns -Path ([string]$Registration.path)
            if (($null -eq $addIn) -and [bool]$Registration.wasAvailableBefore -and (Test-Path -LiteralPath ([string]$Registration.path))) {
                $addIn = $addIns.Add(([string]$Registration.path))
            }
            if ($null -ne $addIn) {
                $addIn.Loaded = $RestoreEnabled
                if (-not [bool]$Registration.wasAvailableBefore -and [bool]$Registration.addedByPackage) {
                    $addIns.Remove($addIn.Name)
                }
            }
        }
        Write-Host "[FAS Add-ins] Restored $($Registration.office) registration state."
    } catch {
        Write-Warning "Could not restore $($Registration.office) registration state. Office may retain a stale Add-ins entry. $($_.Exception.Message)"
    } finally {
        if ($null -ne $addIn) { Release-ComObject $addIn }
        if ($null -ne $addIns) { Release-ComObject $addIns }
        if ($null -ne $application) {
            try { $application.Quit() } catch { }
            Release-ComObject $application
        }
    }
}

$BackupRoot = Join-Path $StateRoot "backups"
$InstallStatePath = Join-Path $StateRoot "installed.json"

if (-not (Test-Path -LiteralPath $InstallStatePath)) {
    Write-Host "[FAS Add-ins] No local install manifest found at $InstallStatePath. Nothing to uninstall."
    return
}

$InstallState = Get-Content -LiteralPath $InstallStatePath -Raw | ConvertFrom-Json
$Artifacts = @($InstallState.artifacts)
$SupportFiles = @()
if ($InstallState.PSObject.Properties.Name -contains "supportFiles") {
    $SupportFiles = @($InstallState.supportFiles)
}
$Registrations = @()
if ($InstallState.PSObject.Properties.Name -contains "registrations") {
    $Registrations = @($InstallState.registrations)
}

foreach ($Artifact in $Artifacts) {
    $InstallPath = [string]$Artifact.path
    if (Test-FileLocked $InstallPath) {
        throw "Cannot uninstall because this add-in appears to be in use:`n$InstallPath`n`nPlease close Excel and PowerPoint completely, then run Uninstall-FASAddins.cmd again."
    }

    if (Test-Path -LiteralPath $InstallPath) {
        $ActualHash = Get-Sha256 $InstallPath
        $ExpectedHash = ([string]$Artifact.sha256).ToLowerInvariant()
        if (($ActualHash -ne $ExpectedHash) -and (-not $Force)) {
            throw "Refusing to remove a changed file:`n$InstallPath`n`nThe file hash no longer matches the install manifest. Re-run with -Force only if you are certain this file should be removed."
        }
        Remove-Item -LiteralPath $InstallPath -Force
        Write-Host "[FAS Add-ins] Removed $InstallPath."
    }

    if ($RestoreLatestBackup) {
        $BackupPath = [string]$Artifact.backupPath
        if ([string]::IsNullOrWhiteSpace($BackupPath) -or (-not (Test-Path -LiteralPath $BackupPath))) {
            $BackupPath = Find-LatestBackup $BackupRoot "$($Artifact.id)-$(Split-Path -Leaf $InstallPath)"
        }
        if (($null -ne $BackupPath) -and (Test-Path -LiteralPath $BackupPath)) {
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $InstallPath) | Out-Null
            Copy-Item -LiteralPath $BackupPath -Destination $InstallPath -Force
            Write-Host "[FAS Add-ins] Restored latest backup for $($Artifact.id) from $BackupPath."
        } else {
            Write-Host "[FAS Add-ins] No backup found for $($Artifact.id)."
        }
    }
}

foreach ($SupportFile in $SupportFiles) {
    $InstallPath = [string]$SupportFile.path
    if (Test-FileLocked $InstallPath) {
        throw "Cannot uninstall because this support file appears to be in use:`n$InstallPath"
    }

    if (Test-Path -LiteralPath $InstallPath) {
        $ActualHash = Get-Sha256 $InstallPath
        $ExpectedHash = ([string]$SupportFile.sha256).ToLowerInvariant()
        if (($ActualHash -ne $ExpectedHash) -and (-not $Force)) {
            throw "Refusing to remove a changed support file:`n$InstallPath`n`nRe-run with -Force only if you are certain this file should be removed."
        }
        Remove-Item -LiteralPath $InstallPath -Force
        Write-Host "[FAS Add-ins] Removed support file $InstallPath."
    }

    if ($RestoreLatestBackup) {
        $BackupPath = [string]$SupportFile.backupPath
        if ([string]::IsNullOrWhiteSpace($BackupPath) -or (-not (Test-Path -LiteralPath $BackupPath))) {
            $BackupPath = Find-LatestBackup $BackupRoot "support-$($SupportFile.id)-$(Split-Path -Leaf $InstallPath)"
        }
        if (($null -ne $BackupPath) -and (Test-Path -LiteralPath $BackupPath)) {
            New-Item -ItemType Directory -Force -Path (Split-Path -Parent $InstallPath) | Out-Null
            Copy-Item -LiteralPath $BackupPath -Destination $InstallPath -Force
            Write-Host "[FAS Add-ins] Restored support-file backup from $BackupPath."
        }
    }
}

foreach ($Registration in $Registrations) {
    $RestoreEnabled = $false
    if ($RestoreLatestBackup) {
        $RestoreEnabled = [bool]$Registration.wasEnabledBefore
    }
    Restore-OfficeRegistration -Registration $Registration -RestoreEnabled $RestoreEnabled
}

if ($RestoreLatestBackup) {
    [pscustomobject]@{
        packageName = $InstallState.packageName
        previousVersion = $InstallState.version
        restoredUtc = (Get-Date).ToUniversalTime().ToString("o")
        artifacts = $Artifacts
        supportFiles = $SupportFiles
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $StateRoot "restored-from-backup.json") -Encoding UTF8
}

Remove-Item -LiteralPath $InstallStatePath -Force
Write-Host "[FAS Add-ins] Uninstall complete."
