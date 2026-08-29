[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$SkipOfficeRegistration,
    [string]$InstallRoot,
    [string]$StateRoot,
    [string]$SupportRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($InstallRoot)) {
    $InstallRoot = Join-Path $env:APPDATA "Microsoft\AddIns"
}
if ([string]::IsNullOrWhiteSpace($StateRoot)) {
    $StateRoot = Join-Path $env:APPDATA "FAS_Addins"
}
if ([string]::IsNullOrWhiteSpace($SupportRoot)) {
    $SupportRoot = Join-Path $env:APPDATA "FAS_Toolkit\Tools"
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

function Get-ArtifactInstallPath([object]$Artifact) {
    Join-Path $InstallRoot (Split-Path -Leaf ([string]$Artifact.installPath))
}

function Get-SupportInstallPath([object]$SupportFile) {
    Join-Path $SupportRoot (Split-Path -Leaf ([string]$SupportFile.installPath))
}

function Register-OfficeAddIn {
    param(
        [Parameter(Mandatory = $true)]$Artifact,
        [Parameter(Mandatory = $true)][string]$InstallPath
    )

    $office = switch ([string]$Artifact.id) {
        "excel" { "Excel"; break }
        "powerpoint" { "PowerPoint"; break }
        default { return $null }
    }

    $application = $null
    $addIns = $null
    $addIn = $null
    try {
        if ($office -eq "Excel") {
            $application = New-Object -ComObject Excel.Application
            $application.Visible = $false
            $application.DisplayAlerts = $false
            $addIns = $application.AddIns
            $addIn = Get-OfficeAddInByPath -AddIns $addIns -Path $InstallPath
            $wasAvailableBefore = ($null -ne $addIn)
            $wasEnabledBefore = if ($wasAvailableBefore) { [bool]$addIn.Installed } else { $false }
            $addedByPackage = $false
            if (-not $wasAvailableBefore) {
                $addIn = $addIns.Add($InstallPath, $false)
                $addedByPackage = $true
            }
            $addIn.Installed = $true
            if (-not [bool]$addIn.Installed) { throw "Excel did not report the add-in as installed." }
        } else {
            $application = New-Object -ComObject PowerPoint.Application
            $application.Visible = $false
            $addIns = $application.AddIns
            $addIn = Get-OfficeAddInByPath -AddIns $addIns -Path $InstallPath
            $wasAvailableBefore = ($null -ne $addIn)
            $wasEnabledBefore = if ($wasAvailableBefore) { [bool]$addIn.Loaded } else { $false }
            $addedByPackage = $false
            if (-not $wasAvailableBefore) {
                $addIn = $addIns.Add($InstallPath)
                $addedByPackage = $true
            }
            $addIn.Loaded = $true
            if (-not [bool]$addIn.Loaded) { throw "PowerPoint did not report the add-in as loaded." }
        }

        Write-Host "[FAS Add-ins] Registered $office add-in: $InstallPath"
        return [pscustomobject]@{
            office = $office
            path = $InstallPath
            registrationAttempted = $true
            wasAvailableBefore = $wasAvailableBefore
            wasEnabledBefore = $wasEnabledBefore
            addedByPackage = $addedByPackage
            error = $null
        }
    } catch {
        Write-Warning "Could not register the $office add-in automatically. The file was installed, but Office may need one manual Add-ins step. $($_.Exception.Message)"
        return [pscustomobject]@{
            office = $office
            path = $InstallPath
            registrationAttempted = $false
            wasAvailableBefore = $false
            wasEnabledBefore = $false
            addedByPackage = $false
            error = $_.Exception.Message
        }
    } finally {
        if ($null -ne $addIn) { Release-ComObject $addIn }
        if ($null -ne $addIns) { Release-ComObject $addIns }
        if ($null -ne $application) {
            try { $application.Quit() } catch { }
            Release-ComObject $application
        }
    }
}

$PackageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ManifestPath = Join-Path $PackageRoot "release.json"
if (-not (Test-Path -LiteralPath $ManifestPath)) {
    throw "release.json was not found next to Install-FASAddins.ps1. Extract the full ZIP before running the installer."
}

$Manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
$BackupRoot = Join-Path $StateRoot "backups"
$InstallStatePath = Join-Path $StateRoot "installed.json"
New-Item -ItemType Directory -Force -Path $StateRoot, $BackupRoot | Out-Null

$Artifacts = @($Manifest.artifacts)
if ($Artifacts.Count -eq 0) { throw "release.json does not contain any artifacts." }
$SupportFiles = @()
if ($Manifest.PSObject.Properties.Name -contains "supportFiles") {
    $SupportFiles = @($Manifest.supportFiles)
}

Write-Host "[FAS Add-ins] Validating package $($Manifest.packageName) $($Manifest.version)."

$ArtifactTargets = @()
foreach ($Artifact in $Artifacts) {
    $PackagePath = Join-Path $PackageRoot ([string]$Artifact.packagePath)
    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) { throw "Packaged artifact is missing: $($Artifact.packagePath)" }
    $ActualHash = Get-Sha256 $PackagePath
    if ($ActualHash -ne $Artifact.sha256.ToLowerInvariant()) {
        throw "Hash mismatch for $($Artifact.packagePath). Expected $($Artifact.sha256), got $ActualHash."
    }
    $InstallPath = Get-ArtifactInstallPath $Artifact
    if (Test-FileLocked $InstallPath) {
        throw "Cannot install because this add-in appears to be in use:`n$InstallPath`n`nPlease close Excel and PowerPoint completely, then run Install-FASAddins.cmd again."
    }
    $ArtifactTargets += [pscustomobject]@{
        definition = $Artifact
        packagePath = $PackagePath
        installPath = $InstallPath
    }
}

$SupportTargets = @()
foreach ($SupportFile in $SupportFiles) {
    $PackagePath = Join-Path $PackageRoot ([string]$SupportFile.packagePath)
    if (-not (Test-Path -LiteralPath $PackagePath -PathType Leaf)) { throw "Packaged support file is missing: $($SupportFile.packagePath)" }
    $ActualHash = Get-Sha256 $PackagePath
    if ($ActualHash -ne $SupportFile.sha256.ToLowerInvariant()) {
        throw "Hash mismatch for $($SupportFile.packagePath). Expected $($SupportFile.sha256), got $ActualHash."
    }
    $InstallPath = Get-SupportInstallPath $SupportFile
    if (Test-FileLocked $InstallPath) {
        throw "Cannot replace a locked support file: $InstallPath"
    }
    $SupportTargets += [pscustomobject]@{
        definition = $SupportFile
        packagePath = $PackagePath
        installPath = $InstallPath
    }
}

$Stamp = (Get-Date).ToUniversalTime().ToString("yyyyMMdd-HHmmss")
$BackupDir = Join-Path (Join-Path $BackupRoot $Manifest.version) $Stamp
New-Item -ItemType Directory -Force -Path $BackupDir | Out-Null
$InstalledArtifacts = @()
$InstalledSupportFiles = @()

foreach ($Target in $ArtifactTargets) {
    $Artifact = $Target.definition
    $BackupPath = $null
    $PreviousHash = $null
    if (Test-Path -LiteralPath $Target.installPath) {
        $PreviousHash = Get-Sha256 $Target.installPath
        $BackupPath = Join-Path $BackupDir "$($Artifact.id)-$(Split-Path -Leaf $Target.installPath)"
        Copy-Item -LiteralPath $Target.installPath -Destination $BackupPath -Force
        Write-Host "[FAS Add-ins] Backed up existing $($Artifact.id) add-in to $BackupPath."
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Target.installPath) | Out-Null
    Copy-Item -LiteralPath $Target.packagePath -Destination $Target.installPath -Force
    $InstalledHash = Get-Sha256 $Target.installPath
    $InstalledArtifacts += [pscustomobject]@{
        id = $Artifact.id
        name = $Artifact.name
        path = $Target.installPath
        packagePath = $Artifact.packagePath
        sha256 = $InstalledHash
        expectedSha256 = $Artifact.sha256
        installedFrom = "$($Manifest.packageName)-$($Manifest.version)"
        installedUtc = (Get-Date).ToUniversalTime().ToString("o")
        backupPath = $BackupPath
        previousSha256 = $PreviousHash
    }
    Write-Host "[FAS Add-ins] Installed $($Artifact.id) add-in to $($Target.installPath)."
}

foreach ($Target in $SupportTargets) {
    $SupportFile = $Target.definition
    $BackupPath = $null
    $PreviousHash = $null
    if (Test-Path -LiteralPath $Target.installPath) {
        $PreviousHash = Get-Sha256 $Target.installPath
        $BackupPath = Join-Path $BackupDir "support-$($SupportFile.id)-$(Split-Path -Leaf $Target.installPath)"
        Copy-Item -LiteralPath $Target.installPath -Destination $BackupPath -Force
        Write-Host "[FAS Add-ins] Backed up existing support file to $BackupPath."
    }

    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Target.installPath) | Out-Null
    Copy-Item -LiteralPath $Target.packagePath -Destination $Target.installPath -Force
    $InstalledHash = Get-Sha256 $Target.installPath
    $InstalledSupportFiles += [pscustomobject]@{
        id = $SupportFile.id
        path = $Target.installPath
        packagePath = $SupportFile.packagePath
        sha256 = $InstalledHash
        expectedSha256 = $SupportFile.sha256
        installedFrom = "$($Manifest.packageName)-$($Manifest.version)"
        installedUtc = (Get-Date).ToUniversalTime().ToString("o")
        backupPath = $BackupPath
        previousSha256 = $PreviousHash
    }
    Write-Host "[FAS Add-ins] Installed support file to $($Target.installPath)."
}

$Registrations = @()
if (-not $SkipOfficeRegistration) {
    foreach ($Target in $ArtifactTargets) {
        $Registrations += Register-OfficeAddIn -Artifact $Target.definition -InstallPath $Target.installPath
    }
} else {
    Write-Host "[FAS Add-ins] Office registration skipped by request."
}

[pscustomobject]@{
    packageName = $Manifest.packageName
    version = $Manifest.version
    installedUtc = (Get-Date).ToUniversalTime().ToString("o")
    releaseCreatedUtc = $Manifest.createdUtc
    sourcePackageRoot = $PackageRoot
    artifacts = $InstalledArtifacts
    supportFiles = $InstalledSupportFiles
    registrations = $Registrations
} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $InstallStatePath -Encoding UTF8

Write-Host "[FAS Add-ins] Installation complete."
Write-Host "[FAS Add-ins] Local install state written to $InstallStatePath."
