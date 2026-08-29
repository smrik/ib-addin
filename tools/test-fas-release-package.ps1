[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ReleaseZip,
    [string]$ScratchRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-File {
    param([Parameter(Mandatory = $true)][string]$Path)
    Assert-True -Condition (Test-Path -LiteralPath $Path -PathType Leaf) -Message "Missing required file: $Path"
}

Assert-File -Path $ReleaseZip

if ([string]::IsNullOrWhiteSpace($ScratchRoot)) {
    $ScratchRoot = Join-Path (Split-Path -Parent $PSScriptRoot) ".scratch"
}

$ZipLeaf = Split-Path -Leaf $ReleaseZip
Assert-True -Condition ($ZipLeaf -match '^FAS-Addins-(\d+\.\d+\.\d+)\.zip$') -Message "ZIP filename must match FAS-Addins-x.y.z.zip."
$VersionFromZip = $Matches[1]
$ExpectedRootName = "FAS-Addins-$VersionFromZip"

$TempRoot = Join-Path $ScratchRoot "fas-release-validate-$([System.Guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Force -Path $TempRoot | Out-Null

try {
    Expand-Archive -LiteralPath $ReleaseZip -DestinationPath $TempRoot -Force

    $TopDirs = @(Get-ChildItem -LiteralPath $TempRoot -Directory)
    Assert-True -Condition ($TopDirs.Count -eq 1) -Message "ZIP must contain exactly one top-level release folder."

    $ReleaseRoot = $TopDirs[0].FullName
    Assert-True -Condition ($TopDirs[0].Name -eq $ExpectedRootName) -Message "Top-level folder must be $ExpectedRootName."

    $ManifestPath = Join-Path $ReleaseRoot "release.json"
    $ChecksumsPath = Join-Path $ReleaseRoot "SHA256SUMS.txt"

    Assert-File -Path (Join-Path $ReleaseRoot "Install-FASAddins.ps1")
    Assert-File -Path (Join-Path $ReleaseRoot "Uninstall-FASAddins.ps1")
    Assert-File -Path (Join-Path $ReleaseRoot "Install-FASAddins.cmd")
    Assert-File -Path (Join-Path $ReleaseRoot "Uninstall-FASAddins.cmd")
    Assert-File -Path (Join-Path $ReleaseRoot "README.md")
    Assert-File -Path $ManifestPath
    Assert-File -Path $ChecksumsPath

    $Manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json

    Assert-True -Condition ($Manifest.packageName -eq "FAS-Addins") -Message "Manifest packageName must be FAS-Addins."
    Assert-True -Condition ($Manifest.version -eq $VersionFromZip) -Message "Manifest version must match ZIP filename."
    Assert-True -Condition ($null -ne ($Manifest.createdUtc -as [datetime])) -Message "Manifest createdUtc must be a valid date."
    Assert-True -Condition (@($Manifest.artifacts).Count -gt 0) -Message "Manifest must include at least one artifact."

    $ChecksumMap = @{}
    foreach ($Line in Get-Content -LiteralPath $ChecksumsPath) {
        if ($Line.Trim().Length -eq 0) { continue }
        Assert-True -Condition ($Line -match '^([a-fA-F0-9]{64})\s\s(.+)$') -Message "Invalid checksum line: $Line"
        $ChecksumMap[$Matches[2]] = $Matches[1].ToLowerInvariant()
    }

    foreach ($Artifact in @($Manifest.artifacts)) {
        Assert-True -Condition ($Artifact.id -in @("excel", "powerpoint")) -Message "Unsupported artifact id: $($Artifact.id)"
        Assert-True -Condition ($Artifact.sha256 -match '^[a-fA-F0-9]{64}$') -Message "Invalid artifact hash for $($Artifact.packagePath)."

        $ArtifactPath = Join-Path $ReleaseRoot $Artifact.packagePath
        Assert-File -Path $ArtifactPath

        $ActualHash = Get-Sha256 -Path $ArtifactPath
        Assert-True -Condition ($ActualHash -eq $Artifact.sha256.ToLowerInvariant()) -Message "Manifest hash mismatch for $($Artifact.packagePath)."
        Assert-True -Condition ($ChecksumMap.ContainsKey($Artifact.packagePath)) -Message "Missing checksum entry for $($Artifact.packagePath)."
        Assert-True -Condition ($ChecksumMap[$Artifact.packagePath] -eq $ActualHash) -Message "SHA256SUMS mismatch for $($Artifact.packagePath)."
    }

    $SupportFiles = @()
    if ($Manifest.PSObject.Properties.Name -contains "supportFiles") {
        $SupportFiles = @($Manifest.supportFiles)
    }
    foreach ($SupportFile in $SupportFiles) {
        Assert-True -Condition ($SupportFile.sha256 -match '^[a-fA-F0-9]{64}$') -Message "Invalid support-file hash for $($SupportFile.packagePath)."
        $SupportPath = Join-Path $ReleaseRoot $SupportFile.packagePath
        Assert-File -Path $SupportPath
        $ActualHash = Get-Sha256 -Path $SupportPath
        Assert-True -Condition ($ActualHash -eq $SupportFile.sha256.ToLowerInvariant()) -Message "Manifest hash mismatch for $($SupportFile.packagePath)."
        Assert-True -Condition ($ChecksumMap.ContainsKey($SupportFile.packagePath)) -Message "Missing checksum entry for $($SupportFile.packagePath)."
        Assert-True -Condition ($ChecksumMap[$SupportFile.packagePath] -eq $ActualHash) -Message "SHA256SUMS mismatch for $($SupportFile.packagePath)."
    }

    foreach ($RelativePath in @(
        "Install-FASAddins.ps1",
        "Uninstall-FASAddins.ps1",
        "Install-FASAddins.cmd",
        "Uninstall-FASAddins.cmd",
        "README.md",
        "release.json"
    )) {
        $FullPath = Join-Path $ReleaseRoot $RelativePath
        Assert-File -Path $FullPath
        Assert-True -Condition ($ChecksumMap.ContainsKey($RelativePath)) -Message "Missing checksum entry for $RelativePath."

        $ActualHash = Get-Sha256 -Path $FullPath
        Assert-True -Condition ($ChecksumMap[$RelativePath] -eq $ActualHash) -Message "SHA256SUMS mismatch for $RelativePath."
    }

    Write-Host "Release package validation passed: $ReleaseZip"
}
finally {
    if (Test-Path -LiteralPath $TempRoot) {
        Remove-Item -LiteralPath $TempRoot -Recurse -Force
    }
}
