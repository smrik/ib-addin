[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [ValidateSet("Excel", "PowerPoint")]
    [string[]]$Include = @("Excel", "PowerPoint"),

    [string]$OutputRoot = "releases",

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-Sha256 {
    param([Parameter(Mandatory = $true)][string]$Path)
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).Hash.ToLowerInvariant()
}

function Assert-FileExists {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Required file not found: $Path"
    }
}

function Copy-RequiredFile {
    param(
        [Parameter(Mandatory = $true)][string]$Source,
        [Parameter(Mandatory = $true)][string]$Destination
    )

    Assert-FileExists -Path $Source
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Destination) | Out-Null
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    throw "Version must use semantic version format x.y.z, for example 1.0.0."
}

$RepoRoot = (Resolve-Path ".").Path
$PackageName = "FAS-Addins"
$ReleaseName = "$PackageName-$Version"
$ReleaseRoot = Join-Path $OutputRoot $ReleaseName
$ZipPath = Join-Path $OutputRoot "$ReleaseName.zip"

if ((Test-Path -LiteralPath $ReleaseRoot) -and (-not $Force)) {
    throw "Release folder already exists: $ReleaseRoot. Re-run with -Force to replace it."
}

if ((Test-Path -LiteralPath $ZipPath) -and (-not $Force)) {
    throw "Release ZIP already exists: $ZipPath. Re-run with -Force to replace it."
}

if (Test-Path -LiteralPath $ReleaseRoot) {
    Remove-Item -LiteralPath $ReleaseRoot -Recurse -Force
}

if (Test-Path -LiteralPath $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
}

New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null

$Definitions = @{
    Excel = [pscustomobject]@{
        id = "excel"
        name = "FAS_Toolkit.xlam"
        sourcePath = "fas-excel-toolkit/dist/FAS_Toolkit.xlam"
        packagePath = "Excel/FAS_Toolkit.xlam"
        installPath = "%APPDATA%\Microsoft\AddIns\FAS_Toolkit.xlam"
    }
    PowerPoint = [pscustomobject]@{
        id = "powerpoint"
        name = "InstrumentaPowerpointToolbar.ppam"
        sourcePath = "Instrumenta/bin/InstrumentaPowerpointToolbar.ppam"
        packagePath = "PowerPoint/InstrumentaPowerpointToolbar.ppam"
        installPath = "%APPDATA%\Microsoft\AddIns\InstrumentaPowerpointToolbar.ppam"
    }
}

$Artifacts = @()
$SupportFiles = @()

foreach ($Item in $Include) {
    $Definition = $Definitions[$Item]
    $Source = Join-Path $RepoRoot $Definition.sourcePath
    $Destination = Join-Path $ReleaseRoot $Definition.packagePath

    Copy-RequiredFile -Source $Source -Destination $Destination
    $Hash = Get-Sha256 -Path $Destination

    $Artifacts += [pscustomobject]@{
        id = $Definition.id
        name = $Definition.name
        sourcePath = $Definition.sourcePath
        packagePath = $Definition.packagePath
        installPath = $Definition.installPath
        sha256 = $Hash
    }
}

if ($Include -contains "Excel") {
    $SupportDefinitions = @(
        [pscustomobject]@{
            id = "fas_project_tool"
            sourcePath = "fas-excel-toolkit/tools/fas_project.ps1"
            packagePath = "Excel/Tools/fas_project.ps1"
            installPath = "%APPDATA%\FAS_Toolkit\Tools\fas_project.ps1"
        },
        [pscustomobject]@{
            id = "fas_search_index_builder"
            sourcePath = "fas-excel-toolkit/tools/build_search_index.ps1"
            packagePath = "Excel/Tools/build_search_index.ps1"
            installPath = "%APPDATA%\FAS_Toolkit\Tools\build_search_index.ps1"
        },
        [pscustomobject]@{
            id = "fas_search_index_query"
            sourcePath = "fas-excel-toolkit/tools/query_search_index.ps1"
            packagePath = "Excel/Tools/query_search_index.ps1"
            installPath = "%APPDATA%\FAS_Toolkit\Tools\query_search_index.ps1"
        }
    )

    foreach ($Definition in $SupportDefinitions) {
        $Source = Join-Path $RepoRoot $Definition.sourcePath
        $Destination = Join-Path $ReleaseRoot $Definition.packagePath

        Copy-RequiredFile -Source $Source -Destination $Destination
        $Hash = Get-Sha256 -Path $Destination

        $SupportFiles += [pscustomobject]@{
            id = $Definition.id
            sourcePath = $Definition.sourcePath
            packagePath = $Definition.packagePath
            installPath = $Definition.installPath
            sha256 = $Hash
        }
    }
}

Copy-RequiredFile -Source "tools/fas-release/Install-FASAddins.ps1" -Destination (Join-Path $ReleaseRoot "Install-FASAddins.ps1")
Copy-RequiredFile -Source "tools/fas-release/Uninstall-FASAddins.ps1" -Destination (Join-Path $ReleaseRoot "Uninstall-FASAddins.ps1")
Copy-RequiredFile -Source "tools/fas-release/Install-FASAddins.cmd" -Destination (Join-Path $ReleaseRoot "Install-FASAddins.cmd")
Copy-RequiredFile -Source "tools/fas-release/Uninstall-FASAddins.cmd" -Destination (Join-Path $ReleaseRoot "Uninstall-FASAddins.cmd")
Assert-FileExists -Path "tools/fas-release/README.template.md"

$IncludedList = foreach ($Artifact in $Artifacts) {
    "- $($Artifact.id): ``$($Artifact.name)``"
}

$Readme = Get-Content -LiteralPath "tools/fas-release/README.template.md" -Raw
$Readme = $Readme.Replace("{{VERSION}}", $Version)
$Readme = $Readme.Replace("{{INCLUDED_LIST}}", ($IncludedList -join [System.Environment]::NewLine))
$SupportList = if ($SupportFiles.Count -gt 0) {
    @(
        "- Excel search helper scripts are installed to ``%APPDATA%\FAS_Toolkit\Tools``"
    ) -join [System.Environment]::NewLine
} else {
    "- None"
}
$Readme = $Readme.Replace("{{SUPPORT_LIST}}", $SupportList)
Set-Content -LiteralPath (Join-Path $ReleaseRoot "README.md") -Value $Readme -Encoding UTF8

$Manifest = [pscustomobject]@{
    packageName = $PackageName
    version = $Version
    createdUtc = (Get-Date).ToUniversalTime().ToString("o")
    included = @($Include)
    artifacts = $Artifacts
    supportFiles = $SupportFiles
}

$Manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $ReleaseRoot "release.json") -Encoding UTF8

$ChecksumLines = foreach ($Artifact in $Artifacts) {
    "$($Artifact.sha256)  $($Artifact.packagePath)"
}
$InstallerHash = Get-Sha256 -Path (Join-Path $ReleaseRoot "Install-FASAddins.ps1")
$UninstallerHash = Get-Sha256 -Path (Join-Path $ReleaseRoot "Uninstall-FASAddins.ps1")
$ReadmeHash = Get-Sha256 -Path (Join-Path $ReleaseRoot "README.md")
$ManifestHash = Get-Sha256 -Path (Join-Path $ReleaseRoot "release.json")
$InstallCmdHash = Get-Sha256 -Path (Join-Path $ReleaseRoot "Install-FASAddins.cmd")
$UninstallCmdHash = Get-Sha256 -Path (Join-Path $ReleaseRoot "Uninstall-FASAddins.cmd")

$ChecksumLines += "$InstallerHash  Install-FASAddins.ps1"
$ChecksumLines += "$UninstallerHash  Uninstall-FASAddins.ps1"
$ChecksumLines += "$ReadmeHash  README.md"
$ChecksumLines += "$ManifestHash  release.json"
$ChecksumLines += "$InstallCmdHash  Install-FASAddins.cmd"
$ChecksumLines += "$UninstallCmdHash  Uninstall-FASAddins.cmd"

foreach ($SupportFile in $SupportFiles) {
    $ChecksumLines += "$($SupportFile.sha256)  $($SupportFile.packagePath)"
}

Set-Content -LiteralPath (Join-Path $ReleaseRoot "SHA256SUMS.txt") -Value $ChecksumLines -Encoding UTF8

Compress-Archive -LiteralPath $ReleaseRoot -DestinationPath $ZipPath -Force

Write-Host "Created release folder: $ReleaseRoot"
Write-Host "Created release ZIP: $ZipPath"
