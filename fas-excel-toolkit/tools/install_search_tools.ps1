[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$targetFolder = Join-Path $env:APPDATA "FAS_Toolkit\Tools"

if (-not (Test-Path -LiteralPath $targetFolder)) {
    New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
}

$toolNames = @(
    "fas_project.ps1",
    "build_search_index.ps1",
    "query_search_index.ps1"
)

foreach ($toolName in $toolNames) {
    $sourcePath = Join-Path (Join-Path $repoRoot "tools") $toolName
    $targetPath = Join-Path $targetFolder $toolName
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    Write-Host "Installed $toolName -> $targetPath"
}

Write-Host "FAS search tools installed to: $targetFolder"
