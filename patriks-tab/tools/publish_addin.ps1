[CmdletBinding()]
param(
    [string]$InstalledPath = "$env:APPDATA\Microsoft\AddIns\patriks-tab.xlam",
    [switch]$SkipSmokeTest,
    [switch]$Force,
    [switch]$ForceCloseExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$buildScript = Join-Path $PSScriptRoot "build_xlam.ps1"
$smokeScript = Join-Path $PSScriptRoot "smoke_test.ps1"
$builtAddInPath = Join-Path $repoRoot "patriks-tab.xlam"
$installedFolder = Split-Path -Path $InstalledPath -Parent
$historyFolder = Join-Path $installedFolder "history"

function Stop-ExcelIfRequested {
    param(
        [switch]$ForceClose
    )

    $excelProcesses = @(Get-Process Excel -ErrorAction SilentlyContinue)
    if ($excelProcesses.Count -eq 0) {
        return
    }

    if (-not $ForceClose) {
        $processList = ($excelProcesses | ForEach-Object { "PID $($_.Id)" }) -join ", "
        throw "Excel is already running ($processList). Close Excel first, or rerun with -ForceCloseExcel."
    }

    $excelProcesses | Stop-Process -Force
}

Stop-ExcelIfRequested -ForceClose:$ForceCloseExcel

$buildParams = @{}
if ($Force) {
    $buildParams.Force = $true
}
if ($ForceCloseExcel) {
    $buildParams.ForceCloseExcel = $true
}

& $buildScript @buildParams

if (-not $SkipSmokeTest) {
    $smokeParams = @{
        SkipBuild = $true
    }
    if ($ForceCloseExcel) {
        $smokeParams.ForceCloseExcel = $true
    }
    & $smokeScript @smokeParams
}

Stop-ExcelIfRequested -ForceClose:$ForceCloseExcel

if (-not (Test-Path -LiteralPath $installedFolder)) {
    New-Item -ItemType Directory -Path $installedFolder -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $historyFolder)) {
    New-Item -ItemType Directory -Path $historyFolder -Force | Out-Null
}

if (Test-Path -LiteralPath $InstalledPath) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupBaseName = [System.IO.Path]::GetFileNameWithoutExtension($InstalledPath)
    $backupPath = Join-Path $historyFolder ("{0}_{1}.xlam" -f $backupBaseName, $stamp)
    Copy-Item -LiteralPath $InstalledPath -Destination $backupPath -Force
    Write-Host "Backup created: $backupPath"
}

Copy-Item -LiteralPath $builtAddInPath -Destination $InstalledPath -Force
Write-Host "Published add-in: $InstalledPath"
