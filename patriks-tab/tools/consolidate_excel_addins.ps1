[CmdletBinding()]
param(
    [string]$CanonicalAddInName = "patriks-tab.xlam",
    [switch]$ForceCloseExcel,
    [switch]$WhatIfOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$addInsFolder = Join-Path $env:APPDATA "Microsoft\AddIns"
$historyFolder = Join-Path $addInsFolder "history"
$canonicalPath = Join-Path $addInsFolder $CanonicalAddInName
$builtPath = Join-Path $repoRoot "patriks-tab.xlam"
$duplicateNames = @(
    'Patriks_Toolkit.xlam',
    'Patriks_Toolkit_v1.xlam',
    '~$Patriks_Toolkit.xlam',
    '~$Patriks_Toolkit_v1.xlam'
)

function Stop-ExcelIfRequested {
    param([switch]$ForceClose)

    $excelProcesses = @(Get-Process Excel -ErrorAction SilentlyContinue)
    if ($excelProcesses.Count -eq 0) {
        return
    }

    if ($WhatIfOnly) {
        Write-Host "Excel is running: $($excelProcesses.Id -join ', ')"
        Write-Host "Real consolidation should be run after closing Excel, or with -ForceCloseExcel."
        return
    }

    if (-not $ForceClose) {
        $processList = ($excelProcesses | ForEach-Object { "PID $($_.Id)" }) -join ", "
        throw "Excel is already running ($processList). Close Excel first, or rerun with -ForceCloseExcel."
    }

    $excelProcesses | Stop-Process -Force
}

if (-not (Test-Path -LiteralPath $builtPath)) {
    throw "Built add-in not found: $builtPath. Run tools\publish_addin.ps1 first."
}

Stop-ExcelIfRequested -ForceClose:$ForceCloseExcel

if (-not (Test-Path -LiteralPath $addInsFolder)) {
    if ($WhatIfOnly) {
        Write-Host "Would create add-ins folder: $addInsFolder"
    } else {
        New-Item -ItemType Directory -Path $addInsFolder -Force | Out-Null
    }
}

if (-not (Test-Path -LiteralPath $historyFolder)) {
    if ($WhatIfOnly) {
        Write-Host "Would create history folder: $historyFolder"
    } else {
        New-Item -ItemType Directory -Path $historyFolder -Force | Out-Null
    }
}

if ($WhatIfOnly) {
    Write-Host "Would copy canonical add-in:"
    Write-Host "  From: $builtPath"
    Write-Host "  To:   $canonicalPath"
} else {
    Copy-Item -LiteralPath $builtPath -Destination $canonicalPath -Force
    Write-Host "Canonical add-in refreshed: $canonicalPath"
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
foreach ($name in $duplicateNames) {
    $duplicatePath = Join-Path $addInsFolder $name
    if (-not (Test-Path -LiteralPath $duplicatePath)) {
        continue
    }

    $archivePath = Join-Path $historyFolder ("{0}_{1}.xlam" -f ([System.IO.Path]::GetFileNameWithoutExtension($name)), $stamp)
    if ($WhatIfOnly) {
        Write-Host "Would archive duplicate:"
        Write-Host "  From: $duplicatePath"
        Write-Host "  To:   $archivePath"
    } else {
        Move-Item -LiteralPath $duplicatePath -Destination $archivePath -Force
        Write-Host "Archived duplicate: $duplicatePath"
    }
}

Write-Host "Keep only '$CanonicalAddInName' checked in Excel Add-ins."
