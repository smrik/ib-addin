[CmdletBinding()]
param(
    [string]$TemplatePath,
    [string]$OutputPath = ".\patriks-tab.xlam",
    [switch]$Visible,
    [switch]$Force,
    [switch]$ForceCloseExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$sourceRoot = Join-Path $repoRoot "vba-files"
$resolvedOutputPath = if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath
} else {
    Join-Path $repoRoot $OutputPath
}

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

function Get-LatestSourceWriteTime {
    $vbaFiles = Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
        Where-Object { $_.Extension -in @(".bas", ".cls", ".frm", ".frx") }

    $customUiRoot = Join-Path $repoRoot "ribbons"
    $customUiFiles = @()
    if (Test-Path -LiteralPath $customUiRoot) {
        $customUiFiles = Get-ChildItem -LiteralPath $customUiRoot -Recurse -File |
            Where-Object { $_.Extension -eq ".xml" }
    }

    $sourceFiles = @($vbaFiles) + @($customUiFiles)

    if (-not $sourceFiles) {
        return [datetime]::MinValue
    }

    return ($sourceFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
}

$outputFolder = Split-Path -Path $resolvedOutputPath -Parent
if (-not (Test-Path -LiteralPath $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

if ((Test-Path -LiteralPath $resolvedOutputPath) -and -not $Force) {
    $latestSourceWriteTime = Get-LatestSourceWriteTime
    $outputWriteTime = (Get-Item -LiteralPath $resolvedOutputPath).LastWriteTime
    if ($outputWriteTime -ge $latestSourceWriteTime) {
        Write-Host "Build skipped; add-in is already up to date: $resolvedOutputPath"
        exit 0
    }
}

Stop-ExcelIfRequested -ForceClose:$ForceCloseExcel

if (-not (Test-Path -LiteralPath $resolvedOutputPath)) {
    if ($TemplatePath) {
        if (-not (Test-Path -LiteralPath $TemplatePath)) {
            throw "Template path not found: $TemplatePath"
        }

        Copy-Item -LiteralPath $TemplatePath -Destination $resolvedOutputPath -Force
    } else {
        $excel = $null
        $workbook = $null

        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = [bool]$Visible
            $excel.DisplayAlerts = -not $Visible
            $excel.EnableEvents = $false
            $excel.ScreenUpdating = $false
            $excel.AskToUpdateLinks = $false
            $excel.AutomationSecurity = 3

            $workbook = $excel.Workbooks.Add()
            $workbook.IsAddin = $true
            $workbook.SaveAs($resolvedOutputPath, 55)
        }
        finally {
            if ($null -ne $workbook) {
                $workbook.Close($true)
            }

            if ($null -ne $excel) {
                $excel.Quit()
            }

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
    }
}

$importParams = @{
    WorkbookPath = $resolvedOutputPath
}

if ($Visible) {
    $importParams.Visible = $true
}

if ($ForceCloseExcel) {
    $importParams.ForceCloseExcel = $true
}

& (Join-Path $PSScriptRoot "import_vba.ps1") @importParams

$customUiPath = Join-Path $repoRoot "ribbons\customUI14.xml"
if (Test-Path -LiteralPath $customUiPath) {
    & (Join-Path $PSScriptRoot "embed_customui.ps1") -WorkbookPath $resolvedOutputPath -CustomUiPath $customUiPath
}

Write-Host "Built add-in: $resolvedOutputPath"
