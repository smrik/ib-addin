[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$formSourcePath = Join-Path $repoRoot "vba-files\Form\frmFASCommandPalette.frm"
$formBinaryPath = [System.IO.Path]::ChangeExtension($formSourcePath, ".frx")
$tempFolder = Join-Path ([System.IO.Path]::GetTempPath()) ("FAS_CommandPalette_" + [guid]::NewGuid().ToString("N"))
$exportFormPath = Join-Path $tempFolder "frmFASCommandPalette.frm"
$exportBinaryPath = [System.IO.Path]::ChangeExtension($exportFormPath, ".frx")

if (-not (Test-Path -LiteralPath $formSourcePath)) {
    throw "Command palette form source not found: $formSourcePath"
}

function Set-CodeModuleContent {
    param(
        [Parameter(Mandatory = $true)]
        $CodeModule,

        [Parameter(Mandatory = $true)]
        [string]$CodeText
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }
    $CodeModule.InsertLines(1, $CodeText)
}

$sourceLines = Get-Content -LiteralPath $formSourcePath
$optionIndex = [Array]::IndexOf($sourceLines, "Option Explicit")
if ($optionIndex -lt 0) {
    throw "Command palette form source does not contain Option Explicit."
}
$codeText = $sourceLines[$optionIndex..($sourceLines.Length - 1)] -join [Environment]::NewLine

New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Add()
    $component = $workbook.VBProject.VBComponents.Add(3)
    $designer = $component.Designer

    $designer.Caption = "FAS Command Palette"
    try {
        $designer.StartUpPosition = 1
    }
    catch {
        # Some VBIDE automation contexts do not expose StartUpPosition.
    }
    try {
        $designer.Width = 520
        $designer.Height = 290
    }
    catch {
        # Default sizing remains usable if design-time sizing is unavailable.
    }

    $txtCommand = $designer.Controls.Add("Forms.TextBox.1", "txtCommand", $true)
    $txtCommand.Left = 12
    $txtCommand.Top = 12
    $txtCommand.Width = 484
    $txtCommand.Height = 22

    $lstCommands = $designer.Controls.Add("Forms.ListBox.1", "lstCommands", $true)
    $lstCommands.Left = 12
    $lstCommands.Top = 42
    $lstCommands.Width = 484
    $lstCommands.Height = 210

    $lblEmpty = $designer.Controls.Add("Forms.Label.1", "lblEmpty", $true)
    $lblEmpty.Left = 18
    $lblEmpty.Top = 52
    $lblEmpty.Width = 360
    $lblEmpty.Height = 18
    $lblEmpty.Caption = "No commands match this search."
    $lblEmpty.Visible = $false

    Set-CodeModuleContent -CodeModule $component.CodeModule -CodeText $codeText
    $component.Name = "frmFASCommandPalette"
    if ($component.Name -ne "frmFASCommandPalette") {
        throw "Excel did not accept the command palette component name."
    }

    $component.Export($exportFormPath)
    if (-not (Test-Path -LiteralPath $exportFormPath) -or -not (Test-Path -LiteralPath $exportBinaryPath)) {
        throw "Excel did not export a complete .frm/.frx pair."
    }

    Copy-Item -LiteralPath $exportFormPath -Destination $formSourcePath -Force
    Copy-Item -LiteralPath $exportBinaryPath -Destination $formBinaryPath -Force
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }
    if ($null -ne $excel) {
        $excel.Quit()
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tempFolder) {
        Remove-Item -LiteralPath $tempFolder -Recurse -Force
    }
}

Write-Host "Exported command palette form pair:"
Write-Host "  $formSourcePath"
Write-Host "  $formBinaryPath"
