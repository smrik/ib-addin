[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FormName,

    [switch]$ForceCloseExcel
)

# Turns a hand-authored UserForm .frm (inline "Begin MSForms.<Type> <name>"
# control blocks, no OleObjectBlob) into a real VBIDE-exported .frm/.frx pair.
#
# AGENTS.md prefers importing a real pair over rebuilding controls by script on
# every build, so run this once per hand-authored form and commit the result.
# Existing exported forms already have their .frx and do not need this.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$formSourcePath = Join-Path $repoRoot "vba-files\Form\$FormName.frm"
$formBinaryPath = [System.IO.Path]::ChangeExtension($formSourcePath, ".frx")

if (-not (Test-Path -LiteralPath $formSourcePath)) {
    throw "Form source not found: $formSourcePath"
}

$sourceText = Get-Content -LiteralPath $formSourcePath -Raw
if ($sourceText -match "OleObjectBlob") {
    throw "$FormName already exports an OleObjectBlob; it needs its existing .frx, not a rebuild."
}

$tempFolder = Join-Path ([System.IO.Path]::GetTempPath()) ("FAS_Form_" + [guid]::NewGuid().ToString("N"))
$exportFormPath = Join-Path $tempFolder "$FormName.frm"
$exportBinaryPath = [System.IO.Path]::ChangeExtension($exportFormPath, ".frx")

# .frm geometry is in twips; the VBIDE designer works in points.
function ConvertTo-Points {
    param([string]$TwipValue)
    return [math]::Round(([double]$TwipValue) / 20, 2)
}

function Get-BlockProperty {
    param([string]$Block, [string]$Name)

    $match = [regex]::Match($Block, "(?m)^\s*$Name\s+=\s+(.+?)\s*$")
    if (-not $match.Success) { return $null }

    $raw = $match.Groups[1].Value.Trim()
    # Strip a trailing VB comment such as: 1  'CenterOwner
    $raw = [regex]::Replace($raw, "\s+'.*$", "")
    return $raw.Trim('"')
}

$controlProgIds = @{
    "Label"         = "Forms.Label.1"
    "TextBox"       = "Forms.TextBox.1"
    "CommandButton" = "Forms.CommandButton.1"
    "CheckBox"      = "Forms.CheckBox.1"
    "ComboBox"      = "Forms.ComboBox.1"
    "ListBox"       = "Forms.ListBox.1"
    "Frame"         = "Forms.Frame.1"
    "OptionButton"  = "Forms.OptionButton.1"
}

# Form-level properties come from the outer Begin block, before the first control.
$firstControlIndex = $sourceText.IndexOf("Begin MSForms.")
if ($firstControlIndex -lt 0) {
    throw "$FormName contains no MSForms control definitions."
}
$formHeader = $sourceText.Substring(0, $firstControlIndex)

$formCaption = Get-BlockProperty -Block $formHeader -Name "Caption"
$clientWidth = Get-BlockProperty -Block $formHeader -Name "ClientWidth"
$clientHeight = Get-BlockProperty -Block $formHeader -Name "ClientHeight"

$controls = @()
foreach ($match in [regex]::Matches($sourceText, "(?ms)^\s{3}Begin\s+MSForms\.(\w+)\s+(\w+)\s*\r?\n(.*?)^\s{3}End\s*$")) {
    $controls += [pscustomobject]@{
        Type  = $match.Groups[1].Value
        Name  = $match.Groups[2].Value
        Block = $match.Groups[3].Value
    }
}

if ($controls.Count -eq 0) {
    throw "$FormName control blocks could not be parsed."
}

$optionIndex = $sourceText.IndexOf("Option Explicit")
if ($optionIndex -lt 0) {
    throw "$FormName does not contain Option Explicit."
}
$codeText = $sourceText.Substring($optionIndex).TrimEnd()

$excelProcesses = @(Get-Process Excel -ErrorAction SilentlyContinue)
if ($excelProcesses.Count -gt 0) {
    if (-not $ForceCloseExcel) {
        throw "Excel is already running. Close it first, or rerun with -ForceCloseExcel."
    }
    $excelProcesses | Stop-Process -Force
    Start-Sleep -Milliseconds 500
}

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

    if ($formCaption) { $designer.Caption = $formCaption }
    try { $designer.StartUpPosition = 1 } catch { }
    try {
        if ($clientWidth) { $designer.Width = ConvertTo-Points $clientWidth }
        if ($clientHeight) { $designer.Height = ConvertTo-Points $clientHeight }
    }
    catch {
        # Default sizing remains usable if design-time sizing is unavailable.
    }

    foreach ($control in $controls) {
        if (-not $controlProgIds.ContainsKey($control.Type)) {
            throw "Unsupported control type '$($control.Type)' on $($control.Name). Add it to `$controlProgIds."
        }

        $added = $designer.Controls.Add($controlProgIds[$control.Type], $control.Name, $true)

        foreach ($geometry in @("Left", "Top", "Width", "Height")) {
            $value = Get-BlockProperty -Block $control.Block -Name $geometry
            if ($null -ne $value) { $added.$geometry = ConvertTo-Points $value }
        }

        $caption = Get-BlockProperty -Block $control.Block -Name "Caption"
        if ($null -ne $caption) {
            try { $added.Caption = $caption } catch { }
        }

        foreach ($flag in @("MultiLine", "EnterKeyBehavior")) {
            $value = Get-BlockProperty -Block $control.Block -Name $flag
            if ($null -ne $value) {
                try { $added.$flag = ($value -eq "-1" -or $value -eq "True") } catch { }
            }
        }

        $scrollBars = Get-BlockProperty -Block $control.Block -Name "ScrollBars"
        if ($null -ne $scrollBars) {
            try { $added.ScrollBars = [int]$scrollBars } catch { }
        }

        $tabIndex = Get-BlockProperty -Block $control.Block -Name "TabIndex"
        if ($null -ne $tabIndex) {
            try { $added.TabIndex = [int]$tabIndex } catch { }
        }
    }

    if ($component.CodeModule.CountOfLines -gt 0) {
        $component.CodeModule.DeleteLines(1, $component.CodeModule.CountOfLines)
    }
    $component.CodeModule.InsertLines(1, $codeText)

    $component.Name = $FormName
    if ($component.Name -ne $FormName) {
        throw "Excel did not accept the component name '$FormName'."
    }

    $component.Export($exportFormPath)
    if (-not (Test-Path -LiteralPath $exportFormPath) -or -not (Test-Path -LiteralPath $exportBinaryPath)) {
        throw "Excel did not export a complete .frm/.frx pair for $FormName."
    }

    Copy-Item -LiteralPath $exportFormPath -Destination $formSourcePath -Force
    Copy-Item -LiteralPath $exportBinaryPath -Destination $formBinaryPath -Force

    Write-Host "Rebuilt $FormName as a .frm/.frx pair ($($controls.Count) controls)."
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tempFolder) {
        Remove-Item -LiteralPath $tempFolder -Recurse -Force
    }
}
