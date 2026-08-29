param(
    [string]$FormPath = (Join-Path $PSScriptRoot '..\src\Forms\FASEditModeForm.frm')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Close-ComObject {
    param($ComObject)

    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
        }
    }
}

function Get-FormCode {
    param([string]$Path)

    $text = Get-Content -LiteralPath $Path -Raw
    $optionIndex = $text.IndexOf('Option Explicit', [StringComparison]::OrdinalIgnoreCase)
    if ($optionIndex -lt 0) {
        throw "Option Explicit was not found in $Path."
    }

    return $text.Substring($optionIndex)
}

function Add-FormControl {
    param(
        $Designer,
        [string]$ProgId,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Caption = ''
    )

    $control = $Designer.Controls.Add($ProgId, $Name, $true)
    $control.Left = $Left
    $control.Top = $Top
    $control.Width = $Width
    $control.Height = $Height
    if (-not [string]::IsNullOrEmpty($Caption)) {
        $control.Caption = $Caption
    }
    return $control
}

$resolvedFormPath = [IO.Path]::GetFullPath($FormPath)
$outputFolder = Split-Path -Parent $resolvedFormPath
$codeText = Get-FormCode -Path $resolvedFormPath
$excel = $null
$workbook = $null
$component = $null
$designer = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AutomationSecurity = 3

    $workbook = $excel.Workbooks.Add()
    $component = $workbook.VBProject.VBComponents.Add(3)
    $component.Name = 'FASEditModeForm'
    $designer = $component.Designer
    $designer.Caption = 'FAS Edit Mode'
    try {
        $designer.Width = 252
        $designer.Height = 292
        $designer.StartUpPosition = 0
    }
    catch {
    }

    $modeLabel = Add-FormControl $designer 'Forms.Label.1' 'lblMode' 12 10 218 18 'EDIT MODE ACTIVE'
    $modeLabel.ForeColor = 0x008000
    [void](Add-FormControl $designer 'Forms.Label.1' 'lblZone' 12 32 218 16 'Zone: Content')
    [void](Add-FormControl $designer 'Forms.Label.1' 'lblSelection' 12 50 218 16 'No shape selected')
    $statusLabel = Add-FormControl $designer 'Forms.Label.1' 'lblStatus' 12 68 218 28 'Ready'
    $statusLabel.WordWrap = $true

    [void](Add-FormControl $designer 'Forms.Label.1' 'lblColumnsTitle' 12 104 48 16 'Columns')
    [void](Add-FormControl $designer 'Forms.ComboBox.1' 'cboColumns' 62 101 48 18)
    [void](Add-FormControl $designer 'Forms.Label.1' 'lblRowsTitle' 124 104 34 16 'Rows')
    [void](Add-FormControl $designer 'Forms.ComboBox.1' 'cboRows' 166 101 48 18)
    [void](Add-FormControl $designer 'Forms.Label.1' 'lblExactColumnTitle' 12 130 72 16 'Exact column')
    [void](Add-FormControl $designer 'Forms.ComboBox.1' 'cboExactColumn' 86 127 128 18)
    [void](Add-FormControl $designer 'Forms.Label.1' 'lblExactRowTitle' 12 156 72 16 'Exact row')
    [void](Add-FormControl $designer 'Forms.ComboBox.1' 'cboExactRow' 86 153 128 18)

    [void](Add-FormControl $designer 'Forms.CommandButton.1' 'cmdZoneContent' 12 181 66 22 'Content')
    [void](Add-FormControl $designer 'Forms.CommandButton.1' 'cmdZoneTitle' 84 181 58 22 'Title')
    [void](Add-FormControl $designer 'Forms.CommandButton.1' 'cmdZoneSubtitle' 148 181 66 22 'Subtitle')

    $keysLabel = Add-FormControl $designer 'Forms.Label.1' 'lblKeys' 12 210 218 34 "Numpad 1-9: place   Ctrl: fit`r`nArrows: move   Ctrl+arrows: resize   Esc: exit"
    $keysLabel.WordWrap = $true
    [void](Add-FormControl $designer 'Forms.CommandButton.1' 'cmdClearExact' 12 250 112 24 'Use keypad anchors')
    [void](Add-FormControl $designer 'Forms.CommandButton.1' 'cmdExit' 148 250 66 24 'Exit')

    $component.CodeModule.AddFromString($codeText)
    $component.Export($resolvedFormPath)

    $frxPath = [IO.Path]::ChangeExtension($resolvedFormPath, '.frx')
    if (-not (Test-Path -LiteralPath $resolvedFormPath) -or -not (Test-Path -LiteralPath $frxPath)) {
        throw 'Excel did not export both the .frm and .frx palette files.'
    }

    Write-Output "Exported tracked FAS Edit Mode form to $resolvedFormPath and $frxPath"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    Close-ComObject $designer
    Close-ComObject $component
    Close-ComObject $workbook
    Close-ComObject $excel
}
