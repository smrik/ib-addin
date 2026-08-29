[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent

function Assert-SourceContains {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RelativePath,

        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $path = Join-Path $repoRoot $RelativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "$Description failed: source file not found: $RelativePath"
    }

    $source = Get-Content -LiteralPath $path -Raw
    if ($source -notmatch $Pattern) {
        throw "$Description failed: pattern '$Pattern' not found in $RelativePath"
    }
}

Assert-SourceContains "src\vba\modules\modCommandHelp.bas" 'Public Function FAS_GetCommandCatalog\(' "command catalog API"
Assert-SourceContains "src\vba\modules\modCommandHelp.bas" 'Public Function FAS_FilterCommandCatalog\(' "command filtering API"
Assert-SourceContains "src\vba\modules\modCommandHelp.bas" 'Public Function FAS_GetCommandCatalogCount\(' "command catalog count API"
Assert-SourceContains "src\vba\modules\modCommandHelp.bas" 'Public Function FAS_CountCommandCatalogMatches\(' "command catalog match-count API"

Assert-SourceContains "src\vba\forms\frmFASCommandPalette.frm" 'Attribute VB_Name = "frmFASCommandPalette"' "command palette form"
Assert-SourceContains "src\vba\forms\frmFASCommandPalette.frm" 'Private Sub txtCommand_Change\(' "live command filtering"
Assert-SourceContains "src\vba\forms\frmFASCommandPalette.frm" 'Private Sub txtCommand_KeyDown\(' "command palette keyboard handling"
Assert-SourceContains "src\vba\modules\modCommandPalette.bas" 'Public Sub FAS_ToggleCommandPalette\(' "command palette shortcut entry point"
if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "src\vba\forms\frmFASCommandPalette.frx"))) {
    throw "command palette form failed: matching frmFASCommandPalette.frx not found"
}

Assert-SourceContains "src\vba\modules\modPrecedentTrace.bas" 'Public Sub FAS_CycleDirectPrecedent\(' "precedent cycle macro"
Assert-SourceContains "src\vba\modules\modAddinLifecycle.bas" 'Application\.OnKey "\^%\{UP\}"' "precedent cycle shortcut"
Assert-SourceContains "src\vba\modules\modAddinLifecycle.bas" 'Application\.OnKey "\^%\{DOWN\}"' "precedent return shortcut"
Assert-SourceContains "src\vba\modules\modAddinLifecycle.bas" 'FAS_GetQualifiedMacroName\("FAS_ToggleCommandPalette"\)' "command palette shortcut binding"
Assert-SourceContains "src\vba\modules\modSubtotalTools.bas" 'Public Sub FAS_InsertSubtotalAboveSelection\(' "above subtotal macro"
Assert-SourceContains "src\vba\modules\modSubtotalTools.bas" 'Selection starts in row 1; no row above for subtotal formulas\.' "above subtotal row-bound guard"
Assert-SourceContains "src\vba\modules\modAddinLifecycle.bas" 'Application\.OnKey "\^\+P", FAS_GetQualifiedMacroName\("FAS_InsertSubtotalAboveSelection"\)' "above subtotal shortcut binding"
Assert-SourceContains "src\vba\modules\modPasteTools.bas" 'Public Sub FAS_CopyVisibleOnly\(' "visible-only copy macro"
Assert-SourceContains "src\vba\modules\modFormulaTools.bas" 'Public Sub FAS_WrapSelectedFormulasRound\(' "ROUND formula wrapper macro"
Assert-SourceContains "src\vba\modules\modFormulaTools.bas" 'Public Sub FAS_WrapSelectedFormulasIfError\(' "IFERROR formula wrapper macro"
Assert-SourceContains "src\vba\modules\modFormulaLinks.bas" 'Public Sub FAS_CopyLinks\(' "linked formula copy macro"
Assert-SourceContains "src\vba\modules\modFormulaLinks.bas" 'Public Sub FAS_PasteLinks\(' "linked formula paste macro"
Assert-SourceContains "src\vba\modules\modProductivityTools.bas" 'Public Sub FAS_CaptureCurrentColor\(' "Excel color capture macro"
Assert-SourceContains "src\vba\modules\modProductivityTools.bas" 'Public Sub FAS_ApplyCurrentColor\(' "Excel color apply macro"
Assert-SourceContains "src\vba\modules\modProductivityTools.bas" 'Public Sub FAS_InsertDelta\(' "Excel delta macro"
Assert-SourceContains "src\vba\modules\modProductivityTools.bas" 'Public Sub FAS_AutoFitWidthWithMargin\(' "Excel width macro"
Assert-SourceContains "src\vba\modules\modFormulaTools.bas" 'FAS_SetSetting "Formula\.RoundDigits"' "saved ROUND digit choice"
Assert-SourceContains "src\vba\modules\modFormulaTools.bas" 'FAS_GetSetting\("Formula\.RoundDigits"' "reused ROUND digit choice"
Assert-SourceContains "src\vba\modules\modCommandRegistry.bas" 'FAS_TryRunFormulaWrapperCommand' "formula wrapper command dispatch"
Assert-SourceContains "src\customUI\customUI14.xml" 'Ribbon_CopyVisibleOnly' "visible-only Ribbon callback"
Assert-SourceContains "src\customUI\customUI14.xml" 'Ribbon_PasteLinks' "linked paste Ribbon callback"

Write-Host "Source contract checks passed."
