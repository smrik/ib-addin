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

# --- Command palette surface -------------------------------------------------

Assert-SourceContains "vba-files\Module\modCommandHelp.bas" 'Public Function FAS_GetCommandCatalog\(' "command catalog API"
Assert-SourceContains "vba-files\Module\modCommandHelp.bas" 'Public Function FAS_FilterCommandCatalog\(' "command filtering API"
Assert-SourceContains "vba-files\Module\modCommandHelp.bas" 'Public Function FAS_GetCommandCatalogCount\(' "command catalog count API"
Assert-SourceContains "vba-files\Module\modCommandHelp.bas" 'Public Function FAS_CountCommandCatalogMatches\(' "command catalog match-count API"

Assert-SourceContains "vba-files\Form\frmFASCommandPalette.frm" 'Attribute VB_Name = "frmFASCommandPalette"' "command palette form"
Assert-SourceContains "vba-files\Form\frmFASCommandPalette.frm" 'Private Sub txtCommand_Change\(' "live command filtering"
Assert-SourceContains "vba-files\Form\frmFASCommandPalette.frm" 'Private Sub txtCommand_KeyDown\(' "command palette keyboard handling"
Assert-SourceContains "vba-files\Module\modCommandPalette.bas" 'Public Sub FAS_ToggleCommandPalette\(' "command palette shortcut entry point"
if (-not (Test-Path -LiteralPath (Join-Path $repoRoot "vba-files\Form\frmFASCommandPalette.frx"))) {
    throw "command palette form failed: matching frmFASCommandPalette.frx not found"
}

# --- Configurable shortcuts --------------------------------------------------
# patriks-tab routes every shortcut through the configurable definition table
# rather than hardcoding Application.OnKey calls in the lifecycle module.

Assert-SourceContains "vba-files\Module\modAddinLifecycle.bas" 'FAS_RegisterConfiguredShortcuts' "configurable shortcut registration"
Assert-SourceContains "vba-files\Module\modAddinLifecycle.bas" 'FAS_UnregisterConfiguredShortcuts' "configurable shortcut teardown"
Assert-SourceContains "vba-files\Module\modShortcutManager.bas" 'Public Function FAS_GetShortcutDefinitions\(' "shortcut definition table"
Assert-SourceContains "vba-files\Module\modShortcutManager.bas" '"FAS_ToggleCommandPalette"' "command palette shortcut definition"
Assert-SourceContains "vba-files\Module\modShortcutManager.bas" '"FAS_CycleDirectPrecedent"' "precedent cycle shortcut definition"
Assert-SourceContains "vba-files\Module\modShortcutManager.bas" '"FAS_BackToTraceOrigin"' "precedent return shortcut definition"
Assert-SourceContains "vba-files\Module\modShortcutManager.bas" '"FAS_InsertSubtotalAboveSelection"' "above subtotal shortcut definition"
Assert-SourceContains "vba-files\Module\modSettings.bas" 'Public Sub FAS_ShowSettings\(' "settings dialog entry point"

# --- Core macros -------------------------------------------------------------

Assert-SourceContains "vba-files\Module\modPrecedentTrace.bas" 'Public Sub FAS_CycleDirectPrecedent\(' "precedent cycle macro"
Assert-SourceContains "vba-files\Module\modSubtotalTools.bas" 'Public Sub FAS_InsertSubtotalAboveSelection\(' "above subtotal macro"
Assert-SourceContains "vba-files\Module\modSubtotalTools.bas" 'Selection starts in row 1; no row above for subtotal formulas\.' "above subtotal row-bound guard"
Assert-SourceContains "vba-files\Module\modPasteTools.bas" 'Public Sub FAS_CopyVisibleOnly\(' "visible-only copy macro"
Assert-SourceContains "vba-files\Module\modFormulaLinks.bas" 'Public Sub FAS_CopyLinks\(' "linked formula copy macro"
Assert-SourceContains "vba-files\Module\modFormulaLinks.bas" 'Public Sub FAS_PasteLinks\(' "linked formula paste macro"
Assert-SourceContains "vba-files\Module\modProductivityTools.bas" 'Public Sub FAS_CaptureCurrentColor\(' "Excel color capture macro"
Assert-SourceContains "vba-files\Module\modProductivityTools.bas" 'Public Sub FAS_ApplyCurrentColor\(' "Excel color apply macro"
Assert-SourceContains "vba-files\Module\modProductivityTools.bas" 'Public Sub FAS_InsertDelta\(' "Excel delta macro"
Assert-SourceContains "vba-files\Module\modProductivityTools.bas" 'Public Sub FAS_AutoFitWidthWithMargin\(' "Excel width macro"

# --- Formula wrappers --------------------------------------------------------

Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'Public Sub FAS_WrapSelectedFormulasRound\(' "ROUND formula wrapper macro"
Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'Public Sub FAS_WrapSelectedFormulasIfError\(' "IFERROR formula wrapper macro"
Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'Public Function FAS_TryRunFormulaWrapperCommand\(' "formula wrapper command parser"
Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'Public Function FAS_IsFormulaWrapperCommand\(' "formula wrapper command recogniser"
Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'FAS_SetSetting FAS_ROUND_DIGITS_SETTING' "saved ROUND digit choice"
Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'FAS_GetSetting\(FAS_ROUND_DIGITS_SETTING' "reused ROUND digit choice"
Assert-SourceContains "vba-files\Module\modFormulaTools.bas" 'FAS_UndoBeginRange' "formula wrapper undo support"

# --- Palette reachability ----------------------------------------------------
# Every tool below is also on the Ribbon; these assertions keep the palette and
# the Ribbon in sync so a command cannot silently become Ribbon-only again.

Assert-SourceContains "vba-files\Module\modCommandRegistry.bas" 'FAS_TryRunFormulaWrapperCommand' "formula wrapper command dispatch"
Assert-SourceContains "vba-files\Module\modCommandRegistry.bas" 'FAS_IsFormulaWrapperCommand' "formula wrapper command recognition"

foreach ($commandCase in @(
        '"visible", "vis", "copyvisible", "visiblecopy"',
        '"linkcopy", "copylinks", "copylink"',
        '"linkpaste", "pastelinks", "pastelink"',
        '"colorcapture", "capturecolor", "getcolor"',
        '"colorapply", "applycolor", "usecolor"',
        '"delta", "deltasign"',
        '"autofit", "fitwidth", "width20", "autofitwidth"',
        '"settings", "options", "config"')) {
    Assert-SourceContains "vba-files\Module\modCommandRegistry.bas" ([regex]::Escape($commandCase)) "palette dispatch for $commandCase"
}

foreach ($catalogToken in @('"visible"', '"linkcopy"', '"linkpaste"', '"round"', '"round <digits>"',
        '"iferror"', '"delta"', '"autofit"', '"colorcapture"', '"colorapply"', '"settings"')) {
    Assert-SourceContains "vba-files\Module\modCommandHelp.bas" ("FAS_AddCommandHelp catalog, " + [regex]::Escape($catalogToken)) "catalog entry for $catalogToken"
}

# --- Ribbon ------------------------------------------------------------------

Assert-SourceContains "ribbons\customUI14.xml" 'Ribbon_CopyVisibleOnly' "visible-only Ribbon callback"
Assert-SourceContains "ribbons\customUI14.xml" 'Ribbon_PasteLinks' "linked paste Ribbon callback"

Write-Host "Source contract checks passed."
