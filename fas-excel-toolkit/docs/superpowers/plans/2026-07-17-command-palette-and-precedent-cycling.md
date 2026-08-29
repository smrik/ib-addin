# FAS Command Palette and Precedent Cycling Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the incomplete command help popup with a VS Code-style searchable palette and add keyboard-driven cycling through direct formula precedents.

**Architecture:** `modCommandHelp` owns one complete command metadata catalog used by filtering, help, suggestions, and palette rows. A dedicated `frmFASCommandPalette` provides thin modeless UI over those module functions. `modPrecedentTrace` owns cycle state and reuses the existing conservative precedent parser, while `modAddinLifecycle` only registers workbook-qualified shortcuts.

**Tech Stack:** Excel VBA, MSForms UserForms, PowerShell, Excel COM automation.

## Global Constraints

- Source of truth remains `src/vba`; do not edit the XLAM binary directly.
- `Ctrl+Shift+J` opens the command palette and `Ctrl+Shift+N` remains the sheet navigator.
- `Ctrl+Alt+Up` cycles direct precedents and `Ctrl+Alt+Down` returns to the origin.
- Existing commands and aliases remain compatible.
- Formula parsing remains conservative and limited to supported same-workbook A1 references.
- Client workbooks receive no configuration sheets or hidden state.

---

### Task 1: Authoritative command catalog

**Files:**
- Modify: `src/vba/modules/modCommandHelp.bas`
- Modify: `src/vba/modules/modCommandRegistry.bas`
- Modify: `tools/smoke_test.ps1`

**Interfaces:**
- Produces: `FAS_GetCommandCatalog() As Collection`, `FAS_FilterCommandCatalog(queryText As String) As Collection`, `FAS_GetCommandCatalogCount() As Long`, and catalog items shaped as `Array(command, aliases, keywords, description, category)`.
- Consumes: existing command-family predicates and `FAS_RunCommand(commandText As String) As Boolean`.

- [ ] Add smoke assertions that `FAS_GetCommandCatalogCount` exceeds 12 and that searches for `chart`, `bookmark`, `precedent`, `layout`, and `mapping` return at least one result.
- [ ] Run `tools/smoke_test.ps1` against the current add-in and confirm failure because the catalog test helpers do not exist.
- [ ] Replace the partial private catalog with complete public catalog and filtering functions.
- [ ] Keep `FAS_TryRunCommandHelp` side-effect-free during command recognition and route interactive help to the new palette entry point.
- [ ] Make unknown-command suggestions consume the same filtered catalog.
- [ ] Make `FAS_GetShortcutBuilderCommandList` derive its list from the catalog rather than a separate hard-coded subset.
- [ ] Build and rerun the focused catalog smoke assertions.

### Task 2: Searchable command palette form

**Files:**
- Create: `src/vba/forms/frmFASCommandPalette.frm`
- Create: `src/vba/modules/modCommandPalette.bas`
- Modify: `tools/import_vba.ps1`
- Modify: `src/vba/modules/modNavigation.bas`

**Interfaces:**
- Consumes: `FAS_FilterCommandCatalog(queryText As String) As Collection` and `FAS_RunCommand(commandText As String) As Boolean`.
- Produces: `FAS_ShowCommandPalette(Optional initialQuery As String = vbNullString)`, `FAS_ToggleCommandPalette()`, and `FAS_ClearCommandPaletteReference()`.

- [ ] Add static source checks for the form, public palette entry point, and required controls `txtCommand`, `lstCommands`, and `lblEmpty`.
- [ ] Run the static checks and confirm failure because the form and module do not exist.
- [ ] Add a source-only `.frm` containing initialization, live filtering, Up/Down navigation, Enter execution, double-click execution, Escape dismissal, and modeless reference cleanup.
- [ ] Add `modCommandPalette.bas` to own the single modeless form instance and initial query.
- [ ] Extend `tools/import_vba.ps1` with `New-FASCommandPaletteForm`, generating the TextBox, two-column ListBox, and empty-state Label when no `.frx` exists.
- [ ] Change `FAS_JumpToSheet` into a compatibility wrapper that opens the palette.
- [ ] Build the XLAM and verify the form imports and VBA compiles.

### Task 3: Direct precedent cycling

**Files:**
- Modify: `src/vba/modules/modPrecedentTrace.bas`
- Modify: `tools/smoke_test.ps1`

**Interfaces:**
- Produces: `FAS_CycleDirectPrecedent()` and deterministic reset through `FAS_BackToTraceOrigin()`.
- Consumes: existing `FAS_ParsePrecedentRanges`, `FAS_SaveTraceOrigin`, and `FAS_ActivatePrecedentRange` helpers.

- [ ] Extend the COM smoke workbook with two same-sheet precedents and one cross-sheet precedent.
- [ ] Add calls proving first jump, next jump, wraparound, cross-sheet activation, and return to origin.
- [ ] Run the focused precedent smoke and confirm failure because `FAS_CycleDirectPrecedent` does not exist.
- [ ] Add module-level cycle source identity, current index, and cached precedent collection.
- [ ] Start or reset a cycle when the active formula source changes; otherwise advance and wrap.
- [ ] Reset cycle state on return, stale workbook/sheet/range, unsupported formula, or missing precedents.
- [ ] Rerun the precedent smoke assertions and confirm all expected addresses.

### Task 4: Shortcut lifecycle, documentation, and full verification

**Files:**
- Modify: `src/vba/modules/modAddinLifecycle.bas`
- Modify: `README.md`
- Modify: `tools/smoke_test.ps1`

**Interfaces:**
- Consumes: `FAS_ToggleCommandPalette`, `FAS_CycleDirectPrecedent`, and `FAS_BackToTraceOrigin`.
- Produces: workbook-qualified `Application.OnKey` bindings for `^+J`, `^%{UP}`, and `^%{DOWN}` plus matching unregistration.

- [ ] Add static assertions for all three shortcut registration and unregistration strings.
- [ ] Run static assertions and confirm the two precedent bindings are absent.
- [ ] Register `Ctrl+Shift+J` to the palette, `Ctrl+Alt+Up` to cycle, and `Ctrl+Alt+Down` to return.
- [ ] Update README command-palette and precedent documentation.
- [ ] Run source validation, build, and full COM smoke with Excel closed.
- [ ] Publish only after smoke passes, preserving the canonical `FAS_Toolkit.xlam` install and timestamped backup.
- [ ] Manually verify palette focus/filter/keyboard behavior and both precedent shortcuts in Excel.

## Execution Notes

This checkout does not contain a `.git` directory, so commit checkpoints are unavailable. Preserve each task as a small source diff and verify before moving to the next task.
