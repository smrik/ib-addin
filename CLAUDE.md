# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

Conventions, product principles, and the hard-won Office/COM gotchas live in [AGENTS.md](AGENTS.md). Read that first; this file covers the command map and the cross-file architecture instead.

## What this workspace is

Office VBA productivity add-ins for financial advisory / investment banking work — the long-term goal is one customizable Excel + PowerPoint toolbar in the shape of UpSlide. Each project follows the same pattern: **text VBA source is the source of truth, the binary Office host is a build input/output that is never hand-edited.**

| Project | Host | Source | Artifact |
|---|---|---|---|
| `patriks-tab` | Excel — **canonical** | `vba-files/{Module,Class,Form}`, `ribbons/customUI14.xml` | `dist/patriks-tab.xlam` |
| `Instrumenta` | PowerPoint | `src/{Modules,Classes,Forms,CustomUI}` + `src/InstrumentaPowerpointToolbar.pptm` | `bin/*.ppam` |
| `Instrumenta-Keys` | standalone | `src/windows/*.ps1`, `src/mac/*` | packaged exe |
| `fas-excel-toolkit` | **retired** | see [RETIRED.md](fas-excel-toolkit/RETIRED.md) | — |

`fas-excel-toolkit` and `patriks-tab` were the same Excel codebase, forked both ways. They were merged into `patriks-tab` on 2026-08-29. Do not add features to `fas-excel-toolkit`; it survives only as a recovery copy until the ported smoke test passes on a machine with VBA project object model access enabled.

**`Instrumenta/README.md` is the upstream project's README and contradicts this workspace.** It says all coding happens inside the `.pptm` and the `.bas` files are "for reference only". Here the inverted rule applies: `src/Modules` etc. are authoritative and `instrumenta-build.ps1 -SyncVba` imports them into the PPTM. Files named `InstrumentaPowerpointToolbar.corrupt-*.pptm` / `.sync-failed-*.pptm` are recovery backups, never build inputs.

## Commands

### Source checks — no Office required, and the only thing CI runs

`.github/workflows/source-checks.yml` runs exactly these three on `windows-latest`. Every build/smoke/publish path below is local-only.

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\patriks-tab\tests\source_contract.ps1
```

```bash
python .\patriks-tab\tools\validate_vba_source.py
```

```bash
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Instrumenta\scripts\test-productivity-features.ps1
```

The VS Code task **Run source checks** runs all three in parallel.

### Local builds and host tests (require desktop Office, usually closed first)

**Prerequisite:** Excel → File → Options → Trust Center → Trust Center Settings → Macro Settings → **Trust access to the VBA project object model**. Without it every import/build script hangs with Excel alive but idle rather than failing cleanly. Check `HKCU:\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM` — it must be `1`.

```bash
powershell -ExecutionPolicy Bypass -File .\patriks-tab\tools\build_xlam.ps1
```

```bash
powershell -ExecutionPolicy Bypass -File .\patriks-tab\tools\smoke_test.ps1 -SkipBuild -ForceCloseExcel
```

```bash
powershell -ExecutionPolicy Bypass -File .\patriks-tab\tools\publish_addin.ps1 -ForceCloseExcel
```

`publish_addin.ps1` is the full flow: build-if-changed → smoke test → timestamped backup → replace the installed add-in in `%APPDATA%\Microsoft\AddIns`. Testing a change in your *installed* Excel requires publish, not just build.

**Two build paths exist — only one is authoritative.** `patriks-tab/config.json` points the XVBA VS Code extension at `patriks-tab/patriks-tab.xlam`, while `tools\build_xlam.ps1` builds `patriks-tab/dist/patriks-tab.xlam`. Same filename, different directories, which is exactly how the original fork started. **The PowerShell build is authoritative**; treat XVBA sync as a dev-loop convenience only, and never ship the root-level `.xlam`.

Note that `smoke_test.ps1` writes to your live `%APPDATA%\FAS_Toolkit` settings — it sets `Formula.RoundDigits=2` and then asserts a bare `round` reuses it. That persistence is required by the test, not incidental.

For the fast edit loop: `tools\start_test_excel.ps1` (isolated Excel instance) plus `tools\watch_vba.ps1 -WorkbookPath <path>` (pushes only the changed module into the open workbook).

PowerPoint:

```bash
powershell -ExecutionPolicy Bypass -File .\Instrumenta\scripts\build-dev.ps1 -SyncVba
```

`build-dev.ps1` / `build-prod.ps1` are thin wrappers over `instrumenta-build.ps1 -Mode dev|prod`. **`-SyncVba` is opt-in**: without it the build only refreshes the embedded customUI XML and exports whatever VBA already sits in the PPTM — edits under `src/Modules` will silently not reach the artifact. Dev and prod produce separate `.ppam` files and separate `HKCU:\...\Office\16.0\powerpoint\AddIns\*` keys, so both can be installed side by side.

### Release packaging

```bash
powershell -ExecutionPolicy Bypass -File .\tools\package-fas-addins.ps1 -Version 1.2.3
```

Installer/uninstaller templates live in `tools/fas-release/`; verify a built package with `tools\test-fas-release-package.ps1`. See [Instrumenta/docs/fas-release-packaging.md](Instrumenta/docs/fas-release-packaging.md).

### Running one test

There is no unit-test framework. The granular equivalents:

- either source-check script on its own (fastest, no Office)
- `patriks-tab\tools\smoke_test.ps1 -SkipBuild` — Excel COM assertions against a temp workbook
- `patriks-tab\tools\run_vba_macro.ps1` — invoke a single public macro in an open workbook
- `Instrumenta\scripts\test-fas-edit-mode.ps1`
- `Instrumenta-Keys\tests\test-edit-mode-routing.ps1` (source-only) vs `test-edit-mode-live.ps1` (needs PowerPoint)

When an Excel COM step misbehaves, isolate with one probe macro rather than re-running the whole smoke suite.

## Excel add-in architecture (`patriks-tab`)

Every tool is reachable two ways — the Ribbon and the command palette. Keeping both in sync is a contract, not a convention: `source_contract.ps1` asserts that each palette token has a dispatch case *and* a catalog entry, because the two trees had already drifted into Ribbon-only and palette-only tools before the merge.

```
Ctrl+Shift+J → modCommandPalette.FAS_ToggleCommandPalette
             → frmFASCommandPalette (thin UI, live filtering)
             → modCommandRegistry.FAS_RunCommand
                 ├─ FAS_TryRun*Command chain (help, IFS, format cycles,
                 │    shortcuts, formula wrappers, chart, slide layout, …)
                 └─ Select Case over command tokens + aliases
             → feature module (modNavigation, modPasteTools, modPrecedentTrace, …)
```

Supporting pieces that are easy to miss:

- `modCommandHelp.bas` — the searchable catalog (`FAS_GetCommandCatalog` / `FAS_FilterCommandCatalog`), what `help`, `commands` and `? <text>` render. Separate from the dispatch table in `modCommandRegistry`.
- `modShortcutManager.bas` — `FAS_GetShortcutDefinitions` is the single table of `(key, label, default, macro, aliases)`. `modAddinLifecycle` just calls `FAS_RegisterConfiguredShortcuts`; there are no hardcoded `Application.OnKey` calls. Users rebind through `frmFASSettings`. Changing shortcut wiring requires re-running registration — `Workbook_Open` does not re-fire on a code sync.
- `modWorkbookTarget.FAS_GetTargetWorkbook` — enforces the "act on the user's workbook, never `ThisWorkbook`" rule. Use it instead of `ActiveWorkbook` in new features.
- `modUndoManager.bas` — `FAS_UndoBeginRange` / `FAS_UndoCommit` / `FAS_UndoCancel`. Any command that writes to cells should wrap itself in these; the formula wrappers are the reference example.

**Adding a command** — the contract enforces this order: dispatch case in `modCommandRegistry.bas` (plus the token in `FAS_IsKnownCommand`) → catalog entry in `modCommandHelp.bas` → Ribbon callback if it deserves a button → assertion in `tests/source_contract.ps1`. The contract is a regex over source text, so renaming a public procedure breaks CI until the assertion is updated too.

**Local state never goes into client workbooks.** Mappings, bookmarks, project search indexes, and helper tools live under `%APPDATA%\FAS_Toolkit\{Mappings,Bookmarks,Projects,Tools}` as pipe-delimited text — encode user-controlled fields before writing.

**Layout quirk:** `vba-files/Class/` holds real class modules *and* the document modules (`ThisWorkbook.cls`, `Sheet1.cls`). Document modules must be code-replaced, never imported as components — `import_vba.ps1` and `sync_vba_file.ps1` both filter on `$documentModuleNames` to keep that straight.

`validate_vba_source.py` opens no Excel. It asserts every `onAction`/`getPressed`/`getEnabled`/`getVisible`/`getLabel` in `ribbons/customUI14.xml` resolves to a `Public Sub`/`Function` in `vba-files/`, and that no public name is declared twice.

## Toward one cross-host plugin

VBA cannot share modules between an `.xlam` and a `.ppam` at runtime, so "one plugin" means one shared *source* directory copied into both hosts at build time, host adapters for the parts touching `Workbook` vs `Presentation`, and one settings file both read.

The nearest seam is already visible: PowerPoint's `ModuleFASAlignment.bas` has a versioned profile system (zones, gutters, presets, `FAS-FM-wide-*` profile version) while Excel's `modSlideLayoutBoxes.bas` hardcodes the same numbers — `0.83 / 32.18 / 12.2` cm — as `Private Const`. The two also store settings differently: PowerPoint uses registry `SaveSetting`, Excel uses a file under `%APPDATA%`. Unifying those onto one shared profile file is the customization layer, and should use PowerPoint's model with Excel's file-based mechanism (registry does not travel between machines; a brand profile is exactly what a team shares).
