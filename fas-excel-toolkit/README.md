# FAS_Toolkit

FAS_Toolkit is a Windows-only Excel VBA `.xlam` add-in for Financial Advisory / FDD work. The source of truth is the text-exported VBA project under `src/vba`, not the built binary add-in in `dist/`.

## Repo layout

```text
fas-excel-toolkit/
  AGENTS.md
  README.md
  dist/
  src/vba/modules/
  src/vba/classes/
  src/vba/forms/
  src/vba/documents/
  tests/
  tools/
```

## Prerequisites

- Windows with desktop Excel installed.
- VBA macros enabled in Excel.
- Excel Trust Center option enabled:
  `File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> Trust access to the VBA project object model`
- Close any stuck Excel instances before running build/import/export/smoke scripts.

## Build the add-in

From the repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1
```

The build creates or updates:

- `dist\FAS_Toolkit.xlam`

Optional template-based build:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1 -TemplatePath "C:\Path\To\Template.xlam"
```

Optional custom output:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1 -OutputPath ".\dist\MyToolkit.xlam"
```

If Excel is already running, the build will stop and ask you to close it. To intentionally close stale/hidden Excel automation processes first:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1 -ForceCloseExcel
```

If no VBA source files changed since the last build, the script skips the import step.

## Live edit loop in VS Code

If you want a save-and-test workflow similar to editing inside the VBA editor, this repo now supports a repo-native live sync loop for an already-open Excel workbook or add-in.

Recommended setup:

1. Open this repo in VS Code.
2. Install the recommended VS Code extensions when prompted. `xlcraft` is the suggested editor layer for VBA IntelliSense, linting, and formatting.
3. Open your dev workbook or dev add-in in Excel. For this loop, prefer a development copy rather than the installed roaming add-in.
4. In VS Code, run the task `Watch VBA (live sync)`.
5. When prompted, enter the full path to the workbook or add-in you have open in Excel.
6. Edit files under `src\vba\...` and save.
7. The watcher pushes only the changed module, class, form, or document module into the open workbook, so you can test immediately in Excel.

There is also a `Watch VBA (live sync + save workbook)` task if you want every synced change persisted to the workbook file right away. The default watcher does not save the binary workbook on every file save, which keeps the loop faster and avoids unnecessary workbook churn.

For fast verification after a save, run the VS Code task `Run macro in Excel` and enter a public no-argument macro such as `FAS_ShowNavigator`, `FAS_ToggleNavigator`, `FAS_PasteValuesOnly`, or `FAS_RegisterShortcuts`.

For a cleaner dev/test split, run the VS Code task `Start test Excel instance`. It starts a fresh Excel automation instance, closes any other auto-loaded add-in workbooks in that test session, opens only the dev add-in you point it at, sets a custom window caption such as `TEST INSTANCE`, runs `FAS_RegisterShortcuts` if available, and creates a blank workbook for testing.

If you want the full loop in one action, run `Start test Excel instance + watch VBA`. It launches the isolated test Excel window and starts the watcher against the same dev add-in path.

Live sync notes:

- The watcher syncs `modules`, `classes`, `forms`, and `documents` from `src\vba`.
- Form code sync expects a real `.frm` and `.frx` pair. For this repo that means keeping `frmFASSidebar.frm` and `frmFASSidebar.frx` together.
- Excel's macro dialog runs only `Public Sub` procedures with no required arguments. Helper `Function` procedures will not appear there even if sync succeeded.
- If you change startup or keyboard shortcut wiring, `Workbook_Open` does not fire again just because code changed. Run `FAS_RegisterShortcuts` manually or reopen the add-in to apply those changes.
- The `Start test Excel instance` script isolates development by launching a fresh Excel COM instance, closing other add-in workbooks in that test session, and opening only the chosen dev add-in. It does not uninstall your normal add-ins from the main Excel profile.
- Shortcut registration now qualifies the macro name with the current add-in workbook name, which avoids collisions when another loaded add-in happens to expose the same public macros.
- The watcher and macro runner now try to bind to the target workbook/add-in by file path first, which makes them much more reliable when both a normal Excel session and a separate test Excel instance are open.
- Structural cleanup is still better through the normal import/build flow. If you delete or rename source files, rebuild or re-import instead of relying on the watcher to remove old components from the workbook.
- The source of truth is still the text files in `src\vba`. Use the watcher for development speed, and keep `build_xlam.ps1`, `smoke_test.ps1`, and `publish_addin.ps1` as the reliable build/test/release path.

## Run the smoke test

From the repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\smoke_test.ps1
```

The smoke test:

- builds `dist\FAS_Toolkit.xlam`
- opens Excel with COM automation
- creates a temporary workbook with representative FAS sheet names
- loads the add-in
- runs core navigation, command-palette dispatch, paste smoke checks, and precedent helper checks
- exits with a non-zero code if any assertion fails

If the build is already current, run the faster smoke path:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\smoke_test.ps1 -SkipBuild -ForceCloseExcel
```

## Publish the installed add-in

The preferred local workflow is now:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\publish_addin.ps1 -ForceCloseExcel
```

This command:

- builds only if VBA source changed
- runs the smoke test
- creates a timestamped backup in `C:\Users\psmid\AppData\Roaming\Microsoft\AddIns\history`
- replaces `C:\Users\psmid\AppData\Roaming\Microsoft\AddIns\FAS_Toolkit.xlam`

## Command palette

`Ctrl+Shift+J` opens the FAS command palette. It works like the VS Code command palette: start typing to filter the complete command catalog, use Up/Down to move through results, press Enter to run, and press Escape to close. Parameterized entries such as `tb <search>` place the command prefix in the search box so you can finish the query before running it.

Current commands:

- `help`, `commands`, `help chart`, `help box`, `help paste`, `? ledger`, or similar searches open the complete command palette with an optional filter
- `p`, `p1`, `p2` jump to P&L slots
- `bs`, `bs2` jump to BS slots
- `t`, `t2` jump to TB/BB slots
- `g`, `g2` jump to GL slots
- `m`, `m2` jump to Mapping slots
- `val` pastes values
- `txt` pastes as text
- `rep` replaces text in selected formulas, with confirmation
- `pre`, `pre2`, `pre3` jump to direct precedents for the active formula cell
- `back` returns to the cell where the last precedent trace started
- `Ctrl+Alt+Up` starts at the active formula cell and cycles through its direct precedents, wrapping to the first after the last
- `Ctrl+Alt+Down` returns to the formula cell where precedent cycling started
- `tb cash`, `bb cash`, `acct cash`, or `account cash` search the resolved TB/BB sheet only
- `dk 489006`, `ledger revaluation`, `ix cash`, or `idx vendor` search the active FAS project ledger index
- `dktb cash` searches the active FAS project TB index
- `dkmap revenue` searches the active FAS project mapping index
- `gl vendor` searches the resolved GL sheet only
- `map revenue`, `mapping revenue`, `class revenue`, or `klas revenue` search the resolved Mapping/Classification sheet only
- `bm set review` or `mark review` saves the active cell as a local bookmark
- `bm review` jumps to a saved local bookmark
- `bm`, `bms`, or `bookmarks` lists local bookmarks
- `bmreset` or `resetbookmarks` clears local bookmarks for the active workbook
- `fcopy H10`, `formulacopy H10`, or `smartcopy H10` copies the selected range to a destination while moving same-sheet references inside the copied range and preserving outside/cross-sheet references as absolute references
- `sub` or `subtotal` inserts `SUBTOTAL(109,...)` formulas below the selected range
- `div` or `divider` inserts a colored divider sheet
- `drill` or `dd` creates a floating drilldown shape near the active cell with Mapping/TB/GL matches
- `drillclear` or `ddclear` removes the active sheet's FAS drilldown shape
- `chart fas`, `chart format`, or `chart clean` applies the FAS chart layout: title top-left, horizontal legend below title, Arial 8 pt, muted axes, value axis title `EUR 000`, all gridlines off, no vertical axis line/ticks, horizontal axis line without ticks, low horizontal/date labels, and expanded plot area
- `series fas`, `labels fas`, or `line fas` formats the selected chart series: data labels above and linked to source number format, label text matching the series color, line and marker fill matching the series color, and marker border de-emphasized
- `legend bottom`, `legend right`, `legend off`, `grid light`, or `grid off` applies focused chart cleanup commands
- `box`, `box full`, `box half`, `box third`, `box 2/3`, `box quarter`, `box row`, `box small`, `box title`, or `box subtitle` inserts a PowerPoint-sized outline box at the active cell using the FAS slide layout
- `size`, `size full`, `size half`, `size third`, `size 2/3`, `size quarter`, `size row`, or `size small` resizes selected shapes/charts to the same FAS slide layout sizes
- `layoutdebug` shows the loaded slide layout command version and calculated dimensions, useful after rebuilding/publishing the add-in
- `fixobj`, `fixshape`, or `float` sets selected shapes/images/charts to "Don't move or size with cells"
- `shapefix`, `nofloat`, or `fixshapes` sets workbook shapes to "Don't move or size with cells"
- `maps`, `mapinfo`, or `mappings` show the active workbook's local sheet mappings
- `mapreset` or `resetmaps` clear the active workbook's local sheet mappings after confirmation

Unknown text falls back to sheet search. Workbook-specific sheet slot mappings are stored locally under `%APPDATA%\FAS_Toolkit\Mappings`, and bookmarks are stored under `%APPDATA%\FAS_Toolkit\Bookmarks`. Neither is written into client workbooks.

Command discovery is built into the palette. Opening it with an empty search shows the complete catalog, while typing searches command names, aliases, categories, keywords, and descriptions. Commands such as `help chart`, `help box`, `help formula`, `help paste`, `help search`, or `? ledger` open the same palette already filtered. If typed text does not match a command or sheet, FAS suggests likely commands in the status bar.

Formula tools are intentionally conservative. Smart formula copy refuses formulas with external links, structured table references, `INDIRECT`, `OFFSET`, spilled refs, or 3D sheet refs instead of silently producing a risky rewrite.

Most commands report through Excel's status bar instead of modal popups. Popups are reserved for command input and risky confirmations such as overwriting cells, resetting mappings, or replacing formulas.

Drilldown uses one reusable floating shape per sheet. Running `drill` again replaces the previous FAS drilldown shape instead of inserting rows or leaving many objects behind.

Chart formatting commands are intentionally scoped to selected charts or the active chart. `chart fas` handles the whole chart surface, while `series fas` handles the selected series and its labels. If no individual series is selected, `series fas` falls back to all series in the active chart.

Slide layout boxes use the same Forvis Mazars wide-slide working area as the PowerPoint FAS tools: content width `32.18 cm`, content height `12.20 cm`, and a `0.83 cm` gutter. Calculated presets are `15.675 cm` for half width, `10.173 cm` for third width, `21.177 cm` for two-thirds width, `7.423 cm` for quarter width, and `5.685 cm` for half content height. Inserted boxes are outline-only and floating so they can be copied into PowerPoint or used as chart sizing guides without changing worksheet layout.

When testing these commands in your normal installed add-in, use `tools\publish_addin.ps1`, not only `tools\build_xlam.ps1`. The build script updates `dist\FAS_Toolkit.xlam`; the publish script replaces the installed roaming add-in at `C:\Users\psmid\AppData\Roaming\Microsoft\AddIns\FAS_Toolkit.xlam`.

If Excel shows old duplicate add-ins such as `Patriks_Toolkit` or `Patriks_Toolkit_V1`, close Excel and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\consolidate_excel_addins.ps1 -ForceCloseExcel
```

This refreshes the canonical `FAS_Toolkit.xlam` and archives the old duplicate files into the AddIns `history` folder. Keep only `FAS Toolkit` checked in Excel Add-ins afterward.

## Project Ledger Index

The project search layer keeps raw file paths and fast search indexes outside client workbooks under `%APPDATA%\FAS_Toolkit\Projects`. This lets the add-in search a local index even if the original workbook is closed, moved, renamed, or too slow to filter in Excel. Rebuilding the index still needs the source file path to be valid.

Install the helper scripts where the installed add-in can find them:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\install_search_tools.ps1
```

Install DuckDB separately by putting `duckdb.exe` in `%APPDATA%\FAS_Toolkit\Tools` or adding it to `PATH`. Without DuckDB, the builder still creates standardized staging CSVs, but it cannot create the fast `.duckdb`/Parquet index.

Create a project and register source files interactively:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\fas_project.ps1
```

The menu lets you create a project, set the active project, and add source files without remembering arguments. When adding a source, use `ledger` for DK/GL files, `tb` for trial balances, and `mapping` for classification/mapping files.

For scripted/bulk entry, the old argument style still works:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\fas_project.ps1 -Action New -ProjectName "Project Mime"

powershell -ExecutionPolicy Bypass -File .\tools\fas_project.ps1 -Action Add `
  -ProjectName "Project Mime" `
  -Kind ledger `
  -FilePath "C:\Path\To\DK 2025.xlsx" `
  -Company "Mediline" `
  -Year "2025" `
  -SheetName "DK-Mediline-25" `
  -HeaderRow 1
```

Repeat source entry for each year/company ledger file.

Build the local index:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\build_search_index.ps1 -ProjectName "Project Mime" -Kind all
```

Search from PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\query_search_index.ps1 -ProjectName "Project Mime" -Kind ledger -Query "489006 revaluation"
```

Search from Excel:

```text
Ctrl+Shift+J -> dk 489006 revaluation
Ctrl+Shift+J -> dktb cash
Ctrl+Shift+J -> dkmap revenue
```

Excel search commands try to call `duckdb.exe` directly for speed. If the add-in cannot locate the active project index or DuckDB executable, it falls back to the PowerShell query script. Rebuilds and source registration still use the PowerShell tools.

The builder auto-detects common ledger/TB columns such as `Konto`, `Debit/Debet`, `Credit/Kredit`, `Opis`, `Klasifikacija`, and date fields. For difficult files, save the target sheet as CSV first and register the CSV as the source. A richer mapping window for manual field mapping is the next natural layer.

## Install the add-in in Excel

1. Close Excel.
2. Build the add-in:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\tools\build_xlam.ps1
   ```

3. Open Excel.
4. Go to `File -> Options -> Add-ins`.
5. At the bottom, choose `Manage: Excel Add-ins`, then click `Go...`.
6. Click `Browse...`.
7. Select `dist\FAS_Toolkit.xlam` from this repo, or copy that file into your roaming add-ins folder first and browse there.
8. Confirm the add-in is checked in the add-ins list.

Your current roaming add-ins folder is:

`C:\Users\psmid\AppData\Roaming\Microsoft\AddIns`

If you want Excel to load this add-in from the roaming folder, copy `dist\FAS_Toolkit.xlam` there after closing Excel. This repo does not overwrite installed add-ins automatically.

## Export VBA from an existing workbook or add-in

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\export_vba.ps1 -WorkbookPath "C:\Path\To\WorkbookOrAddin.xlam"
```

Use `-Force` to allow overwriting existing exported source files:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\export_vba.ps1 -WorkbookPath "C:\Path\To\WorkbookOrAddin.xlam" -Force
```

## Import VBA into a workbook or add-in

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\import_vba.ps1 -WorkbookPath "C:\Path\To\WorkbookOrAddin.xlam"
```

Use `-Visible` if you want to watch Excel during import:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\import_vba.ps1 -WorkbookPath "C:\Path\To\WorkbookOrAddin.xlam" -Visible
```

## Add a new macro

1. Add the macro to a standard module under `src/vba/modules/`.
2. Use `Option Explicit`.
3. Make runnable shortcuts/macros a `Public Sub` with no required arguments.
4. Resolve the target workbook via `FAS_GetTargetWorkbook` unless the code intentionally manages the add-in itself.
5. If the macro is risky, add confirmation and a clear result message.
6. Rebuild with `tools\build_xlam.ps1`.
7. Re-run `tools\smoke_test.ps1` if the new macro affects navigation or active-workbook behavior.

## Assign a macro to the Quick Access Toolbar

1. Open Excel with the add-in installed.
2. Go to `File -> Options -> Quick Access Toolbar`.
3. Change `Choose commands from` to `Macros`.
4. Pick the add-in macro, such as `FAS_ShowNavigator`.
5. Click `Add >>`.
6. Optionally use `Modify...` to rename it or choose an icon.

## Notes on the sidebar form source

`src/vba/forms/frmFASSidebar.frm` contains the form code-behind and form metadata. A full exported VBA UserForm often also has a sibling binary `.frx` file. When a real `.frx` exists next to the `.frm`, the import/export tooling preserves it. For a fresh repo bootstrap, the import script can also generate the initial sidebar layout directly through Excel COM if no `.frx` is present yet.

## Known limitations

- Excel COM automation requires local Windows with desktop Excel installed.
- Import, export, build, and smoke-test scripts require VBA project trust access in Excel.
- Cloud or sandboxed Codex environments may not be able to run Excel COM smoke tests reliably.
- If Excel is already in a bad state, stuck, or showing hidden modal prompts, COM automation can fail until Excel is fully closed and restarted.
