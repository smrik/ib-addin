# AGENTS.md

This repo contains source code for an Excel VBA .xlam add-in.

Rules:
- The source of truth is src/vba, not the .xlam binary.
- Do not edit binary Excel files directly unless explicitly asked.
- VBA modules go in src/vba/modules as .bas files.
- VBA classes go in src/vba/classes as .cls files.
- VBA forms go in src/vba/forms as .frm files.
- Public macros intended for Excel shortcuts must be Public Sub procedures with no required arguments.
- All modules must use Option Explicit.
- All macros must act on the active user workbook, not ThisWorkbook, unless the code is intentionally managing the add-in itself.
- Risky operations must ask for confirmation and report what changed.
- Prefer simple, explicit VBA over clever abstractions.
- Keep UI logic thin. Put matching/search/replacement logic in normal modules so it can be tested.
- Tests should be runnable from PowerShell on Windows using Excel COM automation where possible.
- For the intended interactive development workflow, follow [docs/dev-instance.md](docs/dev-instance.md). That file defines how the isolated Excel test instance, watcher, macro runner, and shortcut registration should work.

Build/debugging notes learned the hard way:
- Do not let Excel COM builds run while normal Excel is already open. Close Excel first, or use the script `-ForceCloseExcel` switch intentionally.
- Hidden Excel processes can keep `.xlam` files locked and make builds appear to hang. Check `Get-Process Excel` before assuming VBA is broken.
- Preserve exported UserForm `.frm` and `.frx` pairs. If the `.frx` is missing, the import script must regenerate controls through VBIDE automation, which is slower and more fragile.
- Prefer importing a real `.frm/.frx` pair over rebuilding UserForm controls by script on every build.
- Set Excel automation knobs before opening workbooks in build/import scripts: `DisplayAlerts = False`, `EnableEvents = False`, `ScreenUpdating = False`, `AskToUpdateLinks = False`, and `AutomationSecurity = 3` to keep add-in events from firing while editing the VBA project.
- Do not use `AutomationSecurity = 3` in smoke tests that need to run add-in macros. Use `AutomationSecurity = 1` there.
- Use `build_xlam.ps1` for source-to-dist builds, `smoke_test.ps1 -SkipBuild` when the build is already current, and `publish_addin.ps1` for the full build/test/backup/install flow.
- If a macro appears unavailable, first check for a runtime VBA error shown in the VBE. Excel may report macro availability errors when a called macro immediately hits a runtime error.
- Keep local workbook mappings outside client workbooks. Current mapping state lives under `%APPDATA%\FAS_Toolkit\Mappings`.
- For performance, command jumps like `p`, `p1`, and `p2` should resolve assigned slots directly before doing broader fuzzy sheet scans.
- Keep `Ctrl+Shift+J` as the main command entry point. Add new command tokens in `modCommandRegistry.bas` first, then keep older public macros as wrappers where useful.
- Common command-palette actions should avoid workbook-wide scans. Slot jumps, paste tools, precedent `pre`, and `back` should feel effectively instant on normal FDD workbooks.
- Precedent tracing v1 intentionally supports only normal same-sheet and same-workbook A1 references. External links, names, structured table refs, `INDIRECT`, and dynamic references should be skipped or reported without crashing.
- TB/account search commands such as `tb cash` and `acct cash` should search only the resolved TB/BB sheet's used range, never every sheet in the workbook.
- Common-sheet lookup commands such as `gl vendor` and `map revenue` should search only the resolved sheet for that command.
- Bookmark commands should store local state under `%APPDATA%\FAS_Toolkit\Bookmarks`, never inside client workbooks.
- Smart formula copy is intentionally conservative. It supports normal same-sheet and same-workbook sheet A1 references first: references fully inside the copied source range move with the destination; other same-sheet/cross-sheet references are preserved as absolute references. Refuse unsupported formulas rather than silently rewriting external links, structured table refs, `INDIRECT`, `OFFSET`, spilled refs, or 3D refs.
- Mapping-management commands may read/write `%APPDATA%\FAS_Toolkit\Mappings`, but must not write hidden config into client workbooks.
- When parsing command payloads such as `tb 100-200`, preserve the user's raw search text after the command prefix. Do not run the payload through sheet-name normalization, because account codes and punctuation may matter.
- Local state files use pipe-delimited text. Encode/decode user-controlled fields before writing them so bookmark names or sheet names containing `|` do not corrupt the file.
- Formula safety checks should ignore text inside quoted formula strings before rejecting brackets, hash characters, or volatile/dynamic references.
- Keep normal workflow feedback in `Application.StatusBar` through `FAS_SetStatus`. Avoid modal popups except for command input and genuinely risky confirmations.
- Command-mode tools should prefer first-match deterministic jumps over picker dialogs. Add indexed variants such as `pre2` when a user needs a non-first target.
- Drilldown output should use one reusable `xlFreeFloating` worksheet shape per sheet. Do not insert rows, alter filters, or write client data.
- Selected-object placement fixes (`fixobj`/`float`) should be confirmation-free and scoped only to the selected shapes/images/charts. Keep whole-workbook shape fixing behind confirmation.
- When debugging Excel COM, isolate with one small probe macro or one macro call at a time. Avoid repeatedly running the full smoke test until the failing step is known.
- Project ledger/TB/mapping search state lives under `%APPDATA%\FAS_Toolkit\Projects`; never write project config or search indexes into client workbooks.
- Project setup should stay friendly: `tools\fas_project.ps1` with no arguments should remain an interactive menu, while argument-based actions remain available for scripted/bulk use.
- Indexed search commands such as `dk 489006`, `ledger cash`, `dktb cash`, and `dkmap revenue` should query the local project index and render lightweight results. They should not filter or scan every Excel worksheet.
- Search-index tooling should tolerate raw files moving after an index is built. Queries use the local index; rebuilds require a current source path or re-adding the source.
- For raw `.xlsx` ingestion, prefer conservative read-only Excel COM extraction. If protected-view/lock issues block opening, support CSV sources as the fallback instead of trying invasive workbook changes.
