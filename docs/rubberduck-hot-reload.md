# Rubberduck VBA hot reload

This is the shared development loop for the Excel and PowerPoint add-ins.

## Goal

Edit VBA source in VS Code while the development host remains open. Reload the changed components with one Rubberduck command. Normal VBA changes do not require a rebuild or host restart.

Rubberduck's `Update Components from Files...` replaces components with matching names, adds missing components, and leaves unrelated components alone. Do not use `Replace Contents from Files...`; that command removes non-document modules. See the [Rubberduck sync guidance](https://rubberduckvba.blog/2020/12/11/sync-project-with-files/).

## Host and source paths

| Host | Development container | VBA source |
| --- | --- | --- |
| Excel | `C:\Users\psmid\Documents\FAS Add-In\fas-excel-toolkit\dist\FAS_Toolkit.xlam` | `C:\Users\psmid\Documents\FAS Add-In\fas-excel-toolkit\src\vba` |
| PowerPoint | `C:\Users\psmid\Documents\FAS Add-In\Instrumenta\bin\InstrumentaPowerpointToolbarDev.ppam` | `C:\Users\psmid\Documents\FAS Add-In\Instrumenta\src` |

Source files are canonical. The `.xlam` and `.ppam` files are development containers, not source files and not production publication targets.

## One-time setup

1. Install Rubberduck for the current Windows user.
2. Enable `Trust access to the VBA project object model` in both Office hosts.
3. Open the development add-in and press `Alt+F11`.
4. In Rubberduck Code Explorer, export each project once to a temporary audit folder. Compare component names with the source tree. Do not export over the canonical source tree.
5. Set a Rubberduck shortcut for `Update Components from Files...`, preferably `Ctrl+Alt+R`.

### ARM64 setup on this machine

The stable Rubberduck 2.5.91 installer is x86/x64-only and rejects Windows ARM64. The verified payload is extracted to:

`%LOCALAPPDATA%\Rubberduck\2.5.9.6316`

Register it with:

```powershell
& "$env:USERPROFILE\Documents\FAS Add-In\tools\Register-RubberduckUserAddIn.ps1"
```

Then close and reopen Excel and PowerPoint once. This is only the initial add-in registration; it is not part of the daily reload loop.

## Daily loop

1. Open the development add-in.
2. Edit and save `.bas` or `.cls` files in VS Code.
3. In VBE, run:

   `Rubberduck Code Explorer -> Project -> Sync Project -> Update Components from Files...`

4. Select the changed source files. Rubberduck's dialog selects files, not a whole project folder; multi-select the changed files. Keep the same-name mapping intact.
5. Test immediately.

If changes span separate source folders, repeat the same sync command for the additional folder. No rebuild is needed.

## Component rules

- Keep most changing logic in ordinary `.bas` and `.cls` files.
- Keep `ThisWorkbook`, `ThisPresentation`, worksheet modules, and Ribbon callbacks small.
- Keep `.frm` and matching `.frx` files together.
- Update document modules in place. Never delete or recreate them during routine sync.
- Do not use `Replace Contents from Files...`.
- Do not export from VBE during normal work; it can overwrite source edits.
- Do not sync into production add-ins.

## Ribbon boundary

Ribbon callback VBA follows the normal hot-reload path. Ribbon XML, icons, and package structure are separate artifacts; changes to those still use the existing package/build workflow and may require a Ribbon reload.

## Collaboration

The source tree is shared. Both the user and Codex may edit it. Different files are safe to edit in parallel; edits to the same file must be sequential. Save all intended changes before syncing.

## Recovery

If sync fails or the host becomes unstable:

1. Revert or fix the source file.
2. Run `Update Components from Files...` again.
3. Compile the VBA project.
4. Restart only if the VBE cannot recover.

The old Excel watcher scripts remain available as legacy recovery code, but are not the default workflow.

## Smoke checklist

- Excel existing module update.
- PowerPoint existing module update.
- New module import.
- Class update.
- Paired form update.
- Document module update without deletion.
- Ribbon callback VBA update.
- Broken source rollback and resync.
- Production add-ins remain unchanged.
