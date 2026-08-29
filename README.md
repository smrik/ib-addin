# FAS Add-In development workspace

Windows development workspace for the FAS Excel and PowerPoint productivity add-ins. Clone it to a development computer and open the folder in VS Code.

## Projects

- `patriks-tab` — the Excel add-in. XVBA-based source under `vba-files`, Ribbon under `ribbons`, with build, live-sync, smoke and publish scripts under `tools`.
- `Instrumenta` — the PowerPoint toolbar. Source under `src`, built through `scripts/build-dev.ps1` / `build-prod.ps1`.
- `Instrumenta-Keys` — optional keyboard companion for Windows and macOS.

The Excel add-in previously lived in two forked trees (`fas-excel-toolkit` and `patriks-tab`). They were merged into `patriks-tab` in August 2026; see `git log` for the merge commits.

Compiled add-ins, generated workbooks, vendored dependencies, backups, logs, and local Office state are intentionally ignored by Git. The editable PowerPoint host file required by `Instrumenta` remains included.

## Requirements

- Windows with desktop Microsoft Excel and PowerPoint.
- VS Code with the recommended extensions from `.vscode/extensions.json`.
- Windows PowerShell 5.1 or PowerShell 7.
- Node.js/npm for XVBA dependency installation.
- Python 3 for the Excel source validator.
- Excel Trust Center: **Trust access to the VBA project object model** must be enabled, or every build/import script will hang rather than fail.

## First setup on another computer

1. Install the requirements above.
2. Clone this repository and open the folder in VS Code.
3. Run:

   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\tools\setup-dev.ps1 -InstallExtensions -InstallNodeDependencies
   ```

4. In VS Code, run the task **Run source checks**.

The setup script only installs VS Code extensions and project dependencies when those switches are supplied. It does not change Office macro security or silently terminate Excel/PowerPoint.

## Development loops

**Excel** is developed either through the XVBA VS Code extension or through the PowerShell scripts; both produce the same artifact, `patriks-tab/patriks-tab.xlam`. XVBA gives the fast edit loop; the scripts are the reliable build/test/publish path.

**PowerPoint** cannot use XVBA — it is an Excel-only extension. That side edits `src/Modules` etc. and imports them with `scripts/build-dev.ps1 -SyncVba`.

## Useful VS Code tasks

- **Run source checks** — static checks for Excel and PowerPoint source. No Office required; this is also what CI runs.
- **Excel: Build add-in** — builds `patriks-tab/patriks-tab.xlam` through Excel COM.
- **Excel: Start test instance** — opens an isolated Excel test instance.
- **Excel: Watch VBA live sync** — syncs edited source modules into an open development add-in.
- **Excel: Smoke test** — runs the Excel COM smoke suite.
- **PowerPoint: Build dev add-in** — builds the development `.ppam`.
- **PowerPoint: Source checks** — runs the PowerPoint source-only contracts.

## Versioning and releases

Git is the archive; do not hand-roll backup copies of Office binaries. Tag a release and build the shipped ZIP:

```powershell
git tag v1.2.3
powershell.exe -ExecutionPolicy Bypass -File .\tools\package-fas-addins.ps1 -Version 1.2.3
```

Packages land in `releases/` (ignored by Git). Installer and uninstaller templates live in `tools/fas-release/`.

## Office host notes

These projects are Office VBA add-ins, not standalone applications. The build and smoke tasks require the corresponding desktop Office application, and usually require it to be closed first. Do not commit the generated `.xlam`/`.ppam` files; rebuild them locally.

For architecture and the full command map, see [`CLAUDE.md`](CLAUDE.md). For the PowerPoint feature list, see [`Instrumenta/README.md`](Instrumenta/README.md) — but note that its build instructions describe the upstream project, not this workspace.
