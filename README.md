# FAS Add-In development workspace

This is a Windows development workspace for the FAS Excel and PowerPoint productivity tooling. It is intended to be cloned to another development computer and opened in VS Code.

## Projects

- `fas-excel-toolkit` — Excel `.xlam` source under `src/vba`, with build, live-sync, and Excel COM smoke scripts.
- `Instrumenta` — PowerPoint toolbar source and editable PowerPoint host under `src`.
- `patriks-tab` — XVBA-based Excel development source under `vba-files`.
- `Instrumenta-Keys` — optional keyboard companion source for Windows and macOS.

Compiled add-ins, generated workbooks, vendored dependencies, backups, logs, and local Office state are intentionally ignored by Git. The editable Office host files required by the PowerPoint projects remain included.

## Requirements

- Windows with desktop Microsoft Excel and PowerPoint.
- VS Code with the recommended extensions from `.vscode/extensions.json`.
- Windows PowerShell 5.1 or PowerShell 7.
- Node.js/npm for XVBA dependency installation.
- Python 3 for the `patriks-tab` source validator.
- Excel Trust Center access to the VBA project object model for build/import workflows.

## First setup on another computer

1. Install the requirements above.
2. Clone this repository.
3. Open the repository folder in VS Code.
4. Run:

   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File .\tools\setup-dev.ps1 -InstallExtensions -InstallNodeDependencies
   ```

5. In VS Code, run the task **Run source checks**.

The setup script only installs VS Code extensions and project dependencies when those switches are supplied. It does not change Office macro security or silently terminate Excel/PowerPoint.

## Publish the development workspace to GitHub

The outer workspace is initialized on branch `main`, with the source snapshot and development configuration staged locally. After creating an empty GitHub repository, run:

```powershell
git config user.name "Your Name"
git config user.email "you@example.com"
git remote add origin https://github.com/<account>/<repository>.git
git commit -m "Prepare development workspace"
git push -u origin main
```

`Instrumenta` and `Instrumenta-Keys` originally contained nested Git repositories. Their source is included in this outer workspace snapshot so a fresh clone opens as one working environment; their nested histories remain local to this checkout.

## Useful VS Code tasks

- **Run source checks** — static checks for Excel, PowerPoint, and XVBA source.
- **Excel: Build add-in** — builds `fas-excel-toolkit/dist/FAS_Toolkit.xlam` through Excel COM.
- **Excel: Start test instance** — opens an isolated Excel test instance.
- **Excel: Watch VBA live sync** — syncs edited source modules into an open development add-in.
- **Excel: Smoke test** — runs the Excel COM smoke suite.
- **PowerPoint: Build dev add-in** — builds the development `.ppam`.
- **PowerPoint: Source checks** — runs the PowerPoint source-only contracts.

## Office host notes

These projects are Office VBA add-ins, not standalone applications. The build and smoke tasks require the corresponding desktop Office application and usually require that application to be closed before a build. Do not commit the generated `.xlam`/`.ppam` files; rebuild them locally.

For the Excel command palette and live-sync workflow, see [`fas-excel-toolkit/README.md`](fas-excel-toolkit/README.md). For the PowerPoint build flow, see [`Instrumenta/README.md`](Instrumenta/README.md).
