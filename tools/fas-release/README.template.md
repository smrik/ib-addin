# FAS Add-ins {{VERSION}}

## Included add-ins

{{INCLUDED_LIST}}

## Included support files

{{SUPPORT_LIST}}

## Install

1. Close Excel and PowerPoint completely.
2. Extract this ZIP to a local folder.
3. Double-click `Install-FASAddins.cmd`.
4. Reopen Excel or PowerPoint.

The installer attempts to register the included add-ins in Office automatically. If Office is not installed, it installs the files and reports that registration was skipped.

If the one-click launcher is unavailable, open PowerShell in the extracted folder and run:

```powershell
.\Install-FASAddins.ps1
```

## Uninstall

To remove the installed FAS add-ins, double-click `Uninstall-FASAddins.cmd`, or run this in PowerShell:

```powershell
.\Uninstall-FASAddins.ps1
```

## Restore previous local version

To uninstall the current package and restore the latest local backup, run:

```powershell
.\Uninstall-FASAddins.ps1 -RestoreLatestBackup
```

## Troubleshooting

### Installer says an add-in is in use

Close Excel and PowerPoint completely, then run the installer again.

If the error continues, check Task Manager and close any remaining `EXCEL.EXE` or `POWERPNT.EXE` processes.

### PowerShell blocks the script

If Windows blocks the script because it came from a downloaded ZIP, open PowerShell in the extracted folder and run:

```powershell
Unblock-File .\Install-FASAddins.ps1
Unblock-File .\Uninstall-FASAddins.ps1
.\Install-FASAddins.ps1
```

### Installed locations

The installer copies add-ins to the standard local Office add-ins folder:

```text
%APPDATA%\Microsoft\AddIns
```

Backups and local install state are stored under:

```text
%APPDATA%\FAS_Addins
```

## Package details

Version: `{{VERSION}}`

This package installs only the add-ins listed above. It backs up existing local add-ins before replacing them and records installation state locally so uninstall is limited to files owned by this package.

## Requirements and limitations

- Windows with desktop Excel and/or PowerPoint installed.
- VBA macros must be allowed by the Office Trust Center. The installer does not weaken macro security.
- Excel project search also needs `duckdb.exe` for the fast index path. The search tools work without it for staging, but the DuckDB index cannot be built until DuckDB is installed separately.
- A container is not useful for these VBA add-ins because they must run inside the host Office applications on the recipient computer.
