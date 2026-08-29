# FAS VBA release packaging design

## Goal

Create a repeatable, low-friction way to distribute the existing Excel and PowerPoint VBA add-ins to a small local team through SharePoint. The workflow must preserve the current VBA implementations, produce versioned release packages, make installation predictable, and retain a recoverable previous installation.

## Scope

The first version packages already-built artifacts:

- Excel: `fas-excel-toolkit/dist/FAS_Toolkit.xlam`
- PowerPoint: `Instrumenta/bin/InstrumentaPowerpointToolbar.ppam`

It does not rewrite VBA, split compiled add-ins into separate feature binaries, introduce an automatic background updater, or modify the existing build/test behavior.

## Release layout

Each release is created under a versioned directory and ZIP archive:

```text
FAS-Addins-1.0.0/
  Excel/FAS_Toolkit.xlam
  PowerPoint/InstrumentaPowerpointToolbar.ppam
  Install-FASAddins.ps1
  Uninstall-FASAddins.ps1
  README.md
  release.json
  SHA256SUMS.txt
```

The packager supports combined, Excel-only, and PowerPoint-only releases. The release manifest records the version, UTC creation time, included artifacts, source paths, and hashes. The ZIP is the SharePoint handoff artifact; users download it locally before installing.

## Installation and rollback

The installer copies files to the current local destinations without changing those destinations:

- Excel: `%APPDATA%\\Microsoft\\AddIns\\FAS_Toolkit.xlam`
- PowerPoint: `%APPDATA%\\Microsoft\\AddIns\\InstrumentaPowerpointToolbar.ppam`

Before replacement, it copies any existing add-in into a versioned backup directory under `%APPDATA%\\FAS_Addins\\backups`. It writes a local installation manifest under `%APPDATA%\\FAS_Addins`. Installation refuses to overwrite a running/locked add-in and provides a clear instruction to close Office first. The uninstaller removes only files owned by this package and can restore the most recent backup.

Rollback is performed by running the installer from an older release ZIP. The package does not silently change a user's installed version.

## Feature-by-feature distribution

The initial package always distributes the tested add-in binary as a whole. Existing source modules remain the feature boundary for development, but no attempt is made to remove modules from an already-built `.xlam` or `.ppam`. Future feature profiles may be added after the desired optional features and their dependencies are known.

## Validation

Packaging validation will check that required artifacts exist, the release version is valid, every manifest artifact has a matching checksum, the ZIP contains the expected layout, and the installer scripts are present. Existing Excel and PowerPoint smoke tests remain the authoritative functional checks and are not replaced by packaging validation.

## SharePoint workflow

The owner runs the existing build/smoke flows, runs the packager, reviews the generated manifest and checksums, and uploads the ZIP to a private SharePoint release folder. Teammates download the ZIP, extract it locally, and run the installer. A future stable-channel manifest or update shortcut can be added without changing the release layout.
