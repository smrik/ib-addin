[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [switch]$SaveAfterSync,

    [int]$PollMilliseconds = 500
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$sourceRoot = Join-Path $repoRoot "vba-files"
$syncScriptPath = Join-Path $PSScriptRoot "sync_vba_file.ps1"
$supportedExtensions = @(".bas", ".cls", ".frm", ".frx")

function Get-TrackedFiles {
    Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
        Where-Object { $supportedExtensions -contains $_.Extension.ToLowerInvariant() } |
        Sort-Object FullName
}

function Get-FileSnapshot {
    $snapshot = @{}

    foreach ($file in Get-TrackedFiles) {
        $snapshot[$file.FullName] = $file.LastWriteTimeUtc.Ticks
    }

    return $snapshot
}

function Get-ChangedSourceFiles {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$PreviousSnapshot,

        [Parameter(Mandatory = $true)]
        [hashtable]$CurrentSnapshot
    )

    $changedFiles = New-Object System.Collections.Generic.List[string]

    foreach ($path in $CurrentSnapshot.Keys) {
        if (-not $PreviousSnapshot.ContainsKey($path) -or $PreviousSnapshot[$path] -ne $CurrentSnapshot[$path]) {
            [void]$changedFiles.Add($path)
        }
    }

    return @($changedFiles)
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook path not found: $WorkbookPath"
}

if (-not (Test-Path -LiteralPath $syncScriptPath)) {
    throw "Sync script not found: $syncScriptPath"
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$snapshot = Get-FileSnapshot

Write-Host "Watching $sourceRoot for VBA changes"
Write-Host "Target workbook/add-in: $resolvedWorkbookPath"
Write-Host "Press Ctrl+C to stop."

while ($true) {
    Start-Sleep -Milliseconds $PollMilliseconds

    $currentSnapshot = Get-FileSnapshot
    $changedPaths = @(Get-ChangedSourceFiles -PreviousSnapshot $snapshot -CurrentSnapshot $currentSnapshot)
    $snapshot = $currentSnapshot

    if ($changedPaths.Length -eq 0) {
        continue
    }

    $syncTargets = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($changedPath in $changedPaths) {
        $extension = [System.IO.Path]::GetExtension($changedPath)
        if ($extension -ieq ".frx") {
            $frmPath = [System.IO.Path]::ChangeExtension($changedPath, ".frm")
            if (Test-Path -LiteralPath $frmPath) {
                [void]$syncTargets.Add($frmPath)
            }

            continue
        }

        [void]$syncTargets.Add($changedPath)
    }

    foreach ($target in @($syncTargets | Sort-Object)) {
        try {
            $syncParams = @{
                WorkbookPath = $resolvedWorkbookPath
                SourceFile = $target
            }

            if ($SaveAfterSync) {
                $syncParams.SaveAfterSync = $true
            }

            & $syncScriptPath @syncParams
        }
        catch {
            Write-Warning "Live sync failed for $target. $($_.Exception.Message)"
        }
    }
}
