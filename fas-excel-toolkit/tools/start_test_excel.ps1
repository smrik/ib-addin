[CmdletBinding()]
param(
    [string]$AddInPath = ".\dist\FAS_Toolkit.xlam",

    [string]$Caption = "TEST INSTANCE",

    [switch]$CreateBlankWorkbook = $true,

    [switch]$RegisterShortcuts = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepoRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return $Path
    }

    $repoRoot = Split-Path -Path $PSScriptRoot -Parent
    return (Join-Path $repoRoot $Path)
}

function Close-OtherAddInWorkbooks {
    param(
        [Parameter(Mandatory = $true)]
        $Excel,

        [Parameter(Mandatory = $true)]
        [string]$KeepAddInPath
    )

    $workbooksToClose = New-Object System.Collections.Generic.List[object]

    foreach ($candidate in $Excel.Workbooks) {
        if (-not $candidate.IsAddin) {
            continue
        }

        if ($candidate.FullName -ieq $KeepAddInPath) {
            continue
        }

        [void]$workbooksToClose.Add($candidate)
    }

    foreach ($candidate in $workbooksToClose) {
        Write-Host "Closing preloaded add-in in test instance: $($candidate.FullName)"
        $candidate.Close($false)
    }
}

function Get-QualifiedMacroName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkbookName,

        [Parameter(Mandatory = $true)]
        [string]$MacroName
    )

    return "'" + $WorkbookName + "'!" + $MacroName
}

if ([string]::IsNullOrWhiteSpace($Caption)) {
    throw "Caption cannot be blank."
}

$resolvedAddInPath = (Resolve-Path -LiteralPath (Get-RepoRelativePath -Path $AddInPath)).Path
if (-not (Test-Path -LiteralPath $resolvedAddInPath)) {
    throw "Add-in path not found: $resolvedAddInPath"
}

$excel = $null
$devAddIn = $null
$testWorkbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.UserControl = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.ScreenUpdating = $true
    $excel.Caption = $Caption

    # A fresh automation-created Excel instance still inherits any profile-level
    # add-ins that Excel auto-loads. Close them in this session so the test
    # window uses only the requested dev add-in.
    Close-OtherAddInWorkbooks -Excel $excel -KeepAddInPath $resolvedAddInPath

    $devAddIn = $excel.Workbooks.Open($resolvedAddInPath, 0, $false)

    if ($RegisterShortcuts) {
        try {
            [void]$excel.Run((Get-QualifiedMacroName -WorkbookName $devAddIn.Name -MacroName "FAS_RegisterShortcuts"))
        }
        catch {
            Write-Warning "Opened the dev add-in, but could not run FAS_RegisterShortcuts. $($_.Exception.Message)"
        }
    }

    if ($CreateBlankWorkbook) {
        $testWorkbook = $excel.Workbooks.Add()
        $testWorkbook.Activate() | Out-Null
    }

    Write-Host "Started isolated Excel test instance."
    Write-Host "Caption: $Caption"
    Write-Host "Loaded dev add-in: $resolvedAddInPath"
    if ($CreateBlankWorkbook) {
        Write-Host "Created a blank workbook for testing."
    }
}
finally {
    if ($null -ne $testWorkbook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($testWorkbook)
    }

    if ($null -ne $devAddIn) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($devAddIn)
    }

    if ($null -ne $excel) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
