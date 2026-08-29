[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$MacroName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-WorkbookObjectByPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResolvedWorkbookPath
    )

    try {
        return [Microsoft.VisualBasic.Interaction]::GetObject($ResolvedWorkbookPath, $null)
    }
    catch {
        return $null
    }
}

function Get-RunningExcelApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        return $null
    }
}

function Get-OpenWorkbookByPath {
    param(
        [Parameter(Mandatory = $true)]
        $Excel,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedWorkbookPath
    )

    foreach ($candidate in $Excel.Workbooks) {
        if ($candidate.FullName -ieq $ResolvedWorkbookPath) {
            return $candidate
        }
    }

    return $null
}

function Get-OpenWorkbookByName {
    param(
        [Parameter(Mandatory = $true)]
        $Excel,

        [Parameter(Mandatory = $true)]
        [string]$WorkbookName
    )

    foreach ($candidate in $Excel.Workbooks) {
        if ($candidate.Name -ieq $WorkbookName) {
            return $candidate
        }
    }

    return $null
}

function Get-OpenWorkbookSummary {
    param(
        [Parameter(Mandatory = $true)]
        $Excel
    )

    $lines = New-Object System.Collections.Generic.List[string]

    foreach ($candidate in $Excel.Workbooks) {
        $pathText = if ([string]::IsNullOrWhiteSpace($candidate.FullName)) {
            $candidate.Name
        } else {
            $candidate.FullName
        }

        [void]$lines.Add([string]$pathText)
    }

    if ($lines.Count -eq 0) {
        return "(none)"
    }

    return ($lines -join [Environment]::NewLine)
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook path not found: $WorkbookPath"
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$trimmedMacroName = $MacroName.Trim()
if ([string]::IsNullOrWhiteSpace($trimmedMacroName)) {
    throw "Macro name cannot be blank."
}

$excel = $null
$workbook = $null

try {
    $workbook = Get-WorkbookObjectByPath -ResolvedWorkbookPath $resolvedWorkbookPath
    if ($null -ne $workbook) {
        $excel = $workbook.Application
    }
    else {
        $excel = Get-RunningExcelApplication
        if ($null -eq $excel) {
            throw "Excel is not running. Open Excel and load the target workbook or add-in first."
        }

        $workbook = Get-OpenWorkbookByPath -Excel $excel -ResolvedWorkbookPath $resolvedWorkbookPath
        if ($null -eq $workbook) {
            $workbook = Get-OpenWorkbookByName -Excel $excel -WorkbookName ([System.IO.Path]::GetFileName($resolvedWorkbookPath))
        }

        if ($null -eq $workbook) {
            $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, 0, $false)
        }
    }

    $qualifiedMacroName = "'" + $workbook.Name + "'!" + $trimmedMacroName
    try {
        [void]$excel.Run($qualifiedMacroName)
    }
    catch {
        $openSummary = Get-OpenWorkbookSummary -Excel $excel
        throw "Could not run macro $qualifiedMacroName. Open workbooks/add-ins in Excel:`n$openSummary`n$($_.Exception.Message)"
    }
    Write-Host "Ran macro $qualifiedMacroName"
}
finally {
    if ($null -ne $workbook) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }

    if ($null -ne $excel) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
