[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Excel must allow "Trust access to the VBA project object model" for VBProject automation to work.

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$sourceRoot = Join-Path $repoRoot "src\vba"
$modulesPath = Join-Path $sourceRoot "modules"
$classesPath = Join-Path $sourceRoot "classes"
$formsPath = Join-Path $sourceRoot "forms"
$documentsPath = Join-Path $sourceRoot "documents"

function Assert-CanWriteFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ((Test-Path -LiteralPath $Path) -and -not $Force) {
        throw "Refusing to overwrite existing file without -Force: $Path"
    }
}

function Export-ComponentFile {
    param(
        [Parameter(Mandatory = $true)]
        $Component,

        [Parameter(Mandatory = $true)]
        [string]$DestinationFolder
    )

    $extension = switch ($Component.Type) {
        1 { ".bas" }
        2 { ".cls" }
        3 { ".frm" }
        default { throw "Unsupported component type '$($Component.Type)' for export." }
    }

    $destinationPath = Join-Path $DestinationFolder ($Component.Name + $extension)
    Assert-CanWriteFile -Path $destinationPath

    if ($extension -eq ".frm") {
        $frxPath = [System.IO.Path]::ChangeExtension($destinationPath, ".frx")
        if ((Test-Path -LiteralPath $frxPath) -and -not $Force) {
            throw "Refusing to overwrite existing file without -Force: $frxPath"
        }
    }

    $Component.Export($destinationPath)
}

function Export-DocumentModuleCode {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    Assert-CanWriteFile -Path $DestinationPath
    $component = $VBProject.VBComponents.Item("ThisWorkbook")
    $codeModule = $component.CodeModule
    $codeText = if ($codeModule.CountOfLines -gt 0) {
        $codeModule.Lines(1, $codeModule.CountOfLines)
    } else {
        "Option Explicit"
    }

    Set-Content -LiteralPath $DestinationPath -Value $codeText -Encoding UTF8
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook path not found: $WorkbookPath"
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, $null, $true)
    $vbProject = $workbook.VBProject

    foreach ($component in $vbProject.VBComponents) {
        switch ($component.Type) {
            1 { Export-ComponentFile -Component $component -DestinationFolder $modulesPath }
            2 { Export-ComponentFile -Component $component -DestinationFolder $classesPath }
            3 { Export-ComponentFile -Component $component -DestinationFolder $formsPath }
            100 { }
            default { }
        }
    }

    Export-DocumentModuleCode -VBProject $vbProject -DestinationPath (Join-Path $documentsPath "ThisWorkbook.cls")
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($false)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
