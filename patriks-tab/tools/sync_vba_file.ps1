[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$SourceFile,

    [switch]$SaveAfterSync
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$sourceRoot = Join-Path $repoRoot "vba-files"
$modulesPath = Join-Path $sourceRoot "Module"
$classesPath = Join-Path $sourceRoot "Class"
$formsPath = Join-Path $sourceRoot "Form"
$documentsPath = Join-Path $sourceRoot "Class"
# Document modules share the Class folder with real classes in this layout,
# so route them by name before the generic class branch.
$documentModuleNames = @("ThisWorkbook", "Sheet1")

function Test-PathWithinRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $normalizedPath = [System.IO.Path]::GetFullPath($Path)
    $normalizedRoot = [System.IO.Path]::GetFullPath($Root)

    if (-not $normalizedRoot.EndsWith([System.IO.Path]::DirectorySeparatorChar)) {
        $normalizedRoot += [System.IO.Path]::DirectorySeparatorChar
    }

    return $normalizedPath.StartsWith($normalizedRoot, [System.StringComparison]::OrdinalIgnoreCase)
}

function Get-RunningExcelApplication {
    try {
        return [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        return $null
    }
}

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

function Set-CodeModuleContent {
    param(
        [Parameter(Mandatory = $true)]
        $CodeModule,

        [Parameter(Mandatory = $true)]
        [string]$CodeText
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }

    if ([string]::IsNullOrWhiteSpace($CodeText)) {
        return
    }

    $CodeModule.InsertLines(1, $CodeText)
}

function Remove-ComponentIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$ComponentName
    )

    try {
        $existing = $VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        $existing = $null
    }

    if ($null -eq $existing) {
        return
    }

    if ($existing.Type -eq 100) {
        throw "Refusing to remove document module '$ComponentName'."
    }

    $VBProject.VBComponents.Remove($existing)
}

function Import-FileComponent {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    Remove-ComponentIfPresent -VBProject $VBProject -ComponentName $File.BaseName
    [void]$VBProject.VBComponents.Import($File.FullName)
}

function Get-DocumentModuleCode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    # Document modules are exported with a .cls header (VERSION / BEGIN..END /
    # Attribute lines). Only the code below that header may be pushed into an
    # existing document component - inserting the header makes the whole project
    # fail to compile with "Expected: end of statement".
    $lines = Get-Content -LiteralPath $Path
    $startIndex = 0
    $inBeginBlock = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i].Trim()

        if ($inBeginBlock) {
            $startIndex = $i + 1
            if ($line -eq "End" -or $line -eq "END") { $inBeginBlock = $false }
            continue
        }
        if ($line -match "^VERSION\s") { $startIndex = $i + 1; continue }
        if ($line -eq "BEGIN") { $inBeginBlock = $true; $startIndex = $i + 1; continue }
        if ($line -match "^Attribute\s") { $startIndex = $i + 1; continue }
        if ($line -eq "") { continue }
        break
    }

    if ($startIndex -ge $lines.Count) { return "" }
    return ($lines[$startIndex..($lines.Count - 1)] -join [Environment]::NewLine)
}

function Set-DocumentModuleCode {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $component = $VBProject.VBComponents.Item($File.BaseName)
    $codeText = Get-DocumentModuleCode -Path $File.FullName
    Set-CodeModuleContent -CodeModule $component.CodeModule -CodeText $codeText
}

function Resolve-SourceFileForSync {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path

    if (-not (Test-PathWithinRoot -Path $resolvedPath -Root $sourceRoot)) {
        throw "Source file must be inside $sourceRoot"
    }

    $extension = [System.IO.Path]::GetExtension($resolvedPath)
    if ($extension -ieq ".frx") {
        $formPath = [System.IO.Path]::ChangeExtension($resolvedPath, ".frm")
        if (-not (Test-Path -LiteralPath $formPath)) {
            throw "Changed .frx file has no matching .frm file: $resolvedPath"
        }

        return (Resolve-Path -LiteralPath $formPath).Path
    }

    return $resolvedPath
}

function Sync-SourceFile {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$ResolvedSourceFile
    )

    $file = Get-Item -LiteralPath $ResolvedSourceFile

    if ((Test-PathWithinRoot -Path $ResolvedSourceFile -Root $modulesPath) -and $file.Extension -ieq ".bas") {
        Import-FileComponent -VBProject $VBProject -File $file
        return "module $($file.Name)"
    }

    if ((Test-PathWithinRoot -Path $ResolvedSourceFile -Root $documentsPath) -and $file.Extension -ieq ".cls" -and
        $documentModuleNames -contains $file.BaseName) {
        Set-DocumentModuleCode -VBProject $VBProject -File $file
        return "document module $($file.Name)"
    }

    if ((Test-PathWithinRoot -Path $ResolvedSourceFile -Root $classesPath) -and $file.Extension -ieq ".cls") {
        Import-FileComponent -VBProject $VBProject -File $file
        return "class $($file.Name)"
    }

    if ((Test-PathWithinRoot -Path $ResolvedSourceFile -Root $formsPath) -and $file.Extension -ieq ".frm") {
        $frxPath = [System.IO.Path]::ChangeExtension($ResolvedSourceFile, ".frx")
        if (-not (Test-Path -LiteralPath $frxPath)) {
            throw "Form sync requires a matching .frx file: $ResolvedSourceFile"
        }

        Import-FileComponent -VBProject $VBProject -File $file
        return "form $($file.Name)"
    }

    if ((Test-PathWithinRoot -Path $ResolvedSourceFile -Root $documentsPath) -and $file.Extension -ieq ".cls") {
        Set-DocumentModuleCode -VBProject $VBProject -File $file
        return "document module $($file.Name)"
    }

    throw "Unsupported source file for live sync: $ResolvedSourceFile"
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook path not found: $WorkbookPath"
}

if (-not (Test-Path -LiteralPath $SourceFile)) {
    throw "Source file not found: $SourceFile"
}

$resolvedWorkbookPath = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedSourceFile = Resolve-SourceFileForSync -Path $SourceFile

$excel = $null
$workbook = $null
$createdExcel = $false
$openedWorkbook = $false

try {
    $workbook = Get-WorkbookObjectByPath -ResolvedWorkbookPath $resolvedWorkbookPath
    if ($null -ne $workbook) {
        $excel = $workbook.Application
    }
    else {
        $excel = Get-RunningExcelApplication
        if ($null -eq $excel) {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $true
            $createdExcel = $true
        }

        $workbook = Get-OpenWorkbookByPath -Excel $excel -ResolvedWorkbookPath $resolvedWorkbookPath
        if ($null -eq $workbook) {
            $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, 0, $false)
            $openedWorkbook = $true
        }
    }

    $previousDisplayAlerts = $excel.DisplayAlerts
    $previousEnableEvents = $excel.EnableEvents
    $previousScreenUpdating = $excel.ScreenUpdating
    $previousAskToUpdateLinks = $excel.AskToUpdateLinks

    try {
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        $excel.ScreenUpdating = $false
        $excel.AskToUpdateLinks = $false

        $syncLabel = Sync-SourceFile -VBProject $workbook.VBProject -ResolvedSourceFile $resolvedSourceFile

        if ($SaveAfterSync) {
            $workbook.Save()
        }
    }
    finally {
        $excel.DisplayAlerts = $previousDisplayAlerts
        $excel.EnableEvents = $previousEnableEvents
        $excel.ScreenUpdating = $previousScreenUpdating
        $excel.AskToUpdateLinks = $previousAskToUpdateLinks
    }

    $saveText = if ($SaveAfterSync) { " and saved workbook" } else { "" }
    Write-Host "Synced $syncLabel -> $resolvedWorkbookPath$saveText"
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
