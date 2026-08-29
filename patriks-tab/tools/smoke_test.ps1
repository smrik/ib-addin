[CmdletBinding()]
param(
    [switch]$SkipBuild,
    [switch]$ForceCloseExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$buildScript = Join-Path $PSScriptRoot "build_xlam.ps1"
$addInPath = Join-Path $repoRoot "patriks-tab.xlam"
$tempWorkbookPath = Join-Path ([System.IO.Path]::GetTempPath()) ("PatriksTab_Smoke_" + [System.Guid]::NewGuid().ToString("N") + ".xlsx")
$smokeLogPath = Join-Path ([System.IO.Path]::GetTempPath()) "PatriksTab_Smoke_Last.log"
$glSheetName = "Dnevnik knji" + [char]0x017E + "enja"
$plStandaloneName = "P&L Standalone"
$plConsolidatedName = "P&L Consolidated"

Remove-Item -LiteralPath $smokeLogPath -Force -ErrorAction SilentlyContinue

function Write-SmokeStep {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    Write-Host $Message
    Add-Content -LiteralPath $smokeLogPath -Value ("{0:u} {1}" -f (Get-Date), $Message)
}

function Stop-ExcelIfRequested {
    param(
        [switch]$ForceClose
    )

    $excelProcesses = @(Get-Process Excel -ErrorAction SilentlyContinue)
    if ($excelProcesses.Count -eq 0) {
        return
    }

    if (-not $ForceClose) {
        $processList = ($excelProcesses | ForEach-Object { "PID $($_.Id)" }) -join ", "
        throw "Excel is already running ($processList). Close Excel first, or rerun with -ForceCloseExcel."
    }

    $excelProcesses | Stop-Process -Force
}

function Assert-ActiveSheetName {
    param(
        [Parameter(Mandatory = $true)]
        $Workbook,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedSheetName,

        [Parameter(Mandatory = $true)]
        [string]$MacroName
    )

    $actualName = $Workbook.ActiveSheet.Name
    if ($actualName -ne $ExpectedSheetName) {
        throw "$MacroName activated '$actualName' instead of '$ExpectedSheetName'."
    }
}

function Assert-ActiveCellAddress {
    param(
        [Parameter(Mandatory = $true)]
        $Excel,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedAddress,

        [Parameter(Mandatory = $true)]
        [string]$CheckName
    )

    $actualAddress = $Excel.ActiveCell.Address($false, $false)
    if ($actualAddress -ne $ExpectedAddress) {
        throw "$CheckName selected '$actualAddress' instead of '$ExpectedAddress'."
    }
}

function Assert-ShapePlacement {
    param(
        [Parameter(Mandatory = $true)]
        $Shape,

        [Parameter(Mandatory = $true)]
        [int]$ExpectedPlacement,

        [Parameter(Mandatory = $true)]
        [string]$CheckName
    )

    if ([int]$Shape.Placement -ne $ExpectedPlacement) {
        throw "$CheckName had placement '$($Shape.Placement)' instead of '$ExpectedPlacement'."
    }
}

Stop-ExcelIfRequested -ForceClose:$ForceCloseExcel

if (-not $SkipBuild) {
    $buildParams = @{}
    if ($ForceCloseExcel) {
        $buildParams.ForceCloseExcel = $true
    }
    & $buildScript @buildParams
}

$excel = $null
$testWorkbook = $null
$addInWorkbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 1

    $testWorkbook = $excel.Workbooks.Add()
    $testWorkbook.Worksheets.Item(1).Name = $plConsolidatedName
    $sheet = $testWorkbook.Worksheets.Add()
    $sheet.Name = $plStandaloneName
    $sheet = $testWorkbook.Worksheets.Add()
    $sheet.Name = "BB 2024"
    $sheet = $testWorkbook.Worksheets.Add()
    $sheet.Name = $glSheetName
    $sheet = $testWorkbook.Worksheets.Add()
    $sheet.Name = "Klasifikacija"
    $sheet = $testWorkbook.Worksheets.Add()
    $sheet.Name = "Random Sheet"
    $testWorkbook.SaveAs($tempWorkbookPath, 51)

    $addInWorkbook = $excel.Workbooks.Open($addInPath)
    $testWorkbook.Activate()

    Write-SmokeStep "Smoke: saving workbook mappings"
    $excel.Run("'patriks-tab.xlam'!FAS_SaveActiveWorkbookAssignment", "PL", 1, $plStandaloneName) | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_SaveActiveWorkbookAssignment", "PL", 2, $plConsolidatedName) | Out-Null

    Write-SmokeStep "Smoke: legacy navigation macros"
    $testWorkbook.Activate()
    $excel.Run("'patriks-tab.xlam'!FAS_GoToPL")
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plStandaloneName -MacroName "FAS_GoToPL"

    $testWorkbook.Activate()
    $excel.Run("'patriks-tab.xlam'!FAS_GoToTB")
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "BB 2024" -MacroName "FAS_GoToTB"

    $testWorkbook.Activate()
    $excel.Run("'patriks-tab.xlam'!FAS_GoToGL")
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $glSheetName -MacroName "FAS_GoToGL"

    $testWorkbook.Worksheets.Item(1).Activate() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_GoToPL")
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plStandaloneName -MacroName "FAS_GoToPL (manual slot)"

    $testWorkbook.Worksheets.Item(1).Activate() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunJumpCommand", "p2") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plConsolidatedName -MacroName "FAS_RunJumpCommand p2"

    Write-SmokeStep "Smoke: command palette navigation"
    $testWorkbook.Worksheets.Item(1).Activate() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "p1") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plStandaloneName -MacroName "FAS_RunCommand p1"

    $testWorkbook.Worksheets.Item(1).Activate() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "p2") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plConsolidatedName -MacroName "FAS_RunCommand p2"

    $testWorkbook.Worksheets.Item(1).Activate() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "Random") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "Random Sheet" -MacroName "FAS_RunCommand fallback search"

    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "rep")) {
        throw "FAS_IsKnownCommand did not recognize rep."
    }

    Write-SmokeStep "Smoke: paste command tokens"
    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "val")) {
        throw "FAS_IsKnownCommand did not recognize val."
    }
    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "txt")) {
        throw "FAS_IsKnownCommand did not recognize txt."
    }

    Write-SmokeStep "Smoke: command catalog"
    $catalogCount = [int]$excel.Run("'patriks-tab.xlam'!FAS_GetCommandCatalogCount")
    if ($catalogCount -le 12) {
        throw "Expected the complete command catalog to contain more than 12 entries; got $catalogCount."
    }
    foreach ($catalogQuery in @("chart", "bookmark", "precedent", "layout", "mapping")) {
        $catalogMatchCount = [int]$excel.Run("'patriks-tab.xlam'!FAS_CountCommandCatalogMatches", $catalogQuery)
        if ($catalogMatchCount -lt 1) {
            throw "Expected at least one command catalog match for '$catalogQuery'."
        }
    }

    $expectedFirstCommands = [ordered]@{
        "p" = "p"
        "b" = "bs"
        "t" = "t"
        "g" = "g"
        "m" = "m"
    }
    foreach ($catalogQuery in $expectedFirstCommands.Keys) {
        $expectedCommand = $expectedFirstCommands[$catalogQuery]
        $suggestionText = [string]$excel.Run("'patriks-tab.xlam'!FAS_GetCommandHelpSuggestionText", $catalogQuery)
        if ($suggestionText -notmatch ("^Try: {0}(,|$)" -f [regex]::Escape($expectedCommand))) {
            throw "Expected exact command or alias '$expectedCommand' to rank first for '$catalogQuery'; got '$suggestionText'."
        }
    }

    Write-SmokeStep "Smoke: command palette form and shortcuts"
    $excel.Run("'patriks-tab.xlam'!FAS_ShowCommandPalette", "chart") | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_ToggleCommandPalette") | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RegisterShortcuts") | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_UnregisterShortcuts") | Out-Null

    Write-SmokeStep "Smoke: precedent helpers"
    $plConsolidatedSheet = $testWorkbook.Worksheets.Item($plConsolidatedName)
    $plStandaloneSheet = $testWorkbook.Worksheets.Item($plStandaloneName)
    $tbSheet = $testWorkbook.Worksheets.Item("BB 2024")
    $plConsolidatedSheet.Range("A1").Value2 = 10
    $plConsolidatedSheet.Range("B1").Value2 = 20
    $plConsolidatedSheet.Range("C1").Formula = "=A1+B1"
    $plStandaloneSheet.Range("A1").Value2 = 30
    $plConsolidatedSheet.Range("D1").Formula = "='" + $plStandaloneName + "'!A1+A1"

    $plConsolidatedSheet.Activate() | Out-Null
    $plConsolidatedSheet.Range("C1").Select() | Out-Null
    $precedentCount = $excel.Run("'patriks-tab.xlam'!FAS_CountPrecedentsInActiveCell")
    if ([int]$precedentCount -ne 2) {
        throw "Expected 2 same-sheet precedents for C1; got $precedentCount."
    }

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "pre") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plConsolidatedName -MacroName "FAS_JumpToPrecedentByIndex same-sheet"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A1" -CheckName "FAS_JumpToPrecedentByIndex same-sheet"

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "back") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plConsolidatedName -MacroName "FAS_RunCommand back"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "C1" -CheckName "FAS_RunCommand back"

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "pre2") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plConsolidatedName -MacroName "FAS_RunCommand pre2"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "B1" -CheckName "FAS_RunCommand pre2"

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "back") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "C1" -CheckName "FAS_RunCommand back after pre2"

    $excel.Run("'patriks-tab.xlam'!FAS_CycleDirectPrecedent") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A1" -CheckName "FAS_CycleDirectPrecedent first"
    $excel.Run("'patriks-tab.xlam'!FAS_CycleDirectPrecedent") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "B1" -CheckName "FAS_CycleDirectPrecedent second"
    $excel.Run("'patriks-tab.xlam'!FAS_CycleDirectPrecedent") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A1" -CheckName "FAS_CycleDirectPrecedent wrap"
    $excel.Run("'patriks-tab.xlam'!FAS_BackToTraceOrigin") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "C1" -CheckName "FAS_BackToTraceOrigin after cycle"

    $plConsolidatedSheet.Range("D1").Select() | Out-Null
    $precedentCount = $excel.Run("'patriks-tab.xlam'!FAS_CountPrecedentsInActiveCell")
    if ([int]$precedentCount -lt 2) {
        throw "Expected at least 2 mixed precedents for D1; got $precedentCount."
    }

    $excel.Run("'patriks-tab.xlam'!FAS_CycleDirectPrecedent") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plStandaloneName -MacroName "FAS_CycleDirectPrecedent cross-sheet"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A1" -CheckName "FAS_CycleDirectPrecedent cross-sheet"
    $excel.Run("'patriks-tab.xlam'!FAS_CycleDirectPrecedent") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $plConsolidatedName -MacroName "FAS_CycleDirectPrecedent mixed second"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A1" -CheckName "FAS_CycleDirectPrecedent mixed second"
    $excel.Run("'patriks-tab.xlam'!FAS_BackToTraceOrigin") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "D1" -CheckName "FAS_BackToTraceOrigin mixed"

    Write-SmokeStep "Smoke: TB account search"
    $tbSheet.Range("A1").Value2 = "Account"
    $tbSheet.Range("B1").Value2 = "Description"
    $tbSheet.Range("A2").Value2 = "1000"
    $tbSheet.Range("B2").Value2 = "Cash and cash equivalents"
    $tbSheet.Range("A3").Value2 = "4000"
    $tbSheet.Range("B3").Value2 = "Revenue"
    $tbSheet.Range("A4").Value2 = "100-200"
    $tbSheet.Range("B4").Value2 = "Hyphenated account"

    $testWorkbook.Activate()
    $matchCount = $excel.Run("'patriks-tab.xlam'!FAS_CountTBAccountMatches", "cash")
    if ([int]$matchCount -ne 1) {
        throw "Expected 1 TB account match for cash; got $matchCount."
    }

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "tb cash") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "BB 2024" -MacroName "FAS_RunCommand tb cash"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "B2" -CheckName "FAS_RunCommand tb cash"

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "tb: 100-200") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "BB 2024" -MacroName "FAS_RunCommand tb punctuation"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A4" -CheckName "FAS_RunCommand tb punctuation"

    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "acct cash")) {
        throw "FAS_IsKnownCommand did not recognize acct search."
    }

    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "mapinfo")) {
        throw "FAS_IsKnownCommand did not recognize mapinfo."
    }

    Write-SmokeStep "Smoke: GL and mapping lookup"
    $glSheet = $testWorkbook.Worksheets.Item($glSheetName)
    $mappingSheet = $testWorkbook.Worksheets.Item("Klasifikacija")
    $glSheet.Range("A2").Value2 = "Vendor cash payment"
    $mappingSheet.Range("A2").Value2 = "Revenue classification"

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "gl vendor") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName $glSheetName -MacroName "FAS_RunCommand gl vendor"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A2" -CheckName "FAS_RunCommand gl vendor"

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "map revenue") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "Klasifikacija" -MacroName "FAS_RunCommand map revenue"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "A2" -CheckName "FAS_RunCommand map revenue"

    Write-SmokeStep "Smoke: drilldown shape"
    $plStandaloneSheet.Activate() | Out-Null
    $plStandaloneSheet.Range("B10").Value2 = "100-200 Revaluation claims"
    $plStandaloneSheet.Range("C10").Formula = "='BB 2024'!B4"
    $mappingSheet.Range("A4").Value2 = "100-200"
    $mappingSheet.Range("B4").Value2 = "Revaluation claims"
    $tbSheet.Range("C4").Value2 = 12345
    $glSheet.Range("A4").Value2 = "100-200"
    $glSheet.Range("B4").Value2 = "Revaluation claims posting"
    $glSheet.Range("C4").Value2 = 12345

    $plStandaloneSheet.Range("C10").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "drill") | Out-Null
    $drillShapeCount = $excel.Run("'patriks-tab.xlam'!FAS_CountDrilldownShapesOnActiveSheet")
    if ([int]$drillShapeCount -ne 1) {
        throw "Expected 1 drilldown shape after drill; got $drillShapeCount."
    }
    $drillPlacement = $excel.Run("'patriks-tab.xlam'!FAS_GetActiveDrilldownPlacement")
    if ([int]$drillPlacement -ne 3) {
        throw "Expected drilldown shape placement 3/xlFreeFloating; got $drillPlacement."
    }
    $drillText = $excel.Run("'patriks-tab.xlam'!FAS_GetActiveDrilldownText")
    if ([string]$drillText -notmatch "100-200" -or [string]$drillText -notmatch "Mapping hits" -or [string]$drillText -notmatch "TB hits" -or [string]$drillText -notmatch "GL hits") {
        throw "Drilldown text did not include expected account and sections: $drillText"
    }

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "dd") | Out-Null
    $drillShapeCount = $excel.Run("'patriks-tab.xlam'!FAS_CountDrilldownShapesOnActiveSheet")
    if ([int]$drillShapeCount -ne 1) {
        throw "Expected drilldown replacement to keep 1 shape; got $drillShapeCount."
    }

    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "drillclear") | Out-Null
    $drillShapeCount = $excel.Run("'patriks-tab.xlam'!FAS_CountDrilldownShapesOnActiveSheet")
    if ([int]$drillShapeCount -ne 0) {
        throw "Expected drillclear to remove shape; got $drillShapeCount."
    }

    Write-SmokeStep "Smoke: selected object shape fix"
    $randomSheet = $testWorkbook.Worksheets.Item("Random Sheet")
    $testShape = $randomSheet.Shapes.AddShape(1, 20, 20, 80, 30)
    $testShape.Placement = 1
    $randomSheet.Activate() | Out-Null
    $testShape.Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "fixobj") | Out-Null
    Assert-ShapePlacement -Shape $testShape -ExpectedPlacement 3 -CheckName "FAS_RunCommand fixobj"

    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "fixshape")) {
        throw "FAS_IsKnownCommand did not recognize fixshape."
    }
    if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", "drillclear")) {
        throw "FAS_IsKnownCommand did not recognize drillclear."
    }

    Write-SmokeStep "Smoke: bookmarks"
    $randomSheet.Activate() | Out-Null
    $randomSheet.Range("B5").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_SaveActiveCellBookmark", "review") | Out-Null
    $randomSheet.Range("A1").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "bm review") | Out-Null
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "Random Sheet" -MacroName "FAS_RunCommand bm review"
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "B5" -CheckName "FAS_RunCommand bm review"

    $randomSheet.Range("C5").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_SaveActiveCellBookmark", "review|one") | Out-Null
    $randomSheet.Range("A1").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_GoToBookmark", "review|one") | Out-Null
    Assert-ActiveCellAddress -Excel $excel -ExpectedAddress "C5" -CheckName "FAS_GoToBookmark encoded name"

    Write-SmokeStep "Smoke: smart formula copy"
    $randomSheet.Range("A1").Value2 = 1
    $randomSheet.Range("B1").Value2 = 2
    $randomSheet.Range("D1").Value2 = 10
    $plStandaloneSheet.Range("A2").Value2 = 7
    $randomSheet.Range("C1").Formula = "=SUM(A1:B1)+'" + $plStandaloneName + "'!A2+D1"
    $randomSheet.Range("A1:C1").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "fcopy E1") | Out-Null
    $copiedFormula = $randomSheet.Range("G1").Formula
    $expectedCopiedFormula = "=SUM(E1:F1)+'" + $plStandaloneName + "'!`$A`$2+`$D`$1"
    if ($copiedFormula -ne $expectedCopiedFormula) {
        throw "Smart formula copy produced '$copiedFormula' instead of '$expectedCopiedFormula'."
    }

    $randomSheet.Range("C2").Formula = "=A1&"" [ok] # """
    $randomSheet.Range("A1:C2").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_CopySelectionFormulasSmartToAddress", "E3") | Out-Null
    $copiedFormula = $randomSheet.Range("G4").Formula
    if ($copiedFormula -ne "=E3&"" [ok] # """) {
        throw "Smart formula copy incorrectly rejected or rewrote quoted bracket/hash text: '$copiedFormula'."
    }

    Write-SmokeStep "Smoke: productivity copy, wrappers, links, color, delta, and width"
    $randomSheet.Range("A10").Value2 = 1
    $randomSheet.Range("A11").Value2 = 2
    $randomSheet.Range("A12").Value2 = 3
    $randomSheet.Rows.Item(11).Hidden = $true
    $randomSheet.Range("A10:A12").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_CopyVisibleOnly") | Out-Null
    if (-not [bool]$excel.CutCopyMode) {
        throw "Visible-only copy did not leave a copy operation ready."
    }
    $randomSheet.Range("E10").Select() | Out-Null
    $randomSheet.Paste($randomSheet.Range("E10"))
    if ($randomSheet.Range("E10").Value2 -ne 1 -or $randomSheet.Range("E11").Value2 -eq 2) {
        throw "Visible-only copy included a hidden row or failed to paste the visible source."
    }
    $randomSheet.Rows.Item(11).Hidden = $false

    $randomSheet.Range("B10").Formula = "=1/3"
    $randomSheet.Range("B10").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_WrapSelectedFormulasRound", 2) | Out-Null
    if ($randomSheet.Range("B10").Formula -ne "=ROUND(1/3,2)") {
        throw "ROUND wrapper produced '$($randomSheet.Range("B10").Formula)'."
    }

    $randomSheet.Range("B11").Formula = "=1/0"
    $randomSheet.Range("B11").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_WrapSelectedFormulasIfError") | Out-Null
    if ($randomSheet.Range("B11").Formula -ne "=IFERROR(1/0,"""")") {
        throw "IFERROR wrapper produced '$($randomSheet.Range("B11").Formula)'."
    }

    $randomSheet.Range("B12").Formula = "=1/3"
    $randomSheet.Range("B12").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_RunCommand", "round") | Out-Null
    if ($randomSheet.Range("B12").Formula -ne "=ROUND(1/3,2)") {
        throw "Saved ROUND digits were not reused by the default command."
    }

    $randomSheet.Range("C10").Value2 = "linked value"
    $randomSheet.Range("D10").Formula = "=1+1"
    $randomSheet.Range("C10").Font.Color = 255
    $randomSheet.Range("C10:D10").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_CopyLinks") | Out-Null
    $plStandaloneSheet.Activate() | Out-Null
    $plStandaloneSheet.Range("H5").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_PasteLinks") | Out-Null
    if ($plStandaloneSheet.Range("H5").Formula -ne "='Random Sheet'!`$C`$10") {
        throw "Linked paste produced '$($plStandaloneSheet.Range("H5").Formula)'."
    }
    if ([int]$plStandaloneSheet.Range("H5").Font.Color -ne 255) {
        throw "Linked paste did not copy source formatting."
    }

    $randomSheet.Activate() | Out-Null
    $randomSheet.Range("F10").Font.Color = 255
    $randomSheet.Range("F10").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_CaptureCurrentColor") | Out-Null
    $randomSheet.Range("F11").Font.Color = 65280
    $randomSheet.Range("F11").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_ApplyCurrentColor") | Out-Null
    if ([int]$randomSheet.Range("F11").Font.Color -ne 255) {
        throw "Captured Excel font color was not applied."
    }

    $randomSheet.Range("F12").ClearContents()
    $randomSheet.Range("F12").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_InsertDelta") | Out-Null
    if ([string]$randomSheet.Range("F12").Value2 -ne ([string][char]0x394)) {
        throw "Delta insertion produced '$($randomSheet.Range("F12").Value2)'."
    }

    $randomSheet.Range("G10").Value2 = "a deliberately longer header for width smoke"
    $randomSheet.Range("G10").EntireColumn.AutoFit()
    $baseWidth = [double]$randomSheet.Range("G10").EntireColumn.ColumnWidth
    $randomSheet.Range("G10").EntireColumn.ColumnWidth = 8.43
    $randomSheet.Range("G10").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_AutoFitWidthWithMargin") | Out-Null
    $actualWidth = [double]$randomSheet.Range("G10").EntireColumn.ColumnWidth
    if ([math]::Abs($actualWidth - ($baseWidth * 1.2)) -gt 0.2) {
        throw "Width autofit produced $actualWidth instead of approximately $($baseWidth * 1.2)."
    }

    foreach ($knownProductivityCommand in @("visible", "round", "iferror", "linkcopy", "linkpaste", "colorcapture", "colorapply", "delta", "autofit")) {
        if (-not $excel.Run("'patriks-tab.xlam'!FAS_IsKnownCommand", $knownProductivityCommand)) {
            throw "FAS_IsKnownCommand did not recognize $knownProductivityCommand."
        }
    }

    Write-SmokeStep "Smoke: subtotals"
    $randomSheet.Range("D1").Value2 = "header"
    $randomSheet.Range("D1:D2").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_InsertSubtotalAboveSelection") | Out-Null
    if ($randomSheet.Range("D1").Value2 -ne "header") {
        throw "Above subtotal changed row 1 instead of refusing a selection with no row above."
    }

    $randomSheet.Range("D5").Value2 = 5
    $randomSheet.Range("D6").Value2 = 6
    $randomSheet.Range("D5:D6").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_InsertSubtotalAboveSelection") | Out-Null
    if ($randomSheet.Range("D4").Formula -ne "=SUBTOTAL(9,D5:D6)") {
        throw "Above subtotal formula was '$($randomSheet.Range("D4").Formula)' instead of '=SUBTOTAL(9,D5:D6)'."
    }
    if ($randomSheet.Range("D5").Value2 -ne 5 -or $randomSheet.Range("D6").Value2 -ne 6) {
        throw "Above subtotal shifted the selected data instead of writing into the existing row above."
    }

    $randomSheet.Range("A2").Value2 = 5
    $randomSheet.Range("A3").Value2 = 6
    $randomSheet.Range("A2:A3").Select() | Out-Null
    $excel.Run("'patriks-tab.xlam'!FAS_InsertSubtotalBelowSelection") | Out-Null
    if ($randomSheet.Range("A4").Formula -ne "=SUBTOTAL(109,A2:A3)") {
        throw "Subtotal formula was '$($randomSheet.Range("A4").Formula)' instead of '=SUBTOTAL(109,A2:A3)'."
    }

    Write-SmokeStep "Smoke: divider sheet"
    $testWorkbook.Worksheets.Item(1).Activate() | Out-Null
    $createdDivider = $excel.Run("'patriks-tab.xlam'!FAS_CreateActiveWorkbookDividerSheet", "P&Ls >", 1)
    if (-not $createdDivider) {
        throw "FAS_CreateActiveWorkbookDividerSheet returned false."
    }
    Assert-ActiveSheetName -Workbook $testWorkbook -ExpectedSheetName "P&Ls >" -MacroName "FAS_CreateDividerSheet"

    Write-Host "Smoke test passed."
}
finally {
    if ($null -ne $testWorkbook) {
        $testWorkbook.Close($false)
    }

    if ($null -ne $addInWorkbook) {
        $addInWorkbook.Close($false)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    if (Test-Path -LiteralPath $tempWorkbookPath) {
        Remove-Item -LiteralPath $tempWorkbookPath -Force
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
