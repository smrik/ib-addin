[CmdletBinding()]
param(
    [switch]$ComSmoke
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $repoRoot 'src\Modules\ModuleFASEditMode.bas'
$alignmentPath = Join-Path $repoRoot 'src\Modules\ModuleFASAlignment.bas'
$eventsPath = Join-Path $repoRoot 'src\Classes\FASEditModeEventsClass.cls'
$formPath = Join-Path $repoRoot 'src\Forms\FASEditModeForm.frm'
$workerPath = Join-Path $repoRoot 'scripts\instrumenta-build-worker.ps1'
$customUiPath = Join-Path $repoRoot 'src\CustomUI\CustomUI.xml'
$featureSearchPath = Join-Path $repoRoot 'src\Modules\ModuleFeatureSearch.bas'
$allowlistPath = Join-Path $repoRoot 'src\Modules\ModuleSecurityAllowlist.bas'
$sourcePptmPath = Join-Path $repoRoot 'src\InstrumentaPowerpointToolbar.pptm'
$devAddInPath = Join-Path $repoRoot 'bin\InstrumentaPowerpointToolbarDev.ppam'

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Near {
    param(
        [double]$Actual,
        [double]$Expected,
        [double]$Tolerance,
        [string]$Message
    )

    if ([math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw "$Message Expected $Expected +/- $Tolerance, got $Actual."
    }
}

function Assert-ContainsLiteral {
    param(
        [string]$Path,
        [string]$Literal,
        [string]$Message
    )

    $content = Get-Content -LiteralPath $Path -Raw
    Assert-True -Condition $content.Contains($Literal) -Message $Message
}

function Close-ComObject {
    param($ComObject)

    if ($null -ne $ComObject) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
        }
    }
}

function Invoke-PowerPointMacro {
    param(
        $Application,
        [string]$MacroName
    )

    $qualifiedMacroName = "'$([IO.Path]::GetFileName($devAddInPath))'!$MacroName"
    try {
        return $Application.Run($qualifiedMacroName)
    }
    catch {
        throw "Macro '$qualifiedMacroName' failed: $($_.Exception.Message)"
    }
}

function Get-EditModeRegistryValue {
    param([string]$Name)

    $path = 'HKCU:\Software\VB and VBA Program Settings\InstrumentaFASEditMode\Session'
    if (-not (Test-Path -LiteralPath $path)) {
        return $null
    }

    $item = Get-ItemProperty -LiteralPath $path
    if ($item.PSObject.Properties.Name -notcontains $Name) {
        return $null
    }

    return [string]$item.$Name
}

foreach ($requiredPath in @(
    $modulePath,
    $alignmentPath,
    $eventsPath,
    $formPath,
    $workerPath,
    $customUiPath,
    $featureSearchPath,
    $allowlistPath,
    $sourcePptmPath
)) {
    Assert-True -Condition (Test-Path -LiteralPath $requiredPath) -Message "Required Edit Mode source is missing: $requiredPath"
}

$publicMacroPatterns = @(
    'Public Sub FASEditModeToggle()',
    'Public Sub FASEditModeExit()',
    'Public Sub FASEditModeNudgeLeftFine()',
    'Public Sub FASEditModeNudgeRightCoarse()',
    'Public Sub FASEditModeResizeWidthIncreaseFine()',
    'Public Sub FASEditModeResizeHeightDecreaseCoarse()',
    'Public Sub FASEditModePlaceTopLeft()',
    'Public Sub FASEditModePlaceCenter()',
    'Public Sub FASEditModeFitBottomRight()',
    'Public Sub FASEditModeBringForward()',
    'Public Sub FASEditModeSendToBack()'
)

foreach ($macroPattern in $publicMacroPatterns) {
    Assert-ContainsLiteral -Path $modulePath -Literal $macroPattern -Message "Edit Mode public macro contract is missing '$macroPattern'."
}

Assert-ContainsLiteral -Path $alignmentPath -Literal 'Public Sub FASApplyEditCell' -Message 'FAS alignment does not expose the shared Edit Mode grid dispatcher.'
Assert-ContainsLiteral -Path $workerPath -Literal 'Sync-FasEditModeForm' -Message 'The build worker does not create and synchronize the modeless Edit Mode form.'
Assert-ContainsLiteral -Path $customUiPath -Literal 'FASEditModeToggleButton' -Message 'The one-tab Ribbon is missing the FAS Edit Mode control.'
Assert-ContainsLiteral -Path $customUiPath -Literal 'TabViewFASEditModeToggleButton' -Message 'The multi-tab Ribbon is missing the FAS Edit Mode control.'
Assert-ContainsLiteral -Path $featureSearchPath -Literal '"FASEditModeToggleButton"' -Message 'Feature Search is missing the FAS Edit Mode entry.'
Assert-ContainsLiteral -Path $allowlistPath -Literal '"faseditmodetoggle"' -Message 'The security allowlist is missing the FAS Edit Mode entry point.'

[xml](Get-Content -LiteralPath $customUiPath -Raw) | Out-Null

if (-not $ComSmoke) {
    Write-Host 'FAS Edit Mode source contracts passed.'
    exit 0
}

if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'PowerPoint must be closed before running the FAS Edit Mode COM smoke.'
}

$testHostPath = $sourcePptmPath
$app = $null
$editModeAddIn = $null
$testPresentation = $null
$slide = $null
$shape = $null
$secondShape = $null

try {
    Assert-True -Condition (Test-Path -LiteralPath $devAddInPath) -Message "Build the dev PPAM before COM smoke: $devAddInPath"

    Write-Host 'COM smoke: starting PowerPoint'
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = -1
    Write-Host 'COM smoke: loading dev add-in'
    $editModeAddIn = $app.AddIns.Add($devAddInPath)
    $editModeAddIn.Loaded = -1
    Write-Host 'COM smoke: creating test slide and shape'
    $testPresentation = $app.Presentations.Add(-1)
    $slide = $testPresentation.Slides.Add(1, 12)
    $shape = $slide.Shapes.AddShape(1, 200, 200, 100, 50)
    $shape.Select()

    Write-Host 'COM smoke: entering Edit Mode without the modeless UI'
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeEnterForAutomation'
    Assert-True -Condition ((Get-EditModeRegistryValue -Name 'Active') -eq '1') -Message 'Toggling Edit Mode did not publish Active=1.'
    Assert-True -Condition ([int](Get-EditModeRegistryValue -Name 'PowerPointPid') -gt 0) -Message 'Toggling Edit Mode did not publish a PowerPoint PID.'

    Write-Host 'COM smoke: testing placement'
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModePlaceTopLeft'
    $pointsPerCm = 72 / 2.54
    Assert-Near -Actual $shape.Left -Expected (0.83 * $pointsPerCm) -Tolerance 0.2 -Message 'Top-left placement used the wrong FAS left boundary.'
    Assert-Near -Actual $shape.Top -Expected (4.3 * $pointsPerCm) -Tolerance 0.2 -Message 'Top-left placement used the wrong FAS top boundary.'

    $leftBeforeNudge = [double]$shape.Left
    Write-Host 'COM smoke: testing nudge'
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeNudgeRightFine'
    Assert-Near -Actual $shape.Left -Expected ($leftBeforeNudge + 1) -Tolerance 0.05 -Message 'Fine right nudge did not move by one point.'

    $widthBeforeResize = [double]$shape.Width
    Write-Host 'COM smoke: testing resize'
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeResizeWidthIncreaseCoarse'
    Assert-Near -Actual $shape.Width -Expected ($widthBeforeResize + 10) -Tolerance 0.1 -Message 'Coarse width increase did not resize by ten points.'

    Write-Host 'COM smoke: testing resize zone clamp'
    $shape.Left = 0
    $shape.Top = 0
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeResizeWidthIncreaseFine'
    Assert-True -Condition ($shape.Left -ge ((0.83 * $pointsPerCm) - 0.2)) -Message 'Resize restored an out-of-zone left position after clamping.'
    Assert-True -Condition ($shape.Top -ge ((4.3 * $pointsPerCm) - 0.2)) -Message 'Resize restored an out-of-zone top position after clamping.'

    Write-Host 'COM smoke: testing oversized selection rejection'
    $shape.Width = 1000
    $leftBeforeOversizedNudge = [double]$shape.Left
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeNudgeRightFine'
    Assert-Near -Actual $shape.Left -Expected $leftBeforeOversizedNudge -Tolerance 0.05 -Message 'An oversized selection should not be nudged outside the active zone.'

    $secondShape = $slide.Shapes.AddShape(1, 350, 200, 100, 50)
    $shape.Select()
    $positionBefore = [int]$shape.ZOrderPosition
    Write-Host 'COM smoke: testing z-order'
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeBringForward'
    Assert-True -Condition ([int]$shape.ZOrderPosition -gt $positionBefore) -Message 'Bring forward did not change the selected shape z-order.'

    Write-Host 'COM smoke: exiting Edit Mode'
    Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeExit'
    Assert-True -Condition ((Get-EditModeRegistryValue -Name 'Active') -eq '0') -Message 'Exiting Edit Mode did not publish Active=0.'

    Write-Host 'FAS Edit Mode COM smoke passed.'
}
finally {
    try {
        if ($null -ne $app) {
            try { Invoke-PowerPointMacro -Application $app -MacroName 'FASEditModeExit' | Out-Null } catch {}
        }
        if ($null -ne $testPresentation) {
            $testPresentation.Close()
        }
        if ($null -ne $app) {
            $app.Quit()
        }
    }
    finally {
        Close-ComObject $secondShape
        Close-ComObject $shape
        Close-ComObject $slide
        Close-ComObject $testPresentation
        Close-ComObject $editModeAddIn
        Close-ComObject $app
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
