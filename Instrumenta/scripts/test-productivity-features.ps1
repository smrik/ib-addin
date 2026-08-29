[CmdletBinding()]
param(
    [switch]$ComSmoke,
    [switch]$ForceClosePowerPoint
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $repoRoot 'src\Modules\ModuleFASProductivity.bas'
$callbacksPath = Join-Path $repoRoot 'src\Modules\ModuleCallbacks.bas'
$featureSearchPath = Join-Path $repoRoot 'src\Modules\ModuleFeatureSearch.bas'
$allowlistPath = Join-Path $repoRoot 'src\Modules\ModuleSecurityAllowlist.bas'
$customUiPath = Join-Path $repoRoot 'src\CustomUI\CustomUI.xml'
$devAddInPath = Join-Path $repoRoot 'bin\InstrumentaPowerpointToolbarDev.ppam'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-ContainsLiteral {
    param([string]$Path, [string]$Literal, [string]$Message)
    $content = Get-Content -LiteralPath $Path -Raw
    Assert-True -Condition $content.Contains($Literal) -Message $Message
}

function Close-ComObject {
    param($ComObject)
    if ($null -ne $ComObject) {
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject) } catch {}
    }
}

function Close-PowerPointSmokeHost {
    param(
        $Presentation,
        $Application
    )

    # Office can briefly reject calls while its worker is closing. Cleanup is
    # best-effort so a shutdown race does not fail feature assertions.
    if ($null -ne $Presentation) {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $Presentation.Close()
                break
            }
            catch {
                Start-Sleep -Milliseconds 250
            }
        }
    }

    if ($null -ne $Application) {
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $Application.Quit()
                break
            }
            catch {
                Start-Sleep -Milliseconds 250
            }
        }
    }
}

foreach ($requiredPath in @($modulePath, $callbacksPath, $featureSearchPath, $allowlistPath, $customUiPath)) {
    Assert-True -Condition (Test-Path -LiteralPath $requiredPath) -Message "Required productivity source is missing: $requiredPath"
}

foreach ($macroPattern in @(
    'Public Sub FASCaptureCurrentColor()',
    'Public Sub FASApplyCurrentColor()',
    'Public Sub FASLightenSelectedColor',
    'Public Sub FASConvertSelectedCirclesToArcs()',
    'Public Sub FASConnectSelectedArcCenters()',
    'Public Sub FASConvertCirclesToArcsAndConnectCenters()',
    'Public Sub FASApplyLastLinePreset()',
    'Public Sub FASApplyBlueDashedLinePreset()',
    'Public Sub FASApplyGraySolidLinePreset()'
)) {
    Assert-ContainsLiteral -Path $modulePath -Literal $macroPattern -Message "Productivity macro contract is missing '$macroPattern'."
}

Assert-ContainsLiteral -Path $callbacksPath -Literal 'Sub TextInsertDelta()' -Message 'Delta callback is missing.'
Assert-ContainsLiteral -Path $featureSearchPath -Literal '"FASConnectSelectedArcCenters"' -Message 'Feature Search is missing the center connector entry.'
Assert-ContainsLiteral -Path $allowlistPath -Literal '"fasconnectselectedarccenters"' -Message 'Security allowlist is missing the center connector entry.'
Assert-ContainsLiteral -Path $allowlistPath -Literal '"fasconvertcirclestoarcsandconnectcenters"' -Message 'Security allowlist is missing the combined arc workflow entry.'
Assert-ContainsLiteral -Path $modulePath -Literal 'SaveSetting "Instrumenta", "Productivity", "LightenPercent"' -Message 'Lighten choice is not persisted.'
Assert-ContainsLiteral -Path $modulePath -Literal 'GetSetting("Instrumenta", "Productivity", "LastLinePreset"' -Message 'Last line preset is not reused.'
Assert-ContainsLiteral -Path $customUiPath -Literal 'id="FASProductivityMenu"' -Message 'The one-tab Ribbon is missing the productivity menu.'
Assert-ContainsLiteral -Path $customUiPath -Literal 'id="TabViewFASProductivityMenu"' -Message 'The multi-tab Ribbon is missing the productivity menu.'
[xml](Get-Content -LiteralPath $customUiPath -Raw) | Out-Null

if (-not $ComSmoke) {
    Write-Host 'PowerPoint productivity source contracts passed.'
    exit 0
}

if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
    if (-not $ForceClosePowerPoint) {
        throw 'PowerPoint must be closed before COM smoke, or rerun with -ForceClosePowerPoint.'
    }
    Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
}

$app = $null
$addIn = $null
$presentation = $null
$slide = $null
$firstShape = $null
$secondShape = $null
$connector = $null

try {
    Assert-True -Condition (Test-Path -LiteralPath $devAddInPath) -Message "Build the dev PPAM before COM smoke: $devAddInPath"
    $app = New-Object -ComObject PowerPoint.Application
    $app.Visible = -1
    $addIn = $app.AddIns.Add($devAddInPath)
    $addIn.Loaded = -1
    $presentation = $app.Presentations.Add(-1)
    $slide = $presentation.Slides.Add(1, 12)

    $firstShape = $slide.Shapes.AddShape(9, 100, 100, 100, 100)
    $secondShape = $slide.Shapes.AddShape(9, 350, 100, 100, 100)

    $firstShape.Select()
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASCaptureCurrentColor") | Out-Null
    $secondShape.Fill.ForeColor.RGB = 65280
    $secondShape.Select()
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASApplyCurrentColor") | Out-Null
    Assert-True -Condition ([int]$secondShape.Fill.ForeColor.RGB -eq 255) -Message 'Captured fill color was not applied.'

    $firstShape.Fill.Transparency = 0.25
    $firstShape.Select()
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASLightenSelectedColor") | Out-Null
    Assert-True -Condition ([int]$firstShape.Fill.ForeColor.RGB -ne 255) -Message 'Lighten color did not change the fill tone.'
    Assert-True -Condition ([math]::Abs([double]$firstShape.Fill.Transparency - 0.25) -lt 0.01) -Message 'Lighten color changed transparency.'

    $firstShape.Select()
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASConvertSelectedCirclesToArcs") | Out-Null
    $secondShape.Select()
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASConvertSelectedCirclesToArcs") | Out-Null
    Assert-True -Condition ([int]$firstShape.AutoShapeType -ne 9) -Message 'First circle was not converted to an arc.'
    Assert-True -Condition ([int]$secondShape.AutoShapeType -ne 9) -Message 'Second circle was not converted to an arc.'

    $firstShape.Select()
    $secondShape.Select($false)
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASConnectSelectedArcCenters") | Out-Null
    Assert-True -Condition ([int]$slide.Shapes.Count -eq 3) -Message 'Center connector was not created.'
    $connector = $slide.Shapes.Item(3)
    $connector.Select()
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASApplyGraySolidLinePreset") | Out-Null
    $connector.Line.ForeColor.RGB = 255
    $app.Run("'$([IO.Path]::GetFileName($devAddInPath))'!FASApplyLastLinePreset") | Out-Null
    Assert-True -Condition ([int]$connector.Line.ForeColor.RGB -eq 8421504) -Message 'Last line preset was not reused.'

    Write-Host 'PowerPoint productivity COM smoke passed.'
}
finally {
    Close-PowerPointSmokeHost -Presentation $presentation -Application $app
    Close-ComObject $connector
    Close-ComObject $secondShape
    Close-ComObject $firstShape
    Close-ComObject $slide
    Close-ComObject $presentation
    Close-ComObject $addIn
    Close-ComObject $app
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
