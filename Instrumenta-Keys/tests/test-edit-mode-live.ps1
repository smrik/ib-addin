[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$companionPath = Join-Path $repoRoot 'bin\windows\Instrumenta Keys.exe'
$smokeDirectory = Join-Path $repoRoot '.instrumenta-live-smoke'
$powerPoint = $null
$presentation = $null
$slide = $null
$shape = $null
$companion = $null

function Close-ComObject {
    param($ComObject)

    if ($null -ne $ComObject) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
        catch {
        }
    }
}

if (-not (Test-Path -LiteralPath $companionPath)) {
    throw "Build Instrumenta Keys before live smoke: $companionPath"
}
if (@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'PowerPoint must be closed before the Instrumenta Keys live smoke.'
}
if (@(Get-Process -Name 'Instrumenta Keys' -ErrorAction SilentlyContinue).Count -gt 0) {
    throw 'Instrumenta Keys must be closed before its live smoke.'
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class InstrumentaLiveSmokeInput {
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
    private const uint KEYEVENTF_KEYUP = 0x0002;

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern void keybd_event(byte virtualKey, byte scanCode, uint flags, UIntPtr extraInfo);

    public static void PressExtendedRightArrow() {
        keybd_event(0x27, 0x4D, KEYEVENTF_EXTENDEDKEY, UIntPtr.Zero);
        keybd_event(0x27, 0x4D, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, UIntPtr.Zero);
    }
}
"@

try {
    New-Item -ItemType Directory -Path $smokeDirectory -Force | Out-Null
    $companion = Start-Process `
        -FilePath $companionPath `
        -WorkingDirectory $smokeDirectory `
        -WindowStyle Hidden `
        -PassThru
    Start-Sleep -Seconds 5
    $companion.Refresh()
    if ($companion.HasExited) {
        throw 'Instrumenta Keys exited during live smoke startup.'
    }

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Add(-1)
    $slide = $presentation.Slides.Add(1, 12)
    $shape = $slide.Shapes.AddShape(1, 200, 200, 100, 50)
    $shape.Select()

    [void]$powerPoint.Run("'InstrumentaPowerpointToolbar.ppam'!FASEditModeEnterForAutomation")
    Start-Sleep -Milliseconds 500
    $leftBefore = [double]$shape.Left

    $powerPointProcess = Get-Process POWERPNT -ErrorAction Stop | Select-Object -First 1
    $windowHandle = [IntPtr]$powerPointProcess.MainWindowHandle
    if ($windowHandle -eq [IntPtr]::Zero) {
        throw 'The temporary PowerPoint process did not expose a main window handle.'
    }
    if (-not [InstrumentaLiveSmokeInput]::SetForegroundWindow($windowHandle)) {
        throw 'Could not foreground the temporary PowerPoint window.'
    }
    Start-Sleep -Milliseconds 500
    [InstrumentaLiveSmokeInput]::PressExtendedRightArrow()
    Start-Sleep -Seconds 2

    $leftAfter = [double]$shape.Left
    if ([Math]::Abs($leftAfter - ($leftBefore + 1)) -gt 0.1) {
        throw "The live Right Arrow did not nudge by 1 point. Before=$leftBefore After=$leftAfter"
    }

    [void]$powerPoint.Run("'InstrumentaPowerpointToolbar.ppam'!FASEditModeExit")
    Write-Host 'Instrumenta Keys live Edit Mode hook smoke passed.'
}
finally {
    if ($null -ne $powerPoint) {
        try { [void]$powerPoint.Run("'InstrumentaPowerpointToolbar.ppam'!FASEditModeExit") } catch {}
    }
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch {}
    }
    if ($null -ne $powerPoint) {
        try { $powerPoint.Quit() } catch {}
    }
    Close-ComObject $shape
    Close-ComObject $slide
    Close-ComObject $presentation
    Close-ComObject $powerPoint

    if ($null -ne $companion) {
        try {
            $companion.Refresh()
            if (-not $companion.HasExited) {
                Stop-Process -Id $companion.Id
                [void]$companion.WaitForExit(5000)
            }
        }
        catch {
        }
    }

    $resolvedSmokeDirectory = [IO.Path]::GetFullPath($smokeDirectory)
    $resolvedRepoRoot = [IO.Path]::GetFullPath($repoRoot)
    if (
        (Test-Path -LiteralPath $resolvedSmokeDirectory) -and
        $resolvedSmokeDirectory.StartsWith($resolvedRepoRoot, [StringComparison]::OrdinalIgnoreCase)
    ) {
        Remove-Item -LiteralPath $resolvedSmokeDirectory -Recurse -Force
    }
}
