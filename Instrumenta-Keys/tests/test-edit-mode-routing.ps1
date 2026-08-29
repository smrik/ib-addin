Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$routingPath = Join-Path $repoRoot 'src\windows\EditModeRouting.ps1'

if (-not (Test-Path -LiteralPath $routingPath)) {
    throw "Edit Mode routing module is missing: $routingPath"
}

. $routingPath

function Assert-Equal {
    param(
        $Actual,
        $Expected,
        [string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected', got '$Actual'."
    }
}

function New-TestKeyEvent {
    param(
        [int]$VirtualKey,
        [int]$ScanCode,
        [bool]$Extended = $false,
        [bool]$Ctrl = $false,
        [bool]$Shift = $false,
        [bool]$Alt = $false,
        [bool]$Repeat = $false
    )

    return [pscustomobject]@{
        VirtualKey = $VirtualKey
        ScanCode   = $ScanCode
        Extended   = $Extended
        Ctrl       = $Ctrl
        Shift      = $Shift
        Alt        = $Alt
        Repeat     = $Repeat
    }
}

$inactiveState = [pscustomobject]@{
    Active               = $false
    PowerPointForeground = $true
    PidMatches           = $true
    NumLockOn            = $false
}

$activeState = [pscustomobject]@{
    Active               = $true
    PowerPointForeground = $true
    PidMatches           = $true
    NumLockOn            = $false
}

$notForegroundState = [pscustomobject]@{
    Active               = $true
    PowerPointForeground = $false
    PidMatches           = $false
    NumLockOn            = $false
}

$tests = @(
    @{
        Name     = 'keypad Insert toggles inactive mode'
        Event    = New-TestKeyEvent -VirtualKey 0x2D -ScanCode 0x52
        State    = $inactiveState
        Macro    = 'FASEditModeToggle'
        Suppress = $true
    },
    @{
        Name     = 'enhanced Insert is not treated as keypad zero'
        Event    = New-TestKeyEvent -VirtualKey 0x2D -ScanCode 0x52 -Extended $true
        State    = $inactiveState
        Macro    = $null
        Suppress = $false
    },
    @{
        Name     = 'keypad Home maps to top-left'
        Event    = New-TestKeyEvent -VirtualKey 0x24 -ScanCode 0x47
        State    = $activeState
        Macro    = 'FASEditModePlaceTopLeft'
        Suppress = $true
    },
    @{
        Name     = 'Ctrl keypad PageDown maps to fit bottom-right'
        Event    = New-TestKeyEvent -VirtualKey 0x22 -ScanCode 0x51 -Ctrl $true
        State    = $activeState
        Macro    = 'FASEditModeFitBottomRight'
        Suppress = $true
    },
    @{
        Name     = 'enhanced Right arrow nudges right'
        Event    = New-TestKeyEvent -VirtualKey 0x27 -ScanCode 0x4D -Extended $true
        State    = $activeState
        Macro    = 'FASEditModeNudgeRightFine'
        Suppress = $true
    },
    @{
        Name     = 'Shift enhanced Left arrow performs coarse nudge'
        Event    = New-TestKeyEvent -VirtualKey 0x25 -ScanCode 0x4B -Extended $true -Shift $true
        State    = $activeState
        Macro    = 'FASEditModeNudgeLeftCoarse'
        Suppress = $true
    },
    @{
        Name     = 'Ctrl enhanced Down arrow increases height'
        Event    = New-TestKeyEvent -VirtualKey 0x28 -ScanCode 0x50 -Extended $true -Ctrl $true
        State    = $activeState
        Macro    = 'FASEditModeResizeHeightIncreaseFine'
        Suppress = $true
    },
    @{
        Name     = 'Ctrl Shift enhanced Up arrow decreases height coarsely'
        Event    = New-TestKeyEvent -VirtualKey 0x26 -ScanCode 0x48 -Extended $true -Ctrl $true -Shift $true
        State    = $activeState
        Macro    = 'FASEditModeResizeHeightDecreaseCoarse'
        Suppress = $true
    },
    @{
        Name     = 'Shift keypad plus brings forward'
        Event    = New-TestKeyEvent -VirtualKey 0x6B -ScanCode 0x4E -Shift $true
        State    = $activeState
        Macro    = 'FASEditModeBringForward'
        Suppress = $true
    },
    @{
        Name     = 'Ctrl Shift keypad minus sends to back'
        Event    = New-TestKeyEvent -VirtualKey 0x6D -ScanCode 0x4A -Ctrl $true -Shift $true
        State    = $activeState
        Macro    = 'FASEditModeSendToBack'
        Suppress = $true
    },
    @{
        Name     = 'Escape exits active mode'
        Event    = New-TestKeyEvent -VirtualKey 0x1B -ScanCode 0x01
        State    = $activeState
        Macro    = 'FASEditModeExit'
        Suppress = $true
    },
    @{
        Name     = 'keys pass through outside PowerPoint'
        Event    = New-TestKeyEvent -VirtualKey 0x27 -ScanCode 0x4D -Extended $true
        State    = $notForegroundState
        Macro    = $null
        Suppress = $false
    }
)

foreach ($test in $tests) {
    $result = Resolve-InstrumentaEditModeAction -KeyEvent $test.Event -State $test.State
    Assert-Equal -Actual $result.Macro -Expected $test.Macro -Message $test.Name
    Assert-Equal -Actual $result.Suppress -Expected $test.Suppress -Message ($test.Name + ' suppression')
}

$numLockState = [pscustomobject]@{
    Active               = $true
    PowerPointForeground = $true
    PidMatches           = $true
    NumLockOn            = $true
}
$numLockResult = Resolve-InstrumentaEditModeAction `
    -KeyEvent (New-TestKeyEvent -VirtualKey 0x61 -ScanCode 0x4F) `
    -State $numLockState
Assert-Equal -Actual $numLockResult.Macro -Expected $null -Message 'NumLock-on numeric entry must pass through'
Assert-Equal -Actual $numLockResult.Suppress -Expected $false -Message 'NumLock-on numeric entry suppression'

Write-Host "Instrumenta Keys Edit Mode routing tests passed ($($tests.Count + 1) cases)."
