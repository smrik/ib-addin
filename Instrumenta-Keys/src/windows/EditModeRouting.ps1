# MIT License
#
# Pure key-event routing for Instrumenta FAS Edit Mode. Keeping this free of
# hooks and COM makes the keyboard contract fast to test.

function New-InstrumentaEditModeAction {
    param(
        [AllowNull()]
        $Macro,
        [bool]$Suppress = $false
    )

    [pscustomobject]@{
        Macro    = $Macro
        Suppress = $Suppress
    }
}

function Test-InstrumentaKeypadKey {
    param(
        $KeyEvent,
        [int]$VirtualKey,
        [int]$ScanCode
    )

    return (
        [int]$KeyEvent.VirtualKey -eq $VirtualKey -and
        [int]$KeyEvent.ScanCode -eq $ScanCode -and
        -not [bool]$KeyEvent.Extended
    )
}

function Resolve-InstrumentaEditModeAction {
    param(
        [Parameter(Mandatory)]
        $KeyEvent,
        [Parameter(Mandatory)]
        $State
    )

    $passThrough = New-InstrumentaEditModeAction -Macro $null -Suppress $false
    if (-not [bool]$State.PowerPointForeground -or -not [bool]$State.PidMatches) {
        return $passThrough
    }

    $isKeypadZero = Test-InstrumentaKeypadKey -KeyEvent $KeyEvent -VirtualKey 0x2D -ScanCode 0x52
    if ($isKeypadZero -and -not $KeyEvent.Ctrl -and -not $KeyEvent.Shift -and -not $KeyEvent.Alt) {
        return New-InstrumentaEditModeAction -Macro 'FASEditModeToggle' -Suppress $true
    }

    if (-not [bool]$State.Active) {
        return $passThrough
    }

    if ([int]$KeyEvent.VirtualKey -eq 0x1B) {
        return New-InstrumentaEditModeAction -Macro 'FASEditModeExit' -Suppress $true
    }

    if ($KeyEvent.Alt) {
        return $passThrough
    }

    $keypadPlacements = @(
        @{ VirtualKey = 0x24; ScanCode = 0x47; Place = 'FASEditModePlaceTopLeft';     Fit = 'FASEditModeFitTopLeft' },
        @{ VirtualKey = 0x26; ScanCode = 0x48; Place = 'FASEditModePlaceTop';         Fit = 'FASEditModeFitTop' },
        @{ VirtualKey = 0x21; ScanCode = 0x49; Place = 'FASEditModePlaceTopRight';    Fit = 'FASEditModeFitTopRight' },
        @{ VirtualKey = 0x25; ScanCode = 0x4B; Place = 'FASEditModePlaceLeft';        Fit = 'FASEditModeFitLeft' },
        @{ VirtualKey = 0x0C; ScanCode = 0x4C; Place = 'FASEditModePlaceCenter';      Fit = 'FASEditModeFitCenter' },
        @{ VirtualKey = 0x27; ScanCode = 0x4D; Place = 'FASEditModePlaceRight';       Fit = 'FASEditModeFitRight' },
        @{ VirtualKey = 0x23; ScanCode = 0x4F; Place = 'FASEditModePlaceBottomLeft';  Fit = 'FASEditModeFitBottomLeft' },
        @{ VirtualKey = 0x28; ScanCode = 0x50; Place = 'FASEditModePlaceBottom';      Fit = 'FASEditModeFitBottom' },
        @{ VirtualKey = 0x22; ScanCode = 0x51; Place = 'FASEditModePlaceBottomRight'; Fit = 'FASEditModeFitBottomRight' }
    )

    foreach ($placement in $keypadPlacements) {
        if (Test-InstrumentaKeypadKey -KeyEvent $KeyEvent -VirtualKey $placement.VirtualKey -ScanCode $placement.ScanCode) {
            if ($KeyEvent.Shift) {
                return $passThrough
            }
            $macro = if ($KeyEvent.Ctrl) { $placement.Fit } else { $placement.Place }
            return New-InstrumentaEditModeAction -Macro $macro -Suppress $true
        }
    }

    if (
        (Test-InstrumentaKeypadKey -KeyEvent $KeyEvent -VirtualKey 0x6B -ScanCode 0x4E) -or
        (Test-InstrumentaKeypadKey -KeyEvent $KeyEvent -VirtualKey 0x6D -ScanCode 0x4A)
    ) {
        if (-not $KeyEvent.Shift) {
            return $passThrough
        }

        $isPlus = ([int]$KeyEvent.VirtualKey -eq 0x6B)
        if ($KeyEvent.Ctrl) {
            $macro = if ($isPlus) { 'FASEditModeBringToFront' } else { 'FASEditModeSendToBack' }
        }
        else {
            $macro = if ($isPlus) { 'FASEditModeBringForward' } else { 'FASEditModeSendBackward' }
        }
        return New-InstrumentaEditModeAction -Macro $macro -Suppress $true
    }

    if (-not [bool]$KeyEvent.Extended) {
        return $passThrough
    }

    $direction = switch ([int]$KeyEvent.VirtualKey) {
        0x25 { 'Left' }
        0x26 { 'Up' }
        0x27 { 'Right' }
        0x28 { 'Down' }
        default { $null }
    }
    if ($null -eq $direction) {
        return $passThrough
    }

    $step = if ($KeyEvent.Shift) { 'Coarse' } else { 'Fine' }
    if ($KeyEvent.Ctrl) {
        $axisAndDirection = switch ($direction) {
            'Left'  { 'WidthDecrease' }
            'Right' { 'WidthIncrease' }
            'Up'    { 'HeightDecrease' }
            'Down'  { 'HeightIncrease' }
        }
        $macro = "FASEditModeResize$axisAndDirection$step"
    }
    else {
        $macro = "FASEditModeNudge$direction$step"
    }

    return New-InstrumentaEditModeAction -Macro $macro -Suppress $true
}
