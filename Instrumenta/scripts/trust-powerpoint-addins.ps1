[CmdletBinding()]
param(
    [ValidateSet('Add', 'Remove', 'Status')]
    [string]$Mode = 'Add'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$OfficeVersion = '16.0'
$TrustedLocationsRoot = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Security\Trusted Locations"
$PolicyTrustedLocationsRoot = "HKCU:\Software\Policies\Microsoft\Office\$OfficeVersion\PowerPoint\Security\Trusted Locations"
$AddInsPath = Join-Path $env:APPDATA 'Microsoft\AddIns'
$TrustedPath = if ($AddInsPath.EndsWith('\')) { $AddInsPath } else { "$AddInsPath\" }
$Description = 'Instrumenta local add-ins'

function Get-TrustedLocationKeys {
    if (-not (Test-Path -LiteralPath $TrustedLocationsRoot)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $TrustedLocationsRoot -ErrorAction SilentlyContinue)
}

function Get-MatchingTrustedLocation {
    foreach ($key in Get-TrustedLocationKeys) {
        $values = Get-ItemProperty -LiteralPath $key.PSPath
        $pathValue = if ($values.PSObject.Properties.Name -contains 'Path') { $values.Path } else { '' }
        if ($pathValue.TrimEnd('\') -ieq $TrustedPath.TrimEnd('\')) {
            return $key
        }
    }

    return $null
}

function Get-NextLocationName {
    $used = New-Object System.Collections.Generic.HashSet[int]
    foreach ($key in Get-TrustedLocationKeys) {
        if ($key.PSChildName -match '^Location(\d+)$') {
            [void]$used.Add([int]$Matches[1])
        }
    }

    for ($i = 1; $i -lt 1000; $i++) {
        if (-not $used.Contains($i)) {
            return "Location$i"
        }
    }

    throw 'No free PowerPoint trusted-location registry slot was found.'
}

function Write-Status {
    $policyLocked = Test-Path -LiteralPath $PolicyTrustedLocationsRoot
    $match = Get-MatchingTrustedLocation

    [pscustomobject]@{
        AddInsPath = $TrustedPath
        Trusted = ($null -ne $match)
        TrustedLocationKey = if ($null -ne $match) { $match.Name } else { $null }
        PolicyTrustedLocationsPresent = $policyLocked
    } | Format-List
}

switch ($Mode) {
    'Status' {
        Write-Status
    }
    'Add' {
        New-Item -ItemType Directory -Path $AddInsPath -Force | Out-Null
        New-Item -Path $TrustedLocationsRoot -Force | Out-Null

        $match = Get-MatchingTrustedLocation
        if ($null -eq $match) {
            $locationName = Get-NextLocationName
            $locationKey = Join-Path $TrustedLocationsRoot $locationName
            New-Item -Path $locationKey -Force | Out-Null
        }
        else {
            $locationKey = $match.PSPath
        }

        New-ItemProperty -LiteralPath $locationKey -Name Path -Value $TrustedPath -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $locationKey -Name Description -Value $Description -PropertyType String -Force | Out-Null
        New-ItemProperty -LiteralPath $locationKey -Name AllowSubFolders -Value 0 -PropertyType DWord -Force | Out-Null
        Write-Status
    }
    'Remove' {
        $match = Get-MatchingTrustedLocation
        if ($null -ne $match) {
            Remove-Item -LiteralPath $match.PSPath -Recurse -Force
        }

        Write-Status
    }
}
