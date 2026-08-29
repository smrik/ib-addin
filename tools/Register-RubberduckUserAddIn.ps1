[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA 'Rubberduck\2.5.9.6316')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Set-RegistryDefault {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    Set-Item -Path $Path -Value $Value
}

function Set-RegistryValue {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string]$Name,
        [Parameter(Mandatory)] [object]$Value,
        [Microsoft.Win32.RegistryValueKind]$Kind = [Microsoft.Win32.RegistryValueKind]::String
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    if ($Kind -eq [Microsoft.Win32.RegistryValueKind]::DWord) {
        New-ItemProperty -Path $Path -Name $Name -Value ([int]$Value) -PropertyType DWord -Force | Out-Null
        return
    }

    New-ItemProperty -Path $Path -Name $Name -Value ([string]$Value) -PropertyType String -Force | Out-Null
}

function Register-RubberduckComClass {
    param(
        [Parameter(Mandatory)] [string]$ProgId,
        [Parameter(Mandatory)] [string]$ClassName,
        [Parameter(Mandatory)] [string]$Clsid,
        [Parameter(Mandatory)] [string]$AssemblyName,
        [Parameter(Mandatory)] [string]$CodeBase
    )

    $classesRoot = 'HKCU:\Software\Classes'
    $progIdPath = Join-Path $classesRoot $ProgId
    $clsidPath = Join-Path $classesRoot "CLSID\$Clsid"
    $inprocPath = Join-Path $clsidPath 'InprocServer32'
    $versionPath = Join-Path $inprocPath '2.5.9.6316'
    $categoryPath = Join-Path $clsidPath 'Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'

    Set-RegistryDefault $progIdPath $ClassName
    Set-RegistryDefault (Join-Path $progIdPath 'CLSID') $Clsid

    Set-RegistryDefault $clsidPath $ClassName
    Set-RegistryDefault $inprocPath 'mscoree.dll'
    Set-RegistryValue $inprocPath 'ThreadingModel' 'Both'
    Set-RegistryValue $inprocPath 'Class' $ClassName
    Set-RegistryValue $inprocPath 'Assembly' $AssemblyName
    Set-RegistryValue $inprocPath 'RuntimeVersion' 'v4.0.30319'
    Set-RegistryValue $inprocPath 'CodeBase' $CodeBase

    Set-RegistryValue $versionPath 'Class' $ClassName
    Set-RegistryValue $versionPath 'Assembly' $AssemblyName
    Set-RegistryValue $versionPath 'RuntimeVersion' 'v4.0.30319'
    Set-RegistryValue $versionPath 'CodeBase' $CodeBase

    Set-RegistryDefault (Join-Path $clsidPath 'ProgId') $ProgId
    New-Item -Path $categoryPath -Force | Out-Null
}

$assemblyPath = Join-Path $InstallRoot 'Rubberduck.dll'
if (-not (Test-Path -LiteralPath $assemblyPath -PathType Leaf)) {
    throw "Rubberduck.dll not found: $assemblyPath"
}

# Rubberduck's legacy installer is x86/x64-only. Register the extracted payload
# under HKCU so this ARM64 machine does not need an all-users install.
$assemblyName = 'Rubberduck, Version=2.5.9.6316, Culture=neutral, PublicKeyToken=null'
$codeBase = ([Uri]$assemblyPath).AbsoluteUri

Register-RubberduckComClass `
    -ProgId 'Rubberduck.Extension' `
    -ClassName 'Rubberduck._Extension' `
    -Clsid '{69E0F697-43F0-3B33-B105-9B8188A6F040}' `
    -AssemblyName $assemblyName `
    -CodeBase $codeBase

Register-RubberduckComClass `
    -ProgId 'Rubberduck.DockableWindowHost' `
    -ClassName 'Rubberduck.UI._DockableWindowHost' `
    -Clsid '{69E0F699-43F0-3B33-B105-9B8188A6F040}' `
    -AssemblyName $assemblyName `
    -CodeBase $codeBase

foreach ($addInRoot in @(
    'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\Rubberduck.Extension',
    'HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\Rubberduck.Extension'
)) {
    Set-RegistryValue $addInRoot 'LoadBehavior' 3 DWord
    Set-RegistryValue $addInRoot 'FriendlyName' 'Rubberduck'
    Set-RegistryValue $addInRoot 'Description' 'Rubberduck'
}

Write-Host "Registered Rubberduck 2.5.91 for the current user."
Write-Host "Close and reopen Excel/PowerPoint once so VBE loads the add-in."
