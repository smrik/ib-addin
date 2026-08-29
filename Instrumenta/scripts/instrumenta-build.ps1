[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('dev', 'prod')]
    [string]$Mode,
    [switch]$ForceClosePowerPoint,
    [switch]$NoLaunchSmoke,
    [switch]$LaunchSmoke,
    [switch]$SyncVba,
    [switch]$UseExistingArtifact
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$RepoRoot = Split-Path -Parent $PSScriptRoot
$SrcDir = Join-Path $RepoRoot 'src'
$ModulesDir = Join-Path $SrcDir 'Modules'
$ClassesDir = Join-Path $SrcDir 'Classes'
$FormsDir = Join-Path $SrcDir 'Forms'
$CustomUiDir = Join-Path $SrcDir 'CustomUI'
$BinDir = Join-Path $RepoRoot 'bin'
$PptmPath = Join-Path $SrcDir 'InstrumentaPowerpointToolbar.pptm'
$CustomUiXmlPath = Join-Path $CustomUiDir 'CustomUI.xml'
$CustomUiOutputPath = Join-Path $CustomUiDir 'CustomUI_output.xml'
$CustomUiJsonPath = Join-Path $CustomUiDir 'CustomUI.json'
$WorkerScriptPath = Join-Path $PSScriptRoot 'instrumenta-build-worker.ps1'

$ProdArtifactPath = Join-Path $BinDir 'InstrumentaPowerpointToolbar.ppam'
$DevArtifactPath = Join-Path $BinDir 'InstrumentaPowerpointToolbarDev.ppam'
$ArtifactPath = if ($Mode -eq 'prod') { $ProdArtifactPath } else { $DevArtifactPath }

$ProdInstallDir = Join-Path $env:APPDATA 'Microsoft\AddIns'
$ProdInstallPath = Join-Path $ProdInstallDir 'InstrumentaPowerpointToolbar.ppam'
$DevInstallPath = Join-Path $ProdInstallDir 'InstrumentaPowerpointToolbarDev.ppam'

$OfficeVersion = '16.0'
$ProdRegistryPath = "HKCU:\Software\Microsoft\Office\$OfficeVersion\powerpoint\AddIns\InstrumentaPowerpointToolbar"
$DevRegistryPath = "HKCU:\Software\Microsoft\Office\$OfficeVersion\powerpoint\AddIns\InstrumentaPowerpointToolbarDev"
$UpSlideUserRegistryPath = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\Finance3point1.UpSlide.PowerPoint'
$UpSlideMachineRegistryPaths = @(
    'HKLM:\Software\Microsoft\Office\PowerPoint\Addins\Finance3point1.UpSlide.PowerPoint',
    'HKLM:\Software\WOW6432Node\Microsoft\Office\PowerPoint\Addins\Finance3point1.UpSlide.PowerPoint'
)

$TempRoot = Join-Path $RepoRoot '.instrumenta-build-temp'
$BuildTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$TempPptmPath = Join-Path $TempRoot "InstrumentaPowerpointToolbar.$Mode.$BuildTimestamp.pptm"
$TempPpamPath = Join-Path $TempRoot "InstrumentaPowerpointToolbar.$Mode.$BuildTimestamp.ppam"
$TempCustomUiXmlPath = Join-Path $TempRoot "CustomUI.$Mode.$BuildTimestamp.xml"

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Open-PowerPointApplication {
    $app = New-Object -ComObject PowerPoint.Application
    try {
        $app.DisplayAlerts = 1
    }
    catch {
    }
    try {
        $app.AutomationSecurity = 3
    }
    catch {
    }
    return $app
}

function Get-RegistryValues {
    param([string]$RegistryPath)

    if (-not (Test-Path -LiteralPath $RegistryPath)) {
        return $null
    }

    return Get-ItemProperty -LiteralPath $RegistryPath
}

function Get-RegistrySummaryLine {
    param([string]$Label, [string]$RegistryPath)

    $values = Get-RegistryValues -RegistryPath $RegistryPath
    if ($null -eq $values) {
        return "${Label}: not registered"
    }

    $pathValue = if ($values.PSObject.Properties.Name -contains 'Path') { $values.Path } else { '<missing Path>' }
    $autoLoadValue = if ($values.PSObject.Properties.Name -contains 'AutoLoad') { $values.AutoLoad } else { '<missing AutoLoad>' }
    return "${Label}: Path=$pathValue ; AutoLoad=$autoLoadValue"
}

function Get-PowerPointOpenPresentationPaths {
    try {
        $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    }
    catch {
        return @()
    }

    $paths = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($presentation in @($app.Presentations)) {
            try {
                $paths.Add($presentation.FullName)
            }
            finally {
                if ($null -ne $presentation) {
                    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($presentation)
                }
            }
        }
    }
    finally {
        if ($null -ne $app) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        }
    }

    return $paths.ToArray()
}

function Get-PowerPointProcesses {
    return @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
}

function Assert-PowerPointClosed {
    $processes = Get-PowerPointProcesses
    if (@($processes).Count -eq 0) {
        return
    }

    if (-not $ForceClosePowerPoint) {
        $presentationLines = Get-PowerPointOpenPresentationPaths
        $messageLines = @(
            'PowerPoint is already running. Close it first, or rerun with -ForceClosePowerPoint.',
            '',
            ('Running POWERPNT processes: ' + (($processes | Select-Object -ExpandProperty Id) -join ', '))
        )

        if (@($presentationLines).Count -gt 0) {
            $messageLines += ''
            $messageLines += 'Open presentations:'
            $messageLines += $presentationLines
        }

        $messageLines += ''
        $messageLines += (Get-RegistrySummaryLine -Label 'Prod registration' -RegistryPath $ProdRegistryPath)
        $messageLines += (Get-RegistrySummaryLine -Label 'Dev registration' -RegistryPath $DevRegistryPath)
        throw ($messageLines -join [Environment]::NewLine)
    }

    Write-Step 'Closing running PowerPoint processes'
    $processes | Stop-Process -Force

    $deadline = (Get-Date).AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 300
        $processes = Get-PowerPointProcesses
    } while (@($processes).Count -gt 0 -and (Get-Date) -lt $deadline)

    if (@($processes).Count -gt 0) {
        throw "Timed out while closing PowerPoint. Remaining PIDs: $((($processes | Select-Object -ExpandProperty Id) -join ', '))"
    }
}

function Assert-RequiredFiles {
    $requiredPaths = @(
        $PptmPath,
        $CustomUiXmlPath,
        $CustomUiJsonPath
    )

    foreach ($path in ($requiredPaths + $WorkerScriptPath)) {
        if (-not (Test-Path -LiteralPath $path)) {
            throw "Required file missing: $path"
        }
    }

    if (@(Get-ChildItem -LiteralPath $ModulesDir -Filter *.bas).Count -eq 0) {
        throw "No VBA modules were found in $ModulesDir"
    }

    if (-not (Test-Path -LiteralPath $ClassesDir)) {
        throw "Required classes folder missing: $ClassesDir"
    }

    if (-not (Test-Path -LiteralPath $FormsDir)) {
        throw "Required forms folder missing: $FormsDir"
    }

    [xml](Get-Content -LiteralPath $CustomUiXmlPath -Raw) | Out-Null
}

function Update-CustomUiOutputMirror {
    Copy-Item -LiteralPath $CustomUiXmlPath -Destination $CustomUiOutputPath -Force
}

function Get-EffectiveCustomUiXmlPath {
    if ($Mode -ne 'dev') {
        return $CustomUiXmlPath
    }

    New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null
    $xml = Get-Content -LiteralPath $CustomUiXmlPath -Raw
    $xml = $xml -replace '<tab id="InstrumentaPowerpointToolbar" label="Instrumenta"', '<tab id="InstrumentaPowerpointToolbarDev" label="Instrumenta Dev"'
    Set-Content -LiteralPath $TempCustomUiXmlPath -Value $xml -Encoding UTF8
    return $TempCustomUiXmlPath
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

function Invoke-PowerPointWorker {
    param(
        [ValidateSet('SyncSource', 'ExportPpam')]
        [string]$Action,
        [hashtable]$Arguments,
        [int]$TimeoutSeconds = 90
    )

    New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null

    $stdoutPath = Join-Path $TempRoot "worker-$Action-stdout-$BuildTimestamp.log"
    $stderrPath = Join-Path $TempRoot "worker-$Action-stderr-$BuildTimestamp.log"
    $enginePath = (Get-Process -Id $PID).Path
    $argumentParts = @('-NoProfile', '-File', (Quote-ProcessArgument $WorkerScriptPath), '-Action', $Action)

    foreach ($key in $Arguments.Keys) {
        $argumentParts += "-$key"
        $argumentParts += (Quote-ProcessArgument ([string]$Arguments[$key]))
    }

    $argumentList = $argumentParts -join ' '

    $process = Start-Process -FilePath $enginePath -ArgumentList $argumentList -PassThru -WindowStyle Hidden -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $succeeded = $false
    try {
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            try { $process.Kill($true) } catch {}
            Get-PowerPointProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
            throw "Timed out during PowerPoint worker step '$Action' after $TimeoutSeconds seconds. Logs: $stdoutPath ; $stderrPath"
        }

        if ($process.ExitCode -ne 0) {
            $stdoutText = if (Test-Path -LiteralPath $stdoutPath) { Get-Content -LiteralPath $stdoutPath -Raw } else { '' }
            $stderrText = if (Test-Path -LiteralPath $stderrPath) { Get-Content -LiteralPath $stderrPath -Raw } else { '' }
            throw "PowerPoint worker step '$Action' failed with exit code $($process.ExitCode).`nSTDOUT:`n$stdoutText`nSTDERR:`n$stderrText"
        }
        $succeeded = $true
    }
    finally {
        if ($succeeded) {
            foreach ($logPath in @($stdoutPath, $stderrPath)) {
                if (Test-Path -LiteralPath $logPath) {
                    Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }
}

function Quote-ProcessArgument {
    param([string]$Value)

    return '"' + ($Value -replace '"', '\"') + '"'
}

function Update-PackageCustomUiXml {
    param(
        [string]$PackagePath,
        [string]$SourceXmlPath
    )

    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry('customUI/customUI.xml')
        if ($null -ne $entry) {
            $entry.Delete()
        }

        $entry = $zip.CreateEntry('customUI/customUI.xml')
        $stream = $entry.Open()
        try {
            $writer = New-Object System.IO.StreamWriter($stream, [System.Text.UTF8Encoding]::new($false))
            try {
                $writer.Write((Get-Content -LiteralPath $SourceXmlPath -Raw))
            }
            finally {
                $writer.Dispose()
            }
        }
        finally {
            $stream.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Get-PackageCustomUiXml {
    param([string]$PackagePath)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        $entry = $zip.GetEntry('customUI/customUI.xml')
        if ($null -eq $entry) {
            throw "customUI/customUI.xml was not found in $PackagePath"
        }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $zip.Dispose()
    }
}

function Assert-BuiltPackage {
    param([string]$PackagePath)

    $packageItem = Get-Item -LiteralPath $PackagePath
    if ($packageItem.Length -lt 200000) {
        throw "The built add-in is unexpectedly small ($($packageItem.Length) bytes): $PackagePath"
    }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        foreach ($requiredEntry in @('ppt/presentation.xml', 'ppt/vbaProject.bin', 'customUI/customUI.xml')) {
            if ($null -eq $zip.GetEntry($requiredEntry)) {
                throw "The built add-in is missing required package entry '$requiredEntry': $PackagePath"
            }
        }
    }
    finally {
        $zip.Dispose()
    }

    $customUiXml = Get-PackageCustomUiXml -PackagePath $PackagePath

    if ($customUiXml -match 'StyleInspector') {
        throw "The built add-in still references StyleInspector. Fix the Ribbon XML before installing."
    }

    if ($customUiXml -notmatch 'FASPlaceMenu') {
        throw "The built add-in does not contain the new FAS Place menu marker."
    }

    if ($customUiXml -notmatch 'RibbonVisibilityMenu') {
        throw "The built add-in does not contain the ribbon declutter menu marker."
    }

    if ($Mode -eq 'dev' -and $customUiXml -notmatch 'Instrumenta Dev') {
        throw "The dev add-in does not contain the Instrumenta Dev ribbon label."
    }
}

function Build-SourcePptm {
    Write-Step 'Refreshing embedded customUI XML in the source PPTM package'
    Update-PackageCustomUiXml -PackagePath $PptmPath -SourceXmlPath $CustomUiXmlPath

    if (-not $SyncVba) {
        Write-Step 'Skipping VBA import into source PPTM. Pass -SyncVba to import tracked modules and classes.'
        return
    }

    Write-Step 'Importing tracked VBA modules and classes into the source PPTM'
    Invoke-PowerPointWorker -Action SyncSource -Arguments @{
        PptmPath   = $PptmPath
        ModulesDir = $ModulesDir
        ClassesDir = $ClassesDir
        FormsDir   = $FormsDir
    } -TimeoutSeconds 120
}

function Export-PpamArtifact {
    if ($UseExistingArtifact) {
        if (-not (Test-Path -LiteralPath $ArtifactPath)) {
            throw "UseExistingArtifact was passed, but the artifact does not exist: $ArtifactPath"
        }

        Write-Step "Using existing $Mode artifact without opening PowerPoint"
        Assert-BuiltPackage -PackagePath $ArtifactPath
        return
    }

    New-Item -ItemType Directory -Path $BinDir -Force | Out-Null
    New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null

    Write-Step "Exporting a fresh $Mode PPAM"
    Invoke-PowerPointWorker -Action ExportPpam -Arguments @{
        TempPptmPath = $PptmPath
        TempPpamPath = $TempPpamPath
    } -TimeoutSeconds 180

    $effectiveCustomUiXmlPath = Get-EffectiveCustomUiXmlPath
    Update-PackageCustomUiXml -PackagePath $TempPpamPath -SourceXmlPath $effectiveCustomUiXmlPath
    Assert-BuiltPackage -PackagePath $TempPpamPath
    Copy-Item -LiteralPath $TempPpamPath -Destination $ArtifactPath -Force
}

function Set-AddInRegistration {
    param(
        [string]$RegistryPath,
        [string]$AddInPath,
        [int]$AutoLoad
    )

    New-Item -Path $RegistryPath -Force | Out-Null
    New-ItemProperty -LiteralPath $RegistryPath -Name Path -Value $AddInPath -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $RegistryPath -Name AutoLoad -Value $AutoLoad -PropertyType DWord -Force | Out-Null
}

function Suspend-InstrumentaAddInAutoload {
    $states = @()

    foreach ($registryPath in @($ProdRegistryPath, $DevRegistryPath)) {
        $values = Get-RegistryValues -RegistryPath $registryPath
        if ($null -eq $values) {
            continue
        }

        $autoLoadValue = if ($values.PSObject.Properties.Name -contains 'AutoLoad') { [int]$values.AutoLoad } else { 0 }
        $states += [pscustomobject]@{
            RegistryPath = $registryPath
            AutoLoad = $autoLoadValue
        }

        if ($autoLoadValue -ne 0) {
            Write-Step "Temporarily disabling Instrumenta autoload at $registryPath"
            New-ItemProperty -LiteralPath $registryPath -Name AutoLoad -Value 0 -PropertyType DWord -Force | Out-Null
        }
    }

    return $states
}

function Restore-InstrumentaAddInAutoload {
    param($States)

    foreach ($state in @($States)) {
        if (Test-Path -LiteralPath $state.RegistryPath) {
            New-ItemProperty -LiteralPath $state.RegistryPath -Name AutoLoad -Value ([int]$state.AutoLoad) -PropertyType DWord -Force | Out-Null
        }
    }
}

function Suspend-UpSlideAutoload {
    $upSlideInstalled = $false
    foreach ($registryPath in $UpSlideMachineRegistryPaths) {
        if (Test-Path -LiteralPath $registryPath) {
            $upSlideInstalled = $true
        }
    }

    if (-not $upSlideInstalled) {
        return $null
    }

    $existingValues = $null
    $existed = Test-Path -LiteralPath $UpSlideUserRegistryPath
    if ($existed) {
        $existingValues = Get-ItemProperty -LiteralPath $UpSlideUserRegistryPath
    }

    Write-Step 'Temporarily disabling UpSlide autoload for this user during the build'
    New-Item -Path $UpSlideUserRegistryPath -Force | Out-Null
    New-ItemProperty -LiteralPath $UpSlideUserRegistryPath -Name LoadBehavior -Value 0 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -LiteralPath $UpSlideUserRegistryPath -Name FriendlyName -Value 'UpSlide for PowerPoint' -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $UpSlideUserRegistryPath -Name Description -Value 'Temporarily disabled during Instrumenta build' -PropertyType String -Force | Out-Null

    return [pscustomobject]@{
        Existed = $existed
        Values = $existingValues
    }
}

function Restore-UpSlideAutoload {
    param($State)

    if ($null -eq $State) {
        return
    }

    if (-not $State.Existed) {
        if (Test-Path -LiteralPath $UpSlideUserRegistryPath) {
            Remove-Item -LiteralPath $UpSlideUserRegistryPath -Recurse -Force
        }
        return
    }

    New-Item -Path $UpSlideUserRegistryPath -Force | Out-Null
    foreach ($property in $State.Values.PSObject.Properties) {
        if ($property.Name -like 'PS*') {
            continue
        }

        New-ItemProperty -LiteralPath $UpSlideUserRegistryPath -Name $property.Name -Value $property.Value -Force | Out-Null
    }
}

function Assert-NoStaleDownloadsPath {
    param([string]$ExpectedPath)

    $prodRegistration = Get-RegistryValues -RegistryPath $ProdRegistryPath
    if ($null -ne $prodRegistration -and $prodRegistration.Path -match '\\Downloads\\InstrumentaPowerpointToolbar\.ppam$') {
        throw "A stale Downloads-based Instrumenta registration is still present: $($prodRegistration.Path). Run the prod build again, then restart PowerPoint."
    }

    if ($ExpectedPath -match '\\Downloads\\InstrumentaPowerpointToolbar\.ppam$') {
        throw "The expected add-in path still points into Downloads: $ExpectedPath"
    }
}

function Get-AddInPathsFromApplication {
    param($Application)

    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($addin in @($Application.AddIns)) {
        try {
            $fullPath = $null
            if ($addin.Path) {
                $fullPath = Join-Path $addin.Path $addin.Name
            }
            elseif ($addin.Name) {
                $fullPath = $addin.Name
            }

            if ($null -ne $fullPath) {
                $paths.Add($fullPath)
            }
        }
        finally {
            Close-ComObject $addin
        }
    }

    return $paths.ToArray()
}

function Invoke-SmokeVerification {
    param(
        [string]$ExpectedPath,
        [string]$RegistryPath,
        [bool]$DevMode
    )

    if ($NoLaunchSmoke -or -not $LaunchSmoke) {
        Write-Step 'Skipping launch smoke. Pass -LaunchSmoke to start PowerPoint for verification.'
        return
    }

    Write-Step 'Launching a fresh PowerPoint smoke instance'

    $app = $null
    try {
        $app = Open-PowerPointApplication

        if ($DevMode) {
            foreach ($addin in @($app.AddIns)) {
                try {
                    $addinPath = if ($addin.Path) { Join-Path $addin.Path $addin.Name } else { $addin.Name }
                    if ($addinPath -eq $ProdInstallPath) {
                        $addin.Loaded = $false
                    }
                }
                finally {
                    Close-ComObject $addin
                }
            }

            $devAddIn = $app.AddIns.Add($ExpectedPath)
            try {
                $devAddIn.Loaded = $true
            }
            finally {
                Close-ComObject $devAddIn
            }
        }

        Start-Sleep -Seconds 2
        $registeredPath = (Get-RegistryValues -RegistryPath $RegistryPath).Path
        if ($registeredPath -ne $ExpectedPath) {
            throw "Smoke verification failed because the registry path did not match the expected path. Expected: $ExpectedPath ; Actual: $registeredPath"
        }

        $loadedPaths = Get-AddInPathsFromApplication -Application $app
        $normalizedLoadedPaths = @(
            foreach ($loadedPath in $loadedPaths) {
                if ([string]::IsNullOrWhiteSpace([IO.Path]::GetExtension($loadedPath))) {
                    "$loadedPath.ppam"
                }
                else {
                    $loadedPath
                }
            }
        )
        if ($normalizedLoadedPaths -contains (Join-Path (Split-Path -Parent $ExpectedPath) (Split-Path -Leaf $ExpectedPath))) {
            # expected path is present
        }
        elseif (-not $DevMode) {
            throw "Smoke verification did not observe the expected add-in path in the PowerPoint AddIns collection. Observed paths: $($loadedPaths -join '; ')"
        }

        if ($normalizedLoadedPaths -match '\\Downloads\\InstrumentaPowerpointToolbar\.ppam$') {
            throw "Smoke verification still found a stale Downloads-based add-in path: $($loadedPaths -join '; ')"
        }
    }
    finally {
        if ($null -ne $app) {
            try { $app.Quit() } catch {}
        }
        Close-ComObject $app
    }
}

function Test-PathInsideTrustedLocation {
    param([string]$PathToCheck)

    $trustedLocationsRoot = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Security\Trusted Locations"
    if (-not (Test-Path -LiteralPath $trustedLocationsRoot)) {
        return $false
    }

    $normalizedPath = $PathToCheck.TrimEnd('\')
    foreach ($locationKey in @(Get-ChildItem -LiteralPath $trustedLocationsRoot -ErrorAction SilentlyContinue)) {
        $values = Get-ItemProperty -LiteralPath $locationKey.PSPath
        $trustedPath = if ($values.PSObject.Properties.Name -contains 'Path') { [string]$values.Path } else { '' }
        if ([string]::IsNullOrWhiteSpace($trustedPath)) {
            continue
        }

        $normalizedTrustedPath = $trustedPath.TrimEnd('\')
        $allowSubFolders = if ($values.PSObject.Properties.Name -contains 'AllowSubFolders') { [int]$values.AllowSubFolders } else { 0 }

        if ($normalizedPath -ieq $normalizedTrustedPath) {
            return $true
        }

        if ($allowSubFolders -ne 0 -and $normalizedPath.StartsWith($normalizedTrustedPath, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Warn-IfPowerPointMacroTrustWillBlock {
    param([string]$InstalledPath)

    $securityPath = "HKCU:\Software\Microsoft\Office\$OfficeVersion\PowerPoint\Security"
    $securityValues = Get-RegistryValues -RegistryPath $securityPath
    if ($null -eq $securityValues -or -not ($securityValues.PSObject.Properties.Name -contains 'VBAWarnings')) {
        return
    }

    $vbaWarnings = [int]$securityValues.VBAWarnings
    if ($vbaWarnings -lt 3) {
        return
    }

    $installedDirectory = Split-Path -Parent $InstalledPath
    if (Test-PathInsideTrustedLocation -PathToCheck $installedDirectory) {
        return
    }

    Write-Warning "PowerPoint macro security is set to block unsigned macros (VBAWarnings=$vbaWarnings), and '$installedDirectory' is not a trusted location."
    Write-Warning "If the add-in loads but commands show 'The macro cannot be found or has been disabled', run: .\scripts\trust-powerpoint-addins.ps1 -Mode Add"
}

function Warn-IfRegisteredPathDrifted {
    param(
        [string]$RegistryPath,
        [string]$ExpectedPath
    )

    $values = Get-RegistryValues -RegistryPath $RegistryPath
    if ($null -eq $values -or -not ($values.PSObject.Properties.Name -contains 'Path')) {
        Write-Warning "PowerPoint add-in registration is missing a Path value at $RegistryPath"
        return
    }

    if ($values.Path -ne $ExpectedPath) {
        Write-Warning "PowerPoint add-in registration path drifted. Expected '$ExpectedPath' but found '$($values.Path)'."
        Write-Warning "Run the build/install script again instead of removing and re-adding the add-in manually."
    }
}

function Install-Artifact {
    if ($Mode -eq 'prod') {
        Write-Step 'Installing prod artifact into the PowerPoint AddIns folder'
        New-Item -ItemType Directory -Path $ProdInstallDir -Force | Out-Null
        Copy-Item -LiteralPath $ArtifactPath -Destination $ProdInstallPath -Force
        Set-AddInRegistration -RegistryPath $ProdRegistryPath -AddInPath $ProdInstallPath -AutoLoad 1
        Assert-NoStaleDownloadsPath -ExpectedPath $ProdInstallPath
        Invoke-SmokeVerification -ExpectedPath $ProdInstallPath -RegistryPath $ProdRegistryPath -DevMode:$false
        Warn-IfRegisteredPathDrifted -RegistryPath $ProdRegistryPath -ExpectedPath $ProdInstallPath
        Warn-IfPowerPointMacroTrustWillBlock -InstalledPath $ProdInstallPath
        return
    }

    Write-Step 'Installing dev artifact into the PowerPoint AddIns folder'
    New-Item -ItemType Directory -Path $ProdInstallDir -Force | Out-Null
    Copy-Item -LiteralPath $ArtifactPath -Destination $DevInstallPath -Force

    Write-Step 'Registering the dev artifact as a separate non-autoload add-in'
    Set-AddInRegistration -RegistryPath $DevRegistryPath -AddInPath $DevInstallPath -AutoLoad 0
    Invoke-SmokeVerification -ExpectedPath $DevInstallPath -RegistryPath $DevRegistryPath -DevMode:$true
    Warn-IfRegisteredPathDrifted -RegistryPath $DevRegistryPath -ExpectedPath $DevInstallPath
    Warn-IfPowerPointMacroTrustWillBlock -InstalledPath $DevInstallPath
}

try {
    Assert-RequiredFiles
    Update-CustomUiOutputMirror
    if ($SyncVba -or -not $UseExistingArtifact -or $LaunchSmoke) {
        Assert-PowerPointClosed
    }

    $suspendedAddInAutoload = @()
    $suspendedUpSlideAutoload = $null
    try {
        if ($SyncVba -or -not $UseExistingArtifact -or $LaunchSmoke) {
            $suspendedAddInAutoload = Suspend-InstrumentaAddInAutoload
            $suspendedUpSlideAutoload = Suspend-UpSlideAutoload
        }

        if (-not $UseExistingArtifact) {
            Build-SourcePptm
        }
        Export-PpamArtifact
    }
    finally {
        Restore-InstrumentaAddInAutoload -States $suspendedAddInAutoload
        Restore-UpSlideAutoload -State $suspendedUpSlideAutoload
    }

    Install-Artifact

    Write-Step 'Build completed successfully'
    Write-Host "Mode: $Mode"
    Write-Host "Artifact: $ArtifactPath"
    if ($Mode -eq 'prod') {
        Write-Host "Installed: $ProdInstallPath"
    }
    else {
        Write-Host "Installed: $DevInstallPath"
    }
}
finally {
    foreach ($tempPath in @($TempPptmPath, $TempPpamPath, $TempCustomUiXmlPath)) {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}
