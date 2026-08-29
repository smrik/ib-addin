[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('SyncSource', 'ExportPpam')]
    [string]$Action,
    [string]$PptmPath,
    [string]$ModulesDir,
    [string]$ClassesDir,
    [string]$FormsDir,
    [string]$TempPptmPath,
    [string]$TempPpamPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$PpSaveAsOpenXMLAddin = 30
$VbextCtClassModule = 2
$VbextCtMsForm = 3

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

function Invoke-ComOperationWithRetry {
    param(
        [string]$Description,
        [scriptblock]$Operation,
        [int]$TimeoutSeconds = 90
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $attempt = 0
    $lastError = $null

    do {
        $attempt += 1
        try {
            return & $Operation
        }
        catch {
            $lastError = $_
            Write-Output "$Description attempt $attempt failed: $($_.Exception.Message)"
            Start-Sleep -Seconds 2
        }
    } while ((Get-Date) -lt $deadline)

    throw $lastError
}

function Sync-VbaComponents {
    param(
        $VBComponents,
        [string]$SourceFolder,
        [string[]]$Extensions,
        [string[]]$SkipComponentNames = @()
    )

    foreach ($extension in $Extensions) {
        foreach ($file in Get-ChildItem -LiteralPath $SourceFolder -Filter "*$extension" | Sort-Object Name) {
            $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            if ($SkipComponentNames -contains $componentName) {
                Write-Output "Skipping component $componentName from $($file.FullName)"
                continue
            }

            Write-Output "Syncing component $componentName from $($file.FullName)"
            $existingComponent = $null
            try {
                $existingComponent = $VBComponents.Item($componentName)
            }
            catch {
            }

            if ($null -ne $existingComponent) {
                $VBComponents.Remove($existingComponent)
                Close-ComObject $existingComponent
            }

            [void]$VBComponents.Import($file.FullName)
        }
    }
}

function Get-VbaCodeTextForComponent {
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path
    $codeLines = New-Object System.Collections.Generic.List[string]

    foreach ($line in $lines) {
        if ($line -match '^\s*Attribute\s+') {
            continue
        }
        if ($line -match '^\s*VERSION\s+') {
            continue
        }
        if ($line -match '^\s*BEGIN\s*$') {
            continue
        }
        if ($line -match '^\s*END\s*$') {
            continue
        }
        if ($line -match '^\s*MultiUse\s*=') {
            continue
        }

        $codeLines.Add($line)
    }

    return ($codeLines -join [Environment]::NewLine)
}

function Sync-VbaClassComponents {
    param(
        $VBComponents,
        [string]$SourceFolder,
        [string[]]$SkipComponentNames = @()
    )

    foreach ($file in Get-ChildItem -LiteralPath $SourceFolder -Filter '*.cls' | Sort-Object Name) {
        $componentName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        if ($SkipComponentNames -contains $componentName) {
            Write-Output "Skipping class component $componentName from $($file.FullName)"
            continue
        }

        Write-Output "Syncing class component $componentName from $($file.FullName)"
        $existingComponent = $null
        try {
            $existingComponent = $VBComponents.Item($componentName)
        }
        catch {
        }

        if ($null -ne $existingComponent) {
            $VBComponents.Remove($existingComponent)
            Close-ComObject $existingComponent
        }

        $newComponent = $VBComponents.Add($VbextCtClassModule)
        $newComponent.Name = $componentName
        $codeText = Get-VbaCodeTextForComponent -Path $file.FullName
        if (-not [string]::IsNullOrWhiteSpace($codeText)) {
            $newComponent.CodeModule.AddFromString($codeText)
        }
        Close-ComObject $newComponent
    }
}

function Add-FasEditModeControl {
    param(
        $Designer,
        [string]$ProgId,
        [string]$Name,
        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height,
        [string]$Caption = ''
    )

    Write-Host "Adding FAS Edit Mode control $Name ($ProgId)"
    if ($null -eq $Designer) {
        throw "Cannot add FAS Edit Mode control '$Name': the UserForm designer is null."
    }

    try {
        $control = $Designer.Controls.Add($ProgId, $Name, $true)
        if ($null -eq $control) {
            throw "Controls.Add returned null."
        }

        $control.Left = $Left
        $control.Top = $Top
        $control.Width = $Width
        $control.Height = $Height
        if (-not [string]::IsNullOrEmpty($Caption)) {
            $control.Caption = $Caption
        }
    }
    catch {
        throw "Failed to add FAS Edit Mode control '$Name' ($ProgId): $($_.Exception.Message)"
    }

    return $control
}

function Sync-FasEditModeForm {
    param(
        $VBComponents,
        [string]$FormsFolder
    )

    $componentName = 'FASEditModeForm'
    $sourcePath = Join-Path $FormsFolder "$componentName.frm"
    $binaryPath = Join-Path $FormsFolder "$componentName.frx"
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Required FAS Edit Mode form source is missing: $sourcePath"
    }
    if (-not (Test-Path -LiteralPath $binaryPath)) {
        throw "Required FAS Edit Mode form binary is missing: $binaryPath. Run scripts\bootstrap-fas-edit-mode-form.ps1."
    }

    Write-Output "Syncing generated form $componentName from $sourcePath"
    $existingComponent = $null
    try {
        $existingComponent = $VBComponents.Item($componentName)
    }
    catch {
    }

    if ($null -ne $existingComponent) {
        $VBComponents.Remove($existingComponent)
        Close-ComObject $existingComponent
    }

    $importedComponent = $null
    try {
        $importedComponent = $VBComponents.Import($sourcePath)
        if ($null -eq $importedComponent) {
            throw "VBComponents.Import returned null for $sourcePath."
        }
        if ($importedComponent.Name -ne $componentName) {
            $importedComponent.Name = $componentName
        }
        Write-Output "Imported tracked form $componentName"
        return
    }
    finally {
        Close-ComObject $importedComponent
    }

    <#
    $newComponent = $VBComponents.Add($VbextCtMsForm)
    $designer = $null
    try {
        $newComponent.Name = $componentName
        $designer = $newComponent.Designer
        $designer.Caption = 'FAS Edit Mode'
        try {
            $designer.Width = 252
            $designer.Height = 292
        }
        catch {
            Write-Output 'PowerPoint VBIDE did not expose UserForm design-time dimensions; using the default form size.'
        }
        try {
            $designer.StartUpPosition = 0
        }
        catch {
            Write-Output 'PowerPoint VBIDE did not expose UserForm StartUpPosition; using the default position.'
        }

        $modeLabel = Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblMode' -Left 12 -Top 10 -Width 218 -Height 18 -Caption 'EDIT MODE ACTIVE'
        $modeLabel.Font.Bold = $true
        $modeLabel.ForeColor = 0x008000

        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblZone' -Left 12 -Top 32 -Width 218 -Height 16 -Caption 'Zone: Content')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblSelection' -Left 12 -Top 50 -Width 218 -Height 16 -Caption 'No shape selected')
        $statusLabel = Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblStatus' -Left 12 -Top 68 -Width 218 -Height 28 -Caption 'Ready'
        $statusLabel.WordWrap = $true

        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblColumnsTitle' -Left 12 -Top 104 -Width 48 -Height 16 -Caption 'Columns')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.ComboBox.1' -Name 'cboColumns' -Left 62 -Top 101 -Width 48 -Height 18)
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblRowsTitle' -Left 124 -Top 104 -Width 34 -Height 16 -Caption 'Rows')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.ComboBox.1' -Name 'cboRows' -Left 166 -Top 101 -Width 48 -Height 18)

        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblExactColumnTitle' -Left 12 -Top 130 -Width 72 -Height 16 -Caption 'Exact column')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.ComboBox.1' -Name 'cboExactColumn' -Left 86 -Top 127 -Width 128 -Height 18)
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblExactRowTitle' -Left 12 -Top 156 -Width 72 -Height 16 -Caption 'Exact row')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.ComboBox.1' -Name 'cboExactRow' -Left 86 -Top 153 -Width 128 -Height 18)

        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.CommandButton.1' -Name 'cmdZoneContent' -Left 12 -Top 181 -Width 66 -Height 22 -Caption 'Content')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.CommandButton.1' -Name 'cmdZoneTitle' -Left 84 -Top 181 -Width 58 -Height 22 -Caption 'Title')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.CommandButton.1' -Name 'cmdZoneSubtitle' -Left 148 -Top 181 -Width 66 -Height 22 -Caption 'Subtitle')

        $keysLabel = Add-FasEditModeControl -Designer $designer -ProgId 'Forms.Label.1' -Name 'lblKeys' -Left 12 -Top 210 -Width 218 -Height 34 -Caption "Numpad 1-9: place   Ctrl: fit`r`nArrows: move   Ctrl+arrows: resize   Esc: exit"
        $keysLabel.WordWrap = $true
        $keysLabel.Font.Size = 8

        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.CommandButton.1' -Name 'cmdClearExact' -Left 12 -Top 250 -Width 112 -Height 24 -Caption 'Use keypad anchors')
        [void](Add-FasEditModeControl -Designer $designer -ProgId 'Forms.CommandButton.1' -Name 'cmdExit' -Left 148 -Top 250 -Width 66 -Height 24 -Caption 'Exit')

        $codeText = Get-VbaCodeTextForComponent -Path $sourcePath
        if (-not [string]::IsNullOrWhiteSpace($codeText)) {
            $newComponent.CodeModule.AddFromString($codeText)
        }
    }
    finally {
        if ($null -ne $designer) {
            Close-ComObject $designer
        }
        Close-ComObject $newComponent
    }
    #>
}

function Remove-VbaComponents {
    param(
        $VBComponents,
        [string[]]$ComponentNames
    )

    foreach ($componentName in $ComponentNames) {
        $existingComponent = $null
        try {
            $existingComponent = $VBComponents.Item($componentName)
        }
        catch {
        }

        if ($null -ne $existingComponent) {
            Write-Output "Removing component $componentName"
            $VBComponents.Remove($existingComponent)
            Close-ComObject $existingComponent
        }
    }
}

switch ($Action) {
    'SyncSource' {
        $app = $null
        $presentation = $null
        try {
            Write-Output 'Starting PowerPoint for SyncSource'
            $app = Open-PowerPointApplication
            Write-Output "Opening $PptmPath"
            $presentation = Invoke-ComOperationWithRetry -Description 'Opening source PPTM' -Operation { $app.Presentations.Open($PptmPath, $false, $false, $false) } -TimeoutSeconds 90
            Write-Output 'Accessing VBProject'
            $vbComponents = $presentation.VBProject.VBComponents
            Remove-VbaComponents -VBComponents $vbComponents -ComponentNames @('RibbonEventsClass')
            Sync-VbaComponents -VBComponents $vbComponents -SourceFolder $ModulesDir -Extensions @('.bas')
            Sync-VbaClassComponents -VBComponents $vbComponents -SourceFolder $ClassesDir -SkipComponentNames @('RibbonEventsClass')
            Sync-FasEditModeForm -VBComponents $vbComponents -FormsFolder $FormsDir
            Write-Output "Skipping other form sync from $FormsDir because tracked .frm files are code-only exports."
            Write-Output 'Saving source PPTM'
            Invoke-ComOperationWithRetry -Description 'Saving source PPTM' -Operation { $presentation.Save() } -TimeoutSeconds 90
        }
        finally {
            if ($null -ne $presentation) {
                try { $presentation.Close() } catch {}
            }
            if ($null -ne $app) {
                try { $app.Quit() } catch {}
            }
            Close-ComObject $presentation
            Close-ComObject $app
        }
    }
    'ExportPpam' {
        $app = $null
        $presentation = $null
        try {
            Write-Output 'Starting PowerPoint for ExportPpam'
            $app = Open-PowerPointApplication
            Write-Output "Opening $TempPptmPath"
            $presentation = Invoke-ComOperationWithRetry -Description 'Opening source PPTM for export' -Operation { $app.Presentations.Open($TempPptmPath, $false, $false, $true) } -TimeoutSeconds 90
            Write-Output "Saving add-in $TempPpamPath"
            Invoke-ComOperationWithRetry -Description 'Saving add-in' -Operation { $presentation.SaveAs($TempPpamPath, $PpSaveAsOpenXMLAddin) } -TimeoutSeconds 120
            if (-not (Test-Path -LiteralPath $TempPpamPath)) {
                $legacyAddInPath = "$TempPpamPath.ppa"
                if (Test-Path -LiteralPath $legacyAddInPath) {
                    throw "PowerPoint saved a legacy .ppa add-in instead of a .ppam add-in: $legacyAddInPath"
                }
                throw "PowerPoint SaveAs returned but did not create $TempPpamPath"
            }
            $savedAddIn = Get-Item -LiteralPath $TempPpamPath
            if ($savedAddIn.Length -lt 200000) {
                throw "PowerPoint SaveAs created an unexpectedly small add-in ($($savedAddIn.Length) bytes): $TempPpamPath"
            }
        }
        finally {
            if ($null -ne $presentation) {
                try { $presentation.Close() } catch {}
            }
            if ($null -ne $app) {
                try { $app.Quit() } catch {}
            }
            Close-ComObject $presentation
            Close-ComObject $app
        }
    }
}
