[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [switch]$Visible,
    [switch]$ForceCloseExcel,
    [switch]$CleanUnmanagedComponents
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Excel must allow "Trust access to the VBA project object model" for VBProject automation to work.

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$sourceRoot = Join-Path $repoRoot "vba-files"
$modulesPath = Join-Path $sourceRoot "Module"
$classesPath = Join-Path $sourceRoot "Class"
$formsPath = Join-Path $sourceRoot "Form"
$documentsPath = Join-Path $sourceRoot "Class"
# In this layout document modules (ThisWorkbook, SheetN) share the Class folder
# with real class modules. They are code-replaced, never imported as components.
$documentModuleNames = @("ThisWorkbook", "Sheet1")

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

function Test-ComponentFilePair {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$FormFile
    )

    $frxPath = [System.IO.Path]::ChangeExtension($FormFile.FullName, ".frx")
    return (Test-Path -LiteralPath $frxPath)
}

function Get-FormCodeText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FormPath
    )

    $lines = Get-Content -LiteralPath $FormPath
    $optionIndex = [Array]::IndexOf($lines, "Option Explicit")
    if ($optionIndex -lt 0) {
        throw "Form file '$FormPath' does not contain Option Explicit."
    }

    return ($lines[$optionIndex..($lines.Length - 1)] -join [Environment]::NewLine)
}

function Remove-ComponentIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$ComponentName
    )

    try {
        $existing = $VBProject.VBComponents.Item($ComponentName)
    } catch {
        $existing = $null
    }

    if ($null -eq $existing) {
        return
    }

    if ($existing.Type -eq 100) {
        throw "Refusing to remove document module '$ComponentName'."
    }

    $VBProject.VBComponents.Remove($existing)
}

function Remove-UnmanagedComponents {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string[]]$ManagedComponentNames
    )

    $managedLookup = @{}
    foreach ($name in $ManagedComponentNames) {
        $managedLookup[$name.ToLowerInvariant()] = $true
    }

    $componentsToRemove = New-Object System.Collections.Generic.List[object]
    foreach ($component in $VBProject.VBComponents) {
        if ($component.Type -eq 100) {
            continue
        }

        if (-not $managedLookup.ContainsKey($component.Name.ToLowerInvariant())) {
            [void]$componentsToRemove.Add($component)
        }
    }

    foreach ($component in $componentsToRemove) {
        $VBProject.VBComponents.Remove($component)
    }
}

function Repair-KnownGeneratedFormNames {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject
    )

    if ($null -ne (Get-VBComponentIfPresent -VBProject $VBProject -ComponentName "frmFASCommandPalette")) {
        return
    }

    $unnamedForms = @($VBProject.VBComponents | Where-Object {
        $_.Type -eq 3 -and $_.Name -ieq "UserForm1"
    })

    if ($unnamedForms.Count -eq 1 -and
        (Test-Path -LiteralPath (Join-Path $formsPath "frmFASCommandPalette.frm")) -and
        -not (Test-Path -LiteralPath (Join-Path $formsPath "frmFASCommandPalette.frx"))) {
        $unnamedForms[0].Name = "frmFASCommandPalette"
        Write-Host "Repaired generated form name: frmFASCommandPalette"
    }
}

function Set-CodeModuleContent {
    param(
        [Parameter(Mandatory = $true)]
        $CodeModule,

        [Parameter(Mandatory = $true)]
        [string]$CodeText
    )

    if ($CodeModule.CountOfLines -gt 0) {
        $CodeModule.DeleteLines(1, $CodeModule.CountOfLines)
    }

    if ([string]::IsNullOrWhiteSpace($CodeText)) {
        return
    }

    $CodeModule.InsertLines(1, $CodeText)
}

function Get-SourceCodeText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $lines = @(Get-Content -LiteralPath $SourcePath)
    $optionIndex = [Array]::IndexOf($lines, "Option Explicit")
    if ($optionIndex -ge 0) {
        return ($lines[$optionIndex..($lines.Length - 1)] -join [Environment]::NewLine)
    }

    return (($lines | Where-Object { $_ -notmatch '^Attribute ' }) -join [Environment]::NewLine)
}

function Import-FileComponent {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File
    )

    $existing = Get-VBComponentIfPresent -VBProject $VBProject -ComponentName $File.BaseName
    if ($null -ne $existing -and $existing.Type -in @(1, 2)) {
        # Updating the existing code module avoids an Excel ARM/VBE crash caused by Remove().
        Set-CodeModuleContent -CodeModule $existing.CodeModule -CodeText (Get-SourceCodeText -SourcePath $File.FullName)
        return
    }

    if ($null -ne $existing) {
        Remove-ComponentIfPresent -VBProject $VBProject -ComponentName $File.BaseName
    }

    if ($File.Extension -ieq '.frm') {
        [void]$VBProject.VBComponents.Import($File.FullName)
        return
    }

    $componentType = if ($File.Extension -ieq '.cls') { 2 } else { 1 }
    $component = $VBProject.VBComponents.Add($componentType)
    $component.Name = $File.BaseName
    Set-CodeModuleContent -CodeModule $component.CodeModule -CodeText (Get-SourceCodeText -SourcePath $File.FullName)
}

function Add-CommandButton {
    param(
        [Parameter(Mandatory = $true)]
        $Designer,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [string]$Caption,

        [double]$Left,
        [double]$Top,
        [double]$Width,
        [double]$Height
    )

    $control = $Designer.Controls.Add("Forms.CommandButton.1", $Name, $true)
    $control.Caption = $Caption
    $control.Left = $Left
    $control.Top = $Top
    $control.Width = $Width
    $control.Height = $Height
    return $control
}

function Get-VBComponentIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$ComponentName
    )

    try {
        return $VBProject.VBComponents.Item($ComponentName)
    }
    catch {
        return $null
    }
}

function Set-ComPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        $Object,

        [Parameter(Mandatory = $true)]
        [string]$PropertyName,

        [Parameter(Mandatory = $true)]
        $Value
    )

    try {
        $Object.$PropertyName = $Value
        return
    }
    catch {
        try {
            $Object.Properties.Item($PropertyName).Value = $Value
            return
        }
        catch {
            throw "Could not set property '$PropertyName'. $($_.Exception.Message)"
        }
    }
}

function Remove-AllDesignerControls {
    param(
        [Parameter(Mandatory = $true)]
        $Designer
    )

    $controlNames = New-Object System.Collections.Generic.List[string]

    foreach ($control in $Designer.Controls) {
        [void]$controlNames.Add([string]$control.Name)
    }

    foreach ($controlName in $controlNames) {
        $Designer.Controls.Remove($controlName)
    }
}

function New-FASSidebarForm {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$FormPath
    )

    $component = Get-VBComponentIfPresent -VBProject $VBProject -ComponentName "frmFASSidebar"
    if ($null -eq $component) {
        $component = $VBProject.VBComponents.Add(3)
        $component.Name = "frmFASSidebar"
    }
    elseif ($component.Type -ne 3) {
        throw "Existing component 'frmFASSidebar' is not a UserForm."
    }

    $designer = $component.Designer
    Remove-AllDesignerControls -Designer $designer
    Set-ComPropertyValue -Object $designer -PropertyName "Caption" -Value "FAS Navigator"
    try {
        Set-ComPropertyValue -Object $designer -PropertyName "StartUpPosition" -Value 0
    }
    catch {
        # Some automation contexts do not expose StartUpPosition during design-time updates.
    }
    try {
        Set-ComPropertyValue -Object $designer -PropertyName "Width" -Value 240
        Set-ComPropertyValue -Object $designer -PropertyName "Height" -Value 330
    }
    catch {
        # Some Excel/MSForms automation environments expose limited form sizing support.
        # The default form size is still usable for the generated sidebar layout.
    }

    $txtSearch = $designer.Controls.Add("Forms.TextBox.1", "txtSearch", $true)
    Set-ComPropertyValue -Object $txtSearch -PropertyName "Left" -Value 12
    Set-ComPropertyValue -Object $txtSearch -PropertyName "Top" -Value 12
    Set-ComPropertyValue -Object $txtSearch -PropertyName "Width" -Value 156
    Set-ComPropertyValue -Object $txtSearch -PropertyName "Height" -Value 18

    $lstSheets = $designer.Controls.Add("Forms.ListBox.1", "lstSheets", $true)
    Set-ComPropertyValue -Object $lstSheets -PropertyName "Left" -Value 12
    Set-ComPropertyValue -Object $lstSheets -PropertyName "Top" -Value 42
    Set-ComPropertyValue -Object $lstSheets -PropertyName "Width" -Value 204
    Set-ComPropertyValue -Object $lstSheets -PropertyName "Height" -Value 210

    $cboGroup = $designer.Controls.Add("Forms.ComboBox.1", "cboGroup", $true)
    Set-ComPropertyValue -Object $cboGroup -PropertyName "Left" -Value 12
    Set-ComPropertyValue -Object $cboGroup -PropertyName "Top" -Value 258
    Set-ComPropertyValue -Object $cboGroup -PropertyName "Width" -Value 84
    Set-ComPropertyValue -Object $cboGroup -PropertyName "Height" -Value 18

    $txtSlot = $designer.Controls.Add("Forms.TextBox.1", "txtSlot", $true)
    Set-ComPropertyValue -Object $txtSlot -PropertyName "Left" -Value 102
    Set-ComPropertyValue -Object $txtSlot -PropertyName "Top" -Value 258
    Set-ComPropertyValue -Object $txtSlot -PropertyName "Width" -Value 30
    Set-ComPropertyValue -Object $txtSlot -PropertyName "Height" -Value 18

    [void](Add-CommandButton -Designer $designer -Name "cmdAssign" -Caption "Set Slot" -Left 138 -Top 258 -Width 78 -Height 24)
    [void](Add-CommandButton -Designer $designer -Name "cmdClearSlot" -Caption "Clear Slot" -Left 12 -Top 288 -Width 96 -Height 24)
    [void](Add-CommandButton -Designer $designer -Name "cmdRefresh" -Caption "Refresh" -Left 120 -Top 288 -Width 48 -Height 24)
    [void](Add-CommandButton -Designer $designer -Name "cmdClose" -Caption "Close" -Left 174 -Top 288 -Width 42 -Height 24)

    Set-CodeModuleContent -CodeModule $component.CodeModule -CodeText (Get-FormCodeText -FormPath $FormPath)
}

function New-FASCommandPaletteForm {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$FormPath
    )

    $component = Get-VBComponentIfPresent -VBProject $VBProject -ComponentName "frmFASCommandPalette"
    if ($null -eq $component) {
        $component = $VBProject.VBComponents.Add(3)
    }
    elseif ($component.Type -ne 3) {
        throw "Existing component 'frmFASCommandPalette' is not a UserForm."
    }

    $designer = $component.Designer
    Remove-AllDesignerControls -Designer $designer
    Set-ComPropertyValue -Object $designer -PropertyName "Caption" -Value "FAS Command Palette"
    try {
        Set-ComPropertyValue -Object $designer -PropertyName "StartUpPosition" -Value 1
        Set-ComPropertyValue -Object $designer -PropertyName "Width" -Value 520
        Set-ComPropertyValue -Object $designer -PropertyName "Height" -Value 290
    }
    catch {
        # Form sizing can be unavailable in restricted VBIDE automation contexts.
    }

    $txtCommand = $designer.Controls.Add("Forms.TextBox.1", "txtCommand", $true)
    Set-ComPropertyValue -Object $txtCommand -PropertyName "Left" -Value 12
    Set-ComPropertyValue -Object $txtCommand -PropertyName "Top" -Value 12
    Set-ComPropertyValue -Object $txtCommand -PropertyName "Width" -Value 484
    Set-ComPropertyValue -Object $txtCommand -PropertyName "Height" -Value 22

    $lstCommands = $designer.Controls.Add("Forms.ListBox.1", "lstCommands", $true)
    Set-ComPropertyValue -Object $lstCommands -PropertyName "Left" -Value 12
    Set-ComPropertyValue -Object $lstCommands -PropertyName "Top" -Value 42
    Set-ComPropertyValue -Object $lstCommands -PropertyName "Width" -Value 484
    Set-ComPropertyValue -Object $lstCommands -PropertyName "Height" -Value 210

    $lblEmpty = $designer.Controls.Add("Forms.Label.1", "lblEmpty", $true)
    Set-ComPropertyValue -Object $lblEmpty -PropertyName "Left" -Value 18
    Set-ComPropertyValue -Object $lblEmpty -PropertyName "Top" -Value 52
    Set-ComPropertyValue -Object $lblEmpty -PropertyName "Width" -Value 360
    Set-ComPropertyValue -Object $lblEmpty -PropertyName "Height" -Value 18
    Set-ComPropertyValue -Object $lblEmpty -PropertyName "Caption" -Value "No commands match this search."
    Set-ComPropertyValue -Object $lblEmpty -PropertyName "Visible" -Value $false

    Set-CodeModuleContent -CodeModule $component.CodeModule -CodeText (Get-FormCodeText -FormPath $FormPath)
    $component.Name = "frmFASCommandPalette"
    if ($component.Name -ne "frmFASCommandPalette") {
        throw "Could not persist the generated command palette form name."
    }
}

function Import-OrGenerateForm {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$FormFile
    )

    if (Test-ComponentFilePair -FormFile $FormFile) {
        $existing = Get-VBComponentIfPresent -VBProject $VBProject -ComponentName $FormFile.BaseName
        if ($null -ne $existing -and $existing.Type -eq 3) {
            Set-CodeModuleContent -CodeModule $existing.CodeModule -CodeText (Get-FormCodeText -FormPath $FormFile.FullName)
        }
        else {
            Import-FileComponent -VBProject $VBProject -File $FormFile
        }
        return
    }

    if ($FormFile.BaseName -ieq "frmFASSidebar") {
        New-FASSidebarForm -VBProject $VBProject -FormPath $FormFile.FullName
        return
    }

    if ($FormFile.BaseName -ieq "frmFASCommandPalette") {
        throw "Command palette form is missing frmFASCommandPalette.frx. Run tools\bootstrap_command_palette_form.ps1 to recreate the tracked form pair."
    }

    throw "Form '$($FormFile.Name)' does not have a matching .frx file."
}

function Set-ThisWorkbookCode {
    param(
        [Parameter(Mandatory = $true)]
        $VBProject,

        [Parameter(Mandatory = $true)]
        [string]$CodePath
    )

    $documentComponent = $VBProject.VBComponents.Item("ThisWorkbook")
    $codeText = Get-Content -LiteralPath $CodePath -Raw
    Set-CodeModuleContent -CodeModule $documentComponent.CodeModule -CodeText $codeText
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "Workbook path not found: $WorkbookPath"
}

Stop-ExcelIfRequested -ForceClose:$ForceCloseExcel

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$Visible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    # This Office build returns a null VBProject while ForceDisable (3) is active.
    # The importer only opens the local build artifact and requires trusted VBOM access.
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path, 0, $false)
    Write-Host "Opened Excel build artifact."
    $vbProject = $workbook.VBProject
    Write-Host "Connected to VBA project."
    Write-Host "Preparing existing VBA components."
    Repair-KnownGeneratedFormNames -VBProject $vbProject
    $managedComponentNames = @(
        (Get-ChildItem -LiteralPath $modulesPath -Filter *.bas | ForEach-Object { $_.BaseName }),
        (Get-ChildItem -LiteralPath $classesPath -Filter *.cls |
            Where-Object { $documentModuleNames -notcontains $_.BaseName } |
            ForEach-Object { $_.BaseName }),
        (Get-ChildItem -LiteralPath $formsPath -Filter *.frm | ForEach-Object { $_.BaseName })
    ) | Where-Object { $_ }

    if ($CleanUnmanagedComponents) {
        Write-Host "Removing unmanaged VBA components."
        Remove-UnmanagedComponents -VBProject $vbProject -ManagedComponentNames $managedComponentNames
    }
    else {
        # Removing components is unreliable in some Office/VBE builds. Keep
        # unmanaged components by default; managed components are still synced
        # by name and new source components are added below.
        Write-Host "Keeping unmanaged VBA components for host stability."
    }

    foreach ($file in Get-ChildItem -LiteralPath $modulesPath -Filter *.bas | Sort-Object Name) {
        Write-Host "Importing module $($file.Name)"
        Import-FileComponent -VBProject $vbProject -File $file
    }

    foreach ($file in Get-ChildItem -LiteralPath $classesPath -Filter *.cls |
        Where-Object { $documentModuleNames -notcontains $_.BaseName } | Sort-Object Name) {
        Import-FileComponent -VBProject $vbProject -File $file
    }

    foreach ($file in Get-ChildItem -LiteralPath $formsPath -Filter *.frm | Sort-Object Name) {
        Import-OrGenerateForm -VBProject $vbProject -FormFile $file
    }

    Set-ThisWorkbookCode -VBProject $vbProject -CodePath (Join-Path $documentsPath "ThisWorkbook.cls")

    $workbook.Save()
}
finally {
    if ($null -ne $workbook) {
        $workbook.Close($true)
    }

    if ($null -ne $excel) {
        $excel.Quit()
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
