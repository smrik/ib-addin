[CmdletBinding()]
param(
    [ValidateSet("Menu", "List", "Create", "New", "Use", "Current", "Show", "AddSource", "Add", "RemoveSource", "Remove")]
    [string]$Action = "Menu",

    [string]$ProjectName,
    [string]$FilePath,

    [ValidateSet("ledger", "tb", "mapping")]
    [string]$Kind = "ledger",

    [string]$Company,
    [string]$Year,
    [string]$SheetName,
    [int]$HeaderRow = 1,
    [string]$SourceId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$fasRoot = Join-Path $env:APPDATA "FAS_Toolkit"
$projectRoot = Join-Path $fasRoot "Projects"
$projectsPath = Join-Path $projectRoot "projects.json"

function Ensure-Folder {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Read-Projects {
    Ensure-Folder -Path $projectRoot
    if (-not (Test-Path -LiteralPath $projectsPath)) {
        return [pscustomobject]@{
            activeProject = $null
            projects = @()
        }
    }

    $json = Get-Content -LiteralPath $projectsPath -Raw
    if ([string]::IsNullOrWhiteSpace($json)) {
        return [pscustomobject]@{
            activeProject = $null
            projects = @()
        }
    }

    return $json | ConvertFrom-Json
}

function Write-Projects {
    param([Parameter(Mandatory = $true)]$State)
    Ensure-Folder -Path $projectRoot
    $State | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $projectsPath -Encoding UTF8
}

function Get-SafeName {
    param([Parameter(Mandatory = $true)][string]$Value)
    $safe = $Value
    foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
        $safe = $safe.Replace($char, "_")
    }
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "Project"
    }
    return $safe
}

function Get-StableHash {
    param([Parameter(Mandatory = $true)][string]$Value)
    $hash = 0L
    foreach ($char in $Value.ToCharArray()) {
        $hash = (($hash * 131L) + [int][char]$char) % 2147483647L
    }
    return "{0:x8}" -f $hash
}

function Normalize-Action {
    param([Parameter(Mandatory = $true)][string]$Value)
    switch ($Value) {
        "New" { return "Create" }
        "Current" { return "Use" }
        "Add" { return "AddSource" }
        "Remove" { return "RemoveSource" }
        default { return $Value }
    }
}

function Get-Project {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($project in @($State.projects)) {
        if ($project.name -ieq $Name) {
            return $project
        }
    }
    return $null
}

function Get-ActiveProject {
    param([Parameter(Mandatory = $true)]$State)
    if ([string]::IsNullOrWhiteSpace([string]$State.activeProject)) {
        return $null
    }
    return Get-Project -State $State -Name ([string]$State.activeProject)
}

function Read-Prompt {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Default
    )

    if ([string]::IsNullOrWhiteSpace($Default)) {
        return (Read-Host $Prompt).Trim()
    }

    $value = (Read-Host "$Prompt [$Default]").Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }
    return $value
}

function Read-RequiredPrompt {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [string]$Default
    )

    do {
        $value = Read-Prompt -Prompt $Prompt -Default $Default
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
        Write-Host "Please enter a value."
    } while ($true)
}

function Read-IntegerPrompt {
    param(
        [Parameter(Mandatory = $true)][string]$Prompt,
        [int]$Default = 1
    )

    do {
        $value = Read-Prompt -Prompt $Prompt -Default ([string]$Default)
        if ([int]::TryParse($value, [ref]$null)) {
            $parsed = [int]$value
            if ($parsed -ge 1) {
                return $parsed
            }
        }
        Write-Host "Enter a positive number."
    } while ($true)
}

function Read-SourceKind {
    param([string]$Default = "ledger")
    do {
        $value = (Read-Prompt -Prompt "Source type: ledger, tb, or mapping" -Default $Default).ToLowerInvariant()
        if ($value -in @("ledger", "tb", "mapping")) {
            return $value
        }
        Write-Host "Use one of: ledger, tb, mapping."
    } while ($true)
}

function Read-PathPrompt {
    param([string]$Default)
    do {
        $value = Read-RequiredPrompt -Prompt "File path (drag/drop works in most terminals)" -Default $Default
        $value = $value.Trim().Trim('"')
        if (Test-Path -LiteralPath $value) {
            return (Resolve-Path -LiteralPath $value).Path
        }
        Write-Host "File not found: $value"
    } while ($true)
}

function Select-ProjectName {
    param(
        [Parameter(Mandatory = $true)]$State,
        [string]$DefaultProjectName
    )

    $projects = @($State.projects)
    if ($projects.Count -eq 0) {
        throw "No FAS projects configured yet. Create one first."
    }

    if (-not [string]::IsNullOrWhiteSpace($DefaultProjectName)) {
        if ($null -ne (Get-Project -State $State -Name $DefaultProjectName)) {
            return $DefaultProjectName
        }
    }

    Write-Host ""
    Write-Host "Projects:"
    for ($index = 0; $index -lt $projects.Count; $index++) {
        $marker = if ($projects[$index].name -eq $State.activeProject) { "*" } else { " " }
        Write-Host ("  {0}. {1} {2}" -f ($index + 1), $marker, $projects[$index].name)
    }

    do {
        $choice = Read-Prompt -Prompt "Pick project number" -Default "1"
        if ([int]::TryParse($choice, [ref]$null)) {
            $projectIndex = [int]$choice
            if ($projectIndex -ge 1 -and $projectIndex -le $projects.Count) {
                return [string]$projects[$projectIndex - 1].name
            }
        }
        Write-Host "Pick a number from the list."
    } while ($true)
}

function New-ProjectRecord {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -ne (Get-Project -State $State -Name $Name)) {
        throw "Project already exists: $Name"
    }

    $safeName = Get-SafeName -Value $Name
    $folder = Join-Path $projectRoot $safeName
    Ensure-Folder -Path $folder
    Ensure-Folder -Path (Join-Path $folder "raw")
    Ensure-Folder -Path (Join-Path $folder "staging")
    Ensure-Folder -Path (Join-Path $folder "index")

    $newProject = [pscustomobject]@{
        name = $Name
        folder = $folder
        createdAt = (Get-Date).ToString("o")
        sources = @()
    }

    $State.projects = @($State.projects) + $newProject
    $State.activeProject = $Name
    Write-Projects -State $State

    Write-Host "Created and activated project: $Name"
    Write-Host $folder
}

function Add-SourceRecord {
    param(
        [Parameter(Mandatory = $true)]$State,
        [Parameter(Mandatory = $true)][string]$TargetProjectName,
        [Parameter(Mandatory = $true)][string]$TargetFilePath,
        [Parameter(Mandatory = $true)][string]$TargetKind,
        [string]$TargetCompany,
        [string]$TargetYear,
        [string]$TargetSheetName,
        [int]$TargetHeaderRow = 1
    )

    $project = Get-Project -State $State -Name $TargetProjectName
    if ($null -eq $project) {
        throw "Project not found: $TargetProjectName"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $TargetFilePath).Path
    $file = Get-Item -LiteralPath $resolvedPath
    $identityText = "{0}|{1}|{2}|{3}" -f $file.FullName.ToLowerInvariant(), $file.Length, $file.LastWriteTimeUtc.Ticks, $TargetKind
    $id = Get-StableHash -Value $identityText

    $source = [pscustomobject]@{
        id = $id
        kind = $TargetKind
        path = $file.FullName
        originalName = $file.Name
        company = $TargetCompany
        year = $TargetYear
        sheetName = $TargetSheetName
        headerRow = $TargetHeaderRow
        addedAt = (Get-Date).ToString("o")
        fingerprint = $identityText
        fieldMap = [pscustomobject]@{}
    }

    $sources = @($project.sources) | Where-Object { $_.id -ne $id }
    $sources += $source
    $project.sources = $sources
    Write-Projects -State $State

    Write-Host "Added $TargetKind source to $($project.name):"
    Write-Host ("  id:   {0}" -f $id)
    Write-Host ("  path: {0}" -f $file.FullName)
}

function Show-Projects {
    param([Parameter(Mandatory = $true)]$State)
    if (@($State.projects).Count -eq 0) {
        Write-Host "No FAS projects configured yet."
        Write-Host "Run: tools\fas_project.ps1"
        return
    }

    foreach ($project in @($State.projects)) {
        $marker = if ($project.name -eq $State.activeProject) { "*" } else { " " }
        $sourceCount = @($project.sources).Count
        Write-Host ("{0} {1}  ({2} source(s))" -f $marker, $project.name, $sourceCount)
        Write-Host ("   {0}" -f $project.folder)
    }
}

function Show-ProjectDetail {
    param([Parameter(Mandatory = $true)]$Project)
    Write-Host ""
    Write-Host ("Project: {0}" -f $Project.name)
    Write-Host ("Folder:  {0}" -f $Project.folder)
    Write-Host ("Sources: {0}" -f @($Project.sources).Count)
    foreach ($source in @($Project.sources)) {
        Write-Host ("  {0}  {1}  {2}  {3}  {4}" -f $source.id, $source.kind, $source.company, $source.year, $source.originalName)
        Write-Host ("      {0}" -f $source.path)
        if (-not [string]::IsNullOrWhiteSpace([string]$source.sheetName)) {
            Write-Host ("      sheet: {0}, header row: {1}" -f $source.sheetName, $source.headerRow)
        }
    }
}

function Invoke-CreateInteractive {
    param([Parameter(Mandatory = $true)]$State)
    $name = Read-RequiredPrompt -Prompt "Project name" -Default $ProjectName
    New-ProjectRecord -State $State -Name $name
}

function Invoke-UseInteractive {
    param([Parameter(Mandatory = $true)]$State)
    $name = Select-ProjectName -State $State -DefaultProjectName $ProjectName
    $project = Get-Project -State $State -Name $name
    $State.activeProject = $project.name
    Write-Projects -State $State
    Write-Host "Active FAS project: $($project.name)"
}

function Invoke-AddSourceInteractive {
    param([Parameter(Mandatory = $true)]$State)
    $defaultProject = $ProjectName
    if ([string]::IsNullOrWhiteSpace($defaultProject)) {
        $activeProject = Get-ActiveProject -State $State
        if ($null -ne $activeProject) {
            $defaultProject = $activeProject.name
        }
    }

    $name = Select-ProjectName -State $State -DefaultProjectName $defaultProject
    $path = if ([string]::IsNullOrWhiteSpace($FilePath)) { Read-PathPrompt } else { (Resolve-Path -LiteralPath $FilePath).Path }
    $kindValue = if ($PSBoundParameters.ContainsKey("Kind")) { $Kind } else { Read-SourceKind -Default "ledger" }
    $companyValue = Read-Prompt -Prompt "Company/entity label" -Default $Company
    $yearValue = Read-Prompt -Prompt "Year/period label" -Default $Year
    $sheetValue = Read-Prompt -Prompt "Sheet name (blank = first sheet / CSV name)" -Default $SheetName
    $headerValue = if ($PSBoundParameters.ContainsKey("HeaderRow")) { $HeaderRow } else { Read-IntegerPrompt -Prompt "Header row" -Default 1 }

    Add-SourceRecord -State $State -TargetProjectName $name -TargetFilePath $path -TargetKind $kindValue -TargetCompany $companyValue -TargetYear $yearValue -TargetSheetName $sheetValue -TargetHeaderRow $headerValue
}

function Invoke-Menu {
    param([Parameter(Mandatory = $true)]$State)

    do {
        Write-Host ""
        Write-Host "FAS Project Manager"
        Write-Host "1. List projects"
        Write-Host "2. Create project"
        Write-Host "3. Set active project"
        Write-Host "4. Add source file"
        Write-Host "5. Show active project"
        Write-Host "6. Quit"

        $choice = Read-Prompt -Prompt "Choose" -Default "4"
        switch ($choice) {
            "1" { Show-Projects -State $State }
            "2" { Invoke-CreateInteractive -State $State }
            "3" { Invoke-UseInteractive -State $State }
            "4" {
                if (@($State.projects).Count -eq 0) {
                    Write-Host "Create a project first."
                    Invoke-CreateInteractive -State $State
                }
                Invoke-AddSourceInteractive -State $State
            }
            "5" {
                $project = Get-ActiveProject -State $State
                if ($null -eq $project) {
                    Write-Host "No active project."
                } else {
                    Show-ProjectDetail -Project $project
                }
            }
            "6" { return }
            default { Write-Host "Pick 1-6." }
        }
    } while ($true)
}

$state = Read-Projects
$normalizedAction = Normalize-Action -Value $Action

switch ($normalizedAction) {
    "Menu" {
        Invoke-Menu -State $state
    }

    "List" {
        Show-Projects -State $state
    }

    "Create" {
        if ([string]::IsNullOrWhiteSpace($ProjectName)) {
            Invoke-CreateInteractive -State $state
        } else {
            New-ProjectRecord -State $state -Name $ProjectName
        }
    }

    "Use" {
        if ([string]::IsNullOrWhiteSpace($ProjectName)) {
            Invoke-UseInteractive -State $state
        } else {
            $project = Get-Project -State $state -Name $ProjectName
            if ($null -eq $project) {
                throw "Project not found: $ProjectName"
            }
            $state.activeProject = $project.name
            Write-Projects -State $state
            Write-Host "Active FAS project: $($project.name)"
        }
    }

    "Show" {
        if ([string]::IsNullOrWhiteSpace($ProjectName)) {
            $project = Get-ActiveProject -State $state
            if ($null -eq $project) {
                throw "No active FAS project."
            }
        } else {
            $project = Get-Project -State $state -Name $ProjectName
            if ($null -eq $project) {
                throw "Project not found: $ProjectName"
            }
        }
        Show-ProjectDetail -Project $project
    }

    "AddSource" {
        if ([string]::IsNullOrWhiteSpace($ProjectName) -or [string]::IsNullOrWhiteSpace($FilePath)) {
            Invoke-AddSourceInteractive -State $state
        } else {
            Add-SourceRecord -State $state -TargetProjectName $ProjectName -TargetFilePath $FilePath -TargetKind $Kind -TargetCompany $Company -TargetYear $Year -TargetSheetName $SheetName -TargetHeaderRow $HeaderRow
        }
    }

    "RemoveSource" {
        $targetProjectName = $ProjectName
        if ([string]::IsNullOrWhiteSpace($targetProjectName)) {
            $targetProjectName = Select-ProjectName -State $state -DefaultProjectName ([string]$state.activeProject)
        }

        $project = Get-Project -State $state -Name $targetProjectName
        if ($null -eq $project) {
            throw "Project not found: $targetProjectName"
        }

        if ([string]::IsNullOrWhiteSpace($SourceId)) {
            Show-ProjectDetail -Project $project
            $SourceId = Read-RequiredPrompt -Prompt "Source id to remove"
        }

        $before = @($project.sources).Count
        $project.sources = @($project.sources) | Where-Object { $_.id -ne $SourceId }
        if (@($project.sources).Count -eq $before) {
            throw "Source id not found: $SourceId"
        }

        Write-Projects -State $state
        Write-Host "Removed source: $SourceId"
    }
}
