[CmdletBinding()]
param(
    [switch]$InstallExtensions,
    [switch]$InstallNodeDependencies
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot

function Write-Check {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][bool]$Passed,
        [Parameter(Mandatory = $true)][string]$Detail
    )

    $status = if ($Passed) { "OK" } else { "WARN" }
    Write-Host ("[{0}] {1}: {2}" -f $status, $Name, $Detail)
}

function Get-CommandPath([string]$Name) {
    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if ($null -eq $command) { return $null }
    return $command.Source
}

function Get-CommandDetail([string]$Path, [string]$MissingDetail) {
    if ($null -ne $Path) { return $Path }
    return $MissingDetail
}

Write-Host "FAS Add-In development setup"
Write-Host "Workspace: $RepoRoot"

$PowerShellPath = Get-CommandPath "powershell.exe"
$NodePath = Get-CommandPath "node.exe"
$NpmPath = Get-CommandPath "npm.cmd"
$PythonPath = Get-CommandPath "python.exe"
$CodePath = Get-CommandPath "code.cmd"

Write-Check -Name "Windows PowerShell" -Passed ($null -ne $PowerShellPath) -Detail (Get-CommandDetail $PowerShellPath "not found")
Write-Check -Name "Node.js" -Passed ($null -ne $NodePath) -Detail (Get-CommandDetail $NodePath "not found; needed for XVBA dependencies")
Write-Check -Name "npm" -Passed ($null -ne $NpmPath) -Detail (Get-CommandDetail $NpmPath "not found; needed for XVBA dependencies")
Write-Check -Name "Python" -Passed ($null -ne $PythonPath) -Detail (Get-CommandDetail $PythonPath "not found; needed for XVBA source checks")
Write-Check -Name "VS Code CLI" -Passed ($null -ne $CodePath) -Detail (Get-CommandDetail $CodePath "not found; needed only with -InstallExtensions")

$ExcelCandidates = @()
$PowerPointCandidates = @()
foreach ($programFilesRoot in @(${env:ProgramFiles}, ${env:ProgramFiles(x86)})) {
    if ([string]::IsNullOrWhiteSpace([string]$programFilesRoot)) { continue }
    $ExcelCandidates += Join-Path $programFilesRoot "Microsoft Office\root\Office16\EXCEL.EXE"
    $PowerPointCandidates += Join-Path $programFilesRoot "Microsoft Office\root\Office16\POWERPNT.EXE"
}

$ExcelPath = $ExcelCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
$PowerPointPath = $PowerPointCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
Write-Check -Name "Desktop Excel" -Passed ($null -ne $ExcelPath) -Detail (Get-CommandDetail $ExcelPath "not found; Excel COM build/smoke tasks will not run")
Write-Check -Name "Desktop PowerPoint" -Passed ($null -ne $PowerPointPath) -Detail (Get-CommandDetail $PowerPointPath "not found; PowerPoint build/smoke tasks will not run")

if ($InstallExtensions) {
    if ($null -eq $CodePath) {
        throw "VS Code CLI 'code' was not found. Install VS Code or remove -InstallExtensions."
    }

    $extensionManifest = Join-Path $RepoRoot ".vscode\extensions.json"
    $extensions = @((Get-Content -LiteralPath $extensionManifest -Raw | ConvertFrom-Json).recommendations)
    foreach ($extension in $extensions) {
        Write-Host "Installing VS Code extension: $extension"
        & $CodePath --install-extension $extension --force
        if ($LASTEXITCODE -ne 0) {
            throw "VS Code failed to install extension: $extension"
        }
    }
}

if ($InstallNodeDependencies) {
    if ($null -eq $NpmPath) {
        throw "npm was not found. Install Node.js or remove -InstallNodeDependencies."
    }

    $NodeProjects = @(
        (Join-Path $RepoRoot "patriks-tab")
    )

    foreach ($project in $NodeProjects) {
        if (-not (Test-Path -LiteralPath (Join-Path $project "package.json"))) {
            continue
        }
        Write-Host "Installing npm dependencies in $project"
        Push-Location $project
        try {
            & $NpmPath install
            if ($LASTEXITCODE -ne 0) {
                throw "npm install failed in $project"
            }
        }
        finally {
            Pop-Location
        }
    }
}

Write-Host "Setup checks complete. Open the workspace in VS Code and run 'Run source checks'."
