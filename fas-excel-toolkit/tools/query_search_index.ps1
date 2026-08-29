[CmdletBinding()]
param(
    [string]$ProjectName,
    [ValidateSet("ledger", "tb", "mapping")]
    [string]$Kind = "ledger",
    [Parameter(Mandatory = $true)]
    [string]$Query,
    [int]$Limit = 15,
    [string]$DuckDbPath,
    [switch]$AsJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$fasRoot = Join-Path $env:APPDATA "FAS_Toolkit"
$projectsPath = Join-Path $fasRoot "Projects\projects.json"

function Read-State {
    if (-not (Test-Path -LiteralPath $projectsPath)) {
        throw "No FAS projects configured. Create one with tools\fas_project.ps1 first."
    }
    return (Get-Content -LiteralPath $projectsPath -Raw | ConvertFrom-Json)
}

function Get-Project {
    param($State, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        $Name = $State.activeProject
    }
    if ([string]::IsNullOrWhiteSpace($Name)) {
        throw "No active FAS project. Run tools\fas_project.ps1 -Action Use -ProjectName <name>."
    }
    foreach ($project in @($State.projects)) {
        if ($project.name -ieq $Name) {
            return $project
        }
    }
    throw "Project not found: $Name"
}

function Resolve-DuckDb {
    param([string]$ProjectFolder)

    if (-not [string]::IsNullOrWhiteSpace($DuckDbPath)) {
        if (-not (Test-Path -LiteralPath $DuckDbPath)) {
            throw "DuckDB path not found: $DuckDbPath"
        }
        return (Resolve-Path -LiteralPath $DuckDbPath).Path
    }

    if (-not [string]::IsNullOrWhiteSpace($ProjectFolder)) {
        $rememberedPathFile = Join-Path (Join-Path $ProjectFolder "index") "duckdb_path.txt"
        if (Test-Path -LiteralPath $rememberedPathFile) {
            $rememberedPath = (Get-Content -LiteralPath $rememberedPathFile -Raw).Trim()
            if (-not [string]::IsNullOrWhiteSpace($rememberedPath) -and (Test-Path -LiteralPath $rememberedPath)) {
                return (Resolve-Path -LiteralPath $rememberedPath).Path
            }
        }
    }

    $command = Get-Command duckdb -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $candidate = Join-Path $fasRoot "Tools\duckdb.exe"
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    throw "DuckDB CLI was not found. Put duckdb.exe in %APPDATA%\FAS_Toolkit\Tools or add it to PATH."
}

function Split-SearchTokens {
    param([Parameter(Mandatory = $true)][string]$Text)
    return @($Text -split "\s+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Escape-SqlLiteral {
    param([Parameter(Mandatory = $true)][string]$Text)
    return $Text.Replace("'", "''")
}

$state = Read-State
$project = Get-Project -State $state -Name $ProjectName
$duckDb = Resolve-DuckDb -ProjectFolder ([string]$project.folder)
$databasePath = Join-Path (Join-Path $project.folder "index") "fas_index.duckdb"
if (-not (Test-Path -LiteralPath $databasePath)) {
    throw "Search index not found: $databasePath. Run tools\build_search_index.ps1 first."
}

if ($Limit -lt 1) {
    $Limit = 15
}
if ($Limit -gt 100) {
    $Limit = 100
}

$tokens = @(Split-SearchTokens -Text $Query)
if ($tokens.Count -eq 0) {
    throw "Search query is blank."
}

$clauses = New-Object System.Collections.Generic.List[string]
foreach ($token in $tokens) {
    $escaped = Escape-SqlLiteral $token.ToLowerInvariant()
    [void]$clauses.Add("lower(coalesce(raw_text, '') || ' ' || coalesce(account, '') || ' ' || coalesce(description, '') || ' ' || coalesce(classification, '')) LIKE '%$escaped%'")
}
$whereSql = ($clauses -join " AND ")

$sql = @"
SELECT
    company,
    year,
    source_file,
    source_sheet,
    source_row,
    posting_date,
    account,
    account_name,
    description,
    debit,
    credit,
    amount,
    classification,
    substr(raw_text, 1, 220) AS snippet
FROM $Kind
WHERE $whereSql
LIMIT $Limit;
"@

$raw = & $duckDb -json $databasePath -c $sql
if ($LASTEXITCODE -ne 0) {
    throw "DuckDB query failed."
}

if ($AsJson) {
    $raw
    exit 0
}

if ([string]::IsNullOrWhiteSpace(($raw -join ""))) {
    Write-Host "No $Kind matches for '$Query'."
    exit 0
}

$rows = $raw | ConvertFrom-Json
if ($null -eq $rows) {
    Write-Host "No $Kind matches for '$Query'."
    exit 0
}

if ($rows -isnot [System.Array]) {
    $rows = @($rows)
}

Write-Host ("FAS {0} search: {1}" -f $Kind, $Query)
Write-Host ("Project: {0}" -f $project.name)
Write-Host ("-" * 72)

$index = 0
foreach ($row in $rows) {
    $index++
    $where = @(@($row.company, $row.year, $row.source_file, $row.source_sheet, ("r" + $row.source_row)) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $headlineParts = @(@($row.account, $row.account_name, $row.description) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $valueParts = @()
    if (-not [string]::IsNullOrWhiteSpace([string]$row.debit)) { $valueParts += "DB $($row.debit)" }
    if (-not [string]::IsNullOrWhiteSpace([string]$row.credit)) { $valueParts += "CR $($row.credit)" }
    if (-not [string]::IsNullOrWhiteSpace([string]$row.amount)) { $valueParts += "AMT $($row.amount)" }
    if (-not [string]::IsNullOrWhiteSpace([string]$row.classification)) { $valueParts += "Class $($row.classification)" }

    Write-Host ("{0}. {1}" -f $index, ($where -join " | "))
    if ($headlineParts.Count -gt 0) {
        Write-Host ("   {0}" -f ($headlineParts -join " | "))
    }
    if ($valueParts.Count -gt 0) {
        Write-Host ("   {0}" -f ($valueParts -join " | "))
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$row.snippet)) {
        Write-Host ("   {0}" -f $row.snippet)
    }
}
