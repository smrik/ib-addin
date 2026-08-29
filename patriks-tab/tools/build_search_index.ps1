[CmdletBinding()]
param(
    [string]$ProjectName,
    [ValidateSet("ledger", "tb", "mapping", "all")]
    [string]$Kind = "all",
    [string]$DuckDbPath,
    [switch]$KeepExcelVisible
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$fasRoot = Join-Path $env:APPDATA "FAS_Toolkit"
$projectsPath = Join-Path $fasRoot "Projects\projects.json"
$canonicalFields = @(
    "kind", "project", "company", "year", "source_id", "source_file", "source_path",
    "source_sheet", "source_row", "posting_date", "document_date", "account",
    "account_original", "account_name", "description", "debit", "credit", "amount",
    "classification", "class_id", "counterparty", "document_number", "raw_text"
)

$fieldSynonyms = @{
    posting_date = @("datum knjizenja", "datum knjiženja", "posting date", "book date", "datum")
    document_date = @("datum obdobja", "document date", "invoice date", "datum dokumenta")
    account = @("konto", "account", "account no", "account number", "gl account", "kontonr")
    account_original = @("original konto", "original account", "izvorni konto")
    account_name = @("naziv konta", "account name", "naziv", "konto naziv")
    description = @("opis", "description", "text", "memo", "narrative", "knjizba", "knjižba")
    debit = @("debet", "debit", "duguje", "promet db", "db")
    credit = @("kredit", "credit", "potrazuje", "promet kr", "kr")
    amount = @("amount", "saldo", "znesek", "net amount", "vrednost")
    classification = @("klasifikacija", "classification", "class", "mapping")
    class_id = @("class id", "classification id", "map id", "id klasifikacije")
    counterparty = @("partner", "dobavitelj", "kupec", "counterparty", "vendor", "customer")
    document_number = @("dokument", "document", "temeljnica", "document number", "invoice", "racun", "račun")
}

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

function Normalize-Header {
    param([object]$Value)
    $text = [string]$Value
    $text = $text.Trim().ToLowerInvariant()
    $text = $text -replace "[._\-\/\\:;\,\(\)\[\]\{\}`"']", " "
    $text = $text -replace "\s+", " "
    return $text.Trim()
}

function Get-ColumnMap {
    param([object[]]$Headers)
    $map = @{}
    for ($index = 0; $index -lt $Headers.Count; $index++) {
        $headerText = Normalize-Header $Headers[$index]
        if ([string]::IsNullOrWhiteSpace($headerText)) {
            continue
        }

        foreach ($field in $fieldSynonyms.Keys) {
            if ($map.ContainsKey($field)) {
                continue
            }
            foreach ($synonym in $fieldSynonyms[$field]) {
                if ($headerText -eq (Normalize-Header $synonym) -or $headerText.Contains((Normalize-Header $synonym))) {
                    $map[$field] = $index
                    break
                }
            }
        }
    }
    return $map
}

function Escape-Csv {
    param([object]$Value)
    if ($null -eq $Value) {
        return '""'
    }
    $text = [string]$Value
    $text = $text.Replace('"', '""')
    return '"' + $text + '"'
}

function Get-CellText {
    param($Values, [int]$Row, [int]$Column)
    $value = $Values[$Row, $Column]
    if ($null -eq $value) {
        return ""
    }
    return [string]$value
}

function Resolve-DuckDb {
    if (-not [string]::IsNullOrWhiteSpace($DuckDbPath)) {
        if (-not (Test-Path -LiteralPath $DuckDbPath)) {
            throw "DuckDB path not found: $DuckDbPath"
        }
        return (Resolve-Path -LiteralPath $DuckDbPath).Path
    }

    $command = Get-Command duckdb -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $candidate = Join-Path $fasRoot "Tools\duckdb.exe"
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return $null
}

function Write-KindCsv {
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)][string]$KindToBuild
    )

    $stagingFolder = Join-Path $Project.folder "staging"
    if (-not (Test-Path -LiteralPath $stagingFolder)) {
        New-Item -ItemType Directory -Path $stagingFolder -Force | Out-Null
    }

    $csvPath = Join-Path $stagingFolder "$KindToBuild.csv"
    $writer = [System.IO.StreamWriter]::new($csvPath, $false, [System.Text.Encoding]::UTF8)
    try {
        $writer.WriteLine(($canonicalFields | ForEach-Object { Escape-Csv $_ }) -join ",")

        foreach ($source in @($Project.sources)) {
            if ($source.kind -ne $KindToBuild) {
                continue
            }
            if (-not (Test-Path -LiteralPath $source.path)) {
                Write-Warning "Source missing, skipping: $($source.path)"
                continue
            }

            if ([System.IO.Path]::GetExtension([string]$source.path) -ieq ".csv") {
                Write-CsvSourceRows -Writer $writer -Project $Project -Source $source -KindToBuild $KindToBuild
                continue
            }

            $workbook = $null
            try {
                $workbook = Open-WorkbookReadOnly -Excel $Excel -Path ([string]$source.path)
                if ([string]::IsNullOrWhiteSpace([string]$source.sheetName)) {
                    $sheet = $workbook.Worksheets.Item(1)
                } else {
                    $sheet = $workbook.Worksheets.Item($source.sheetName)
                }

                $used = $sheet.UsedRange
                $values = $used.Value2
                if ($null -eq $values) {
                    continue
                }

                $rowCount = $used.Rows.Count
                $columnCount = $used.Columns.Count
                $headerRow = [int]$source.headerRow
                if ($headerRow -lt 1) {
                    $headerRow = 1
                }

                $headers = @()
                for ($column = 1; $column -le $columnCount; $column++) {
                    $headers += (Get-CellText -Values $values -Row $headerRow -Column $column)
                }
                $columnMap = Get-ColumnMap -Headers $headers

                for ($row = $headerRow + 1; $row -le $rowCount; $row++) {
                    $allText = New-Object System.Collections.Generic.List[string]
                    for ($column = 1; $column -le $columnCount; $column++) {
                        $cellText = Get-CellText -Values $values -Row $row -Column $column
                        if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                            [void]$allText.Add($cellText)
                        }
                    }
                    if ($allText.Count -eq 0) {
                        continue
                    }

                    $record = @{}
                    foreach ($field in $canonicalFields) {
                        $record[$field] = ""
                    }
                    $record.kind = $KindToBuild
                    $record.project = $Project.name
                    $record.company = [string]$source.company
                    $record.year = [string]$source.year
                    $record.source_id = [string]$source.id
                    $record.source_file = [string]$source.originalName
                    $record.source_path = [string]$source.path
                    $record.source_sheet = [string]$sheet.Name
                    $record.source_row = [string]($used.Row + $row - 1)
                    $record.raw_text = ($allText -join " | ")

                    foreach ($field in $fieldSynonyms.Keys) {
                        if ($columnMap.ContainsKey($field)) {
                            $record[$field] = Get-CellText -Values $values -Row $row -Column ([int]$columnMap[$field] + 1)
                        }
                    }

                    $writer.WriteLine(($canonicalFields | ForEach-Object { Escape-Csv $record[$_] }) -join ",")
                }
            }
            finally {
                if ($null -ne $workbook) {
                    $workbook.Close($false)
                }
            }
        }
    }
    finally {
        $writer.Close()
    }

    return $csvPath
}

function Open-WorkbookReadOnly {
    param(
        [Parameter(Mandatory = $true)]$Excel,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    try {
        return $Excel.Workbooks.Open($resolvedPath, 0, $true, [Type]::Missing, "", "", $true, [Type]::Missing, [Type]::Missing, $false, $false, [Type]::Missing, $false, $true)
    }
    catch {
        throw "Could not open source workbook '$resolvedPath'. Close it in Excel, unblock it if it came from Downloads, or save the target sheet as CSV and add that CSV as the source. $($_.Exception.Message)"
    }
}

function Write-CsvSourceRows {
    param(
        [Parameter(Mandatory = $true)][System.IO.StreamWriter]$Writer,
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)]$Source,
        [Parameter(Mandatory = $true)][string]$KindToBuild
    )

    $rows = @(Import-Csv -LiteralPath ([string]$Source.path))
    if ($rows.Count -eq 0) {
        return
    }

    $headers = @($rows[0].PSObject.Properties.Name)
    $columnMap = Get-ColumnMap -Headers $headers
    $sourceRow = [int]$Source.headerRow

    foreach ($row in $rows) {
        $sourceRow++
        $allText = New-Object System.Collections.Generic.List[string]
        foreach ($header in $headers) {
            $cellText = [string]$row.$header
            if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                [void]$allText.Add($cellText)
            }
        }

        if ($allText.Count -eq 0) {
            continue
        }

        $record = @{}
        foreach ($field in $canonicalFields) {
            $record[$field] = ""
        }

        $record.kind = $KindToBuild
        $record.project = $Project.name
        $record.company = [string]$Source.company
        $record.year = [string]$Source.year
        $record.source_id = [string]$Source.id
        $record.source_file = [string]$Source.originalName
        $record.source_path = [string]$Source.path
        $record.source_sheet = [System.IO.Path]::GetFileNameWithoutExtension([string]$Source.originalName)
        $record.source_row = [string]$sourceRow
        $record.raw_text = ($allText -join " | ")

        foreach ($field in $fieldSynonyms.Keys) {
            if ($columnMap.ContainsKey($field)) {
                $headerName = $headers[[int]$columnMap[$field]]
                $record[$field] = [string]$row.$headerName
            }
        }

        $Writer.WriteLine(($canonicalFields | ForEach-Object { Escape-Csv $record[$_] }) -join ",")
    }
}

function Build-DuckDbTables {
    param(
        [Parameter(Mandatory = $true)][string]$DuckDb,
        [Parameter(Mandatory = $true)]$Project,
        [Parameter(Mandatory = $true)][string[]]$KindsBuilt
    )

    $indexFolder = Join-Path $Project.folder "index"
    if (-not (Test-Path -LiteralPath $indexFolder)) {
        New-Item -ItemType Directory -Path $indexFolder -Force | Out-Null
    }
    $databasePath = Join-Path $indexFolder "fas_index.duckdb"
    Set-Content -LiteralPath (Join-Path $indexFolder "duckdb_path.txt") -Value $DuckDb -Encoding UTF8

    foreach ($kindName in $KindsBuilt) {
        $csvPath = Join-Path (Join-Path $Project.folder "staging") "$kindName.csv"
        if (-not (Test-Path -LiteralPath $csvPath)) {
            continue
        }
        $parquetPath = Join-Path $indexFolder "$kindName.parquet"
        $csvSql = $csvPath.Replace("'", "''").Replace("\", "/")
        $parquetSql = $parquetPath.Replace("'", "''").Replace("\", "/")

        $sql = @"
CREATE OR REPLACE TABLE $kindName AS
SELECT * FROM read_csv('$csvSql', header = true, all_varchar = true);
COPY $kindName TO '$parquetSql' (FORMAT parquet);
"@
        & $DuckDb $databasePath -c $sql | Out-Null
    }

    return $databasePath
}

$state = Read-State
$project = Get-Project -State $state -Name $ProjectName
$kindsToBuild = if ($Kind -eq "all") { @("ledger", "tb", "mapping") } else { @($Kind) }

$excel = $null
$builtKinds = New-Object System.Collections.Generic.List[string]
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$KeepExcelVisible
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    foreach ($kindName in $kindsToBuild) {
        $csvPath = Write-KindCsv -Excel $excel -Project $project -KindToBuild $kindName
        [void]$builtKinds.Add($kindName)
        Write-Host "Wrote $kindName staging CSV: $csvPath"
    }
}
finally {
    if ($null -ne $excel) {
        $excel.Quit()
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$duckDb = Resolve-DuckDb
if ($null -eq $duckDb) {
    Write-Warning "DuckDB CLI was not found. Staging CSVs are ready, but DuckDB/Parquet index was not built."
    Write-Host "Put duckdb.exe in %APPDATA%\FAS_Toolkit\Tools or add it to PATH, then rerun this script."
    exit 0
}

$databasePath = Build-DuckDbTables -DuckDb $duckDb -Project $project -KindsBuilt $builtKinds.ToArray()
Write-Host "Built DuckDB index: $databasePath"
