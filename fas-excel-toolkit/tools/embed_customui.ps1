[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [string]$CustomUiPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not $CustomUiPath) {
    $repoRoot = Split-Path -Path $PSScriptRoot -Parent
    $CustomUiPath = Join-Path $repoRoot "src\customUI\customUI14.xml"
}

if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
    throw "Workbook not found: $WorkbookPath"
}
if (-not (Test-Path -LiteralPath $CustomUiPath -PathType Leaf)) {
    throw "customUI file not found: $CustomUiPath"
}

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryText {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name)
    if ($null -eq $entry) { throw "Missing ZIP entry: $Name" }
    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally {
        $stream.Dispose()
    }
}

function Write-ZipEntryText {
    param($Zip, [string]$Name, [string]$Text)
    $existing = $Zip.GetEntry($Name)
    if ($null -ne $existing) { $existing.Delete() }
    $entry = $Zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        $writer = [System.IO.StreamWriter]::new($stream, $utf8NoBom)
        try { $writer.Write($Text) } finally { $writer.Dispose() }
    } finally {
        $stream.Dispose()
    }
}

function Ensure-ContentTypeOverride {
    param([xml]$Xml)
    $ns = "http://schemas.openxmlformats.org/package/2006/content-types"
    foreach ($node in $Xml.Types.Override) {
        if ($node.PartName -eq "/customUI/customUI14.xml") { return }
    }
    $override = $Xml.CreateElement("Override", $ns)
    $override.SetAttribute("PartName", "/customUI/customUI14.xml")
    $override.SetAttribute("ContentType", "application/vnd.ms-office.customUI+xml")
    [void]$Xml.Types.AppendChild($override)
}

function Ensure-CustomUiRelationship {
    param([xml]$Xml)
    $ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    $type = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    foreach ($node in $Xml.Relationships.Relationship) {
        if ($node.Type -eq $type -and $node.Target -eq "customUI/customUI14.xml") { return }
    }

    $used = @{}
    foreach ($node in $Xml.Relationships.Relationship) { $used[[string]$node.Id] = $true }
    $i = 1
    do {
        $id = "rIdCustomUI$i"
        $i++
    } while ($used.ContainsKey($id))

    $rel = $Xml.CreateElement("Relationship", $ns)
    $rel.SetAttribute("Id", $id)
    $rel.SetAttribute("Type", $type)
    $rel.SetAttribute("Target", "customUI/customUI14.xml")
    [void]$Xml.Relationships.AppendChild($rel)
}

$resolvedWorkbook = (Resolve-Path -LiteralPath $WorkbookPath).Path
$resolvedCustomUi = (Resolve-Path -LiteralPath $CustomUiPath).Path
$customUiText = Get-Content -LiteralPath $resolvedCustomUi -Raw

$zip = [System.IO.Compression.ZipFile]::Open($resolvedWorkbook, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    Write-ZipEntryText -Zip $zip -Name "customUI/customUI14.xml" -Text $customUiText

    [xml]$contentTypes = Read-ZipEntryText -Zip $zip -Name "[Content_Types].xml"
    Ensure-ContentTypeOverride -Xml $contentTypes
    Write-ZipEntryText -Zip $zip -Name "[Content_Types].xml" -Text $contentTypes.OuterXml

    [xml]$rels = Read-ZipEntryText -Zip $zip -Name "_rels/.rels"
    Ensure-CustomUiRelationship -Xml $rels
    Write-ZipEntryText -Zip $zip -Name "_rels/.rels" -Text $rels.OuterXml
}
finally {
    $zip.Dispose()
}

Write-Host "Embedded customUI into: $resolvedWorkbook"
