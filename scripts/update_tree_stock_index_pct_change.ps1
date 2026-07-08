param(
    [string]$TreePath = "",
    [string]$OutputPath = "",
    [string]$SheetName = "重点策略跟踪情况(V2.5)(1)"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-XmlDoc {
    param([string]$XmlText)
    $doc = [System.Xml.XmlDocument]::new()
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($XmlText)
    return $doc
}

function Get-ZipEntryText {
    param($Archive, [string]$Path)
    $entry = $Archive.GetEntry($Path)
    if ($null -eq $entry) { throw "Zip entry not found: $Path" }
    $stream = $entry.Open()
    try {
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Get-SharedStrings {
    param($Archive)
    $entry = $Archive.GetEntry("xl/sharedStrings.xml")
    if ($null -eq $entry) { return @() }
    $doc = Read-XmlDoc (Get-ZipEntryText $Archive "xl/sharedStrings.xml")
    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $strings = New-Object System.Collections.Generic.List[string]
    foreach ($si in $doc.SelectNodes("//x:si", $ns)) {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($t in $si.SelectNodes(".//x:t", $ns)) { $parts.Add($t.InnerText) }
        $strings.Add(($parts -join ""))
    }
    return $strings.ToArray()
}

function Get-SheetPath {
    param($Archive, [string]$Name)
    $workbook = Read-XmlDoc (Get-ZipEntryText $Archive "xl/workbook.xml")
    $wbNs = [System.Xml.XmlNamespaceManager]::new($workbook.NameTable)
    $wbNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $wbNs.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    $sheet = $workbook.SelectSingleNode("//x:sheet[@name='$Name']", $wbNs)
    if ($null -eq $sheet) {
        $sheet = $workbook.SelectSingleNode("//x:sheets/x:sheet[1]", $wbNs)
        if ($null -eq $sheet) { throw "Sheet not found: $Name" }
    }
    $rid = $sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

    $rels = Read-XmlDoc (Get-ZipEntryText $Archive "xl/_rels/workbook.xml.rels")
    $relsNs = [System.Xml.XmlNamespaceManager]::new($rels.NameTable)
    $relsNs.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
    $rel = $rels.SelectSingleNode("//r:Relationship[@Id='$rid']", $relsNs)
    if ($null -eq $rel) { throw "Workbook relationship not found: $rid" }
    $target = $rel.GetAttribute("Target")
    if ($target.StartsWith("/")) { return $target.TrimStart("/") }
    return "xl/" + $target.TrimStart("/")
}

function Get-CellText {
    param($Cell, [string[]]$Shared, $Ns)
    if ($null -eq $Cell) { return $null }
    $type = $Cell.GetAttribute("t")
    if ($type -eq "inlineStr") {
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($t in $Cell.SelectNodes(".//x:t", $Ns)) { $parts.Add($t.InnerText) }
        return ($parts -join "")
    }
    $v = $Cell.SelectSingleNode("x:v", $Ns)
    if ($null -eq $v) { return $null }
    if ($type -eq "s") {
        $idx = [int]$v.InnerText
        if ($idx -ge 0 -and $idx -lt $Shared.Length) { return $Shared[$idx] }
    }
    return $v.InnerText
}

function Set-NumericCell {
    param($Doc, $Ns, [int]$RowNumber, [string]$Column, [double]$Value, [string]$StyleId)
    $row = $Doc.SelectSingleNode("//x:sheetData/x:row[@r='$RowNumber']", $Ns)
    if ($null -eq $row) { throw "Row not found: $RowNumber" }
    $cellRef = "$Column$RowNumber"
    $cell = $row.SelectSingleNode("x:c[@r='$cellRef']", $Ns)
    if ($null -eq $cell) { throw "Cell not found: $cellRef" }
    while ($cell.HasChildNodes) { [void]$cell.RemoveChild($cell.FirstChild) }
    $cell.RemoveAttribute("t")
    if (-not [string]::IsNullOrWhiteSpace($StyleId)) { $cell.SetAttribute("s", $StyleId) }
    $v = $Doc.CreateElement("v", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $v.InnerText = $Value.ToString("0.###############", [System.Globalization.CultureInfo]::InvariantCulture)
    [void]$cell.AppendChild($v)
}

function Get-NextVersionPath {
    param([System.IO.FileInfo]$Source)
    $baseName = $Source.BaseName
    $ext = $Source.Extension
    if ($baseName -match "^(.*)_V(\d+)(?:\.(\d+))?$") {
        $prefix = $matches[1]
        $major = [int]$matches[2]
        $minor = if ($matches[3]) { [int]$matches[3] + 1 } else { 1 }
        for ($i = $minor; $i -le 99; $i++) {
            $candidate = Join-Path $Source.DirectoryName ("{0}_V{1}.{2}{3}" -f $prefix, $major, $i, $ext)
            if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
        }
    }
    for ($i = 1; $i -le 99; $i++) {
        $candidate = Join-Path $Source.DirectoryName ("{0}_pctchange_V1.{1}{2}" -f $baseName, $i, $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    throw "No available output path."
}

if ([string]::IsNullOrWhiteSpace($TreePath)) {
    $source = Get-ChildItem -File -Filter "*TREE*.xlsx" |
        Where-Object { $_.Name -notlike "~$*" } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $source) { throw "TREE workbook not found." }
    $TreePath = $source.FullName
}

$sourceItem = Get-Item -LiteralPath $TreePath
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Get-NextVersionPath $sourceItem }
if (Test-Path -LiteralPath $OutputPath) { throw "Output already exists: $OutputPath" }

Copy-Item -LiteralPath $TreePath -Destination $OutputPath

$targetRows = @(50, 51, 52, 53, 60, 61, 112, 113, 114)
$archive = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    $shared = Get-SharedStrings $archive
    $sheetPath = Get-SheetPath $archive $SheetName
    $doc = Read-XmlDoc (Get-ZipEntryText $archive $sheetPath)
    $ns = [System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    $percentStyle = ($doc.SelectSingleNode("//x:c[@r='M6']", $ns)).GetAttribute("s")

    $results = New-Object System.Collections.Generic.List[object]
    foreach ($row in $targetRows) {
        $name = Get-CellText ($doc.SelectSingleNode("//x:c[@r='D$row']", $ns)) $shared $ns
        $currentText = Get-CellText ($doc.SelectSingleNode("//x:c[@r='K$row']", $ns)) $shared $ns
        $changeText = Get-CellText ($doc.SelectSingleNode("//x:c[@r='M$row']", $ns)) $shared $ns
        if ([string]::IsNullOrWhiteSpace($currentText) -or [string]::IsNullOrWhiteSpace($changeText)) { continue }
        $current = [double]::Parse($currentText, [System.Globalization.CultureInfo]::InvariantCulture)
        $change = [double]::Parse($changeText, [System.Globalization.CultureInfo]::InvariantCulture)
        $previous = $current - $change
        if ([math]::Abs($previous) -lt 0.0000001) { continue }
        $pctChange = $change / $previous
        Set-NumericCell $doc $ns $row "M" $pctChange $percentStyle
        $results.Add([pscustomobject]@{
            Row = $row
            Indicator = $name
            Current = $current
            PointChange = $change
            PctChange = $pctChange
        })
    }

    $oldEntry = $archive.GetEntry($sheetPath)
    $oldEntry.Delete()
    $newEntry = $archive.CreateEntry($sheetPath)
    $stream = $newEntry.Open()
    try {
        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.Encoding]::UTF8)
        try { $doc.Save($writer) } finally { $writer.Dispose() }
    } finally { $stream.Dispose() }
} finally {
    $archive.Dispose()
}

[pscustomobject]@{
    OutputPath = $OutputPath
    UpdatedRows = $results.Count
    Rows = $results
} | ConvertTo-Json -Depth 5
