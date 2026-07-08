$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$source = Get-ChildItem -File -Filter "*add12macro.xlsx" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $source) { throw "Source add12 workbook not found." }

$output = Join-Path $source.DirectoryName "20260526日报数据版V3_add12macro_trendfixed.xlsx"

function Read-EntryText($zip, [string]$name) {
    $entry = $zip.GetEntry($name)
    if ($null -eq $entry) { throw "Missing zip entry: $name" }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function XmlEscape([string]$s) {
    if ($null -eq $s) { return "" }
    return [System.Security.SecurityElement]::Escape($s)
}

function ColumnNumber([string]$letters) {
    $n = 0
    foreach ($ch in $letters.ToUpperInvariant().ToCharArray()) {
        $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $n
}

function ColumnLetters([int]$n) {
    $s = ""
    while ($n -gt 0) {
        $n--
        $s = [char]([int][char]'A' + ($n % 26)) + $s
        $n = [math]::Floor($n / 26)
    }
    return $s
}

function Parse-Cells([string]$sheetXml) {
    [xml]$ws = $sheetXml
    $cells = @{}
    foreach ($c in $ws.GetElementsByTagName("c")) {
        $ref = $c.GetAttribute("r")
        $type = $c.GetAttribute("t")
        $vNode = $c.GetElementsByTagName("v")
        $v = ""
        if ($vNode.Count -gt 0) { $v = $vNode.Item(0).InnerText }
        $cells[$ref] = [pscustomobject]@{ Type = $type; Value = $v }
    }
    return $cells
}

function Get-LastSixValues($cells, [int]$row) {
    $vals = New-Object System.Collections.Generic.List[object]
    for ($col = (ColumnNumber "F"); $col -le (ColumnNumber "SZ"); $col++) {
        $ref = (ColumnLetters $col) + $row
        if ($cells.ContainsKey($ref)) {
            $v = [string]$cells[$ref].Value
            if (-not [string]::IsNullOrWhiteSpace($v) -and $v -ne "#N/A") {
                $vals.Add($v)
            }
        }
    }
    $out = @()
    $start = [Math]::Max(0, $vals.Count - 6)
    for ($i = $start; $i -lt $vals.Count; $i++) { $out += [string]$vals[$i] }
    while ($out.Count -lt 6) { $out = @("#N/A") + $out }
    return $out
}

function CellXml([string]$ref, [string]$formula, [string]$value) {
    $f = XmlEscape $formula
    if ($value -eq "#N/A" -or [string]::IsNullOrWhiteSpace($value)) {
        return '<c r="' + $ref + '" t="e"><f>' + $f + '</f><v>#N/A</v></c>'
    }
    return '<c r="' + $ref + '"><f>' + $f + '</f><v>' + (XmlEscape $value) + '</v></c>'
}

function Build-AuxRows($dataCells) {
    $rows = @()
    for ($i = 0; $i -lt 12; $i++) {
        $srcRow = 114 + $i
        $helperRow = 72 + ($i * 2)
        $valueRow = $helperRow + 1
        $lastSix = Get-LastSixValues $dataCells $srcRow

        $helperCells = @()
        foreach ($col in @("B","C","D","E","F")) {
            $nextCol = [string][char](([int][char]$col) + 1)
            $nextCell = $nextCol + $helperRow
            $firstCell = '宏观数据!$F$' + $srcRow
            $formula = 'IFERROR(IF(' + $nextCell + '<=COLUMN(' + $firstCell + '),-1,LOOKUP(2,1/(OFFSET(' + $firstCell + ',0,0,1,' + $nextCell + '-COLUMN(' + $firstCell + '))<>""),COLUMN(OFFSET(' + $firstCell + ',0,0,1,' + $nextCell + '-COLUMN(' + $firstCell + '))))),-1)'
            $helperCells += CellXml ($col + $helperRow) $formula "-1"
        }
        $series = '宏观数据!$F$' + $srcRow + ':$SZ$' + $srcRow
        $lastColFormula = 'IFERROR(LOOKUP(2,1/(' + $series + '<>""),COLUMN(' + $series + ')),-1)'
        $helperCells += CellXml ("G$helperRow") $lastColFormula "-1"
        $rows += '<row r="' + $helperRow + '">' + ($helperCells -join "") + '</row>'

        $valueCells = @()
        $idx = 0
        foreach ($col in @("B","C","D","E","F","G")) {
            $helperCell = $col + $helperRow
            $firstCell = '宏观数据!$F$' + $srcRow
            $formula = 'IF(' + $helperCell + '=-1,NA(),INDEX(' + $series + ',1,' + $helperCell + '-COLUMN(' + $firstCell + ')+1))'
            $valueCells += CellXml ($col + $valueRow) $formula $lastSix[$idx]
            $idx++
        }
        $rows += '<row r="' + $valueRow + '">' + ($valueCells -join "") + '</row>'
    }
    return $rows -join ""
}

function Replace-AuxRows([string]$xml, [string]$newRows) {
    for ($r = 72; $r -le 95; $r++) {
        $pattern = '<row[^>]*\sr="' + $r + '"[\s\S]*?</row>'
        $xml = [regex]::Replace($xml, $pattern, "")
    }
    $xml = $xml -replace '<dimension ref="[^"]+"', '<dimension ref="A1:BI95"'
    return $xml -replace '</sheetData>', ($newRows + '</sheetData>')
}

function Build-ChartCache([string[]]$values) {
    $pts = ""
    for ($i = 0; $i -lt $values.Count; $i++) {
        $v = $values[$i]
        if ($v -eq "#N/A" -or [string]::IsNullOrWhiteSpace($v)) { $v = "0" }
        $pts += '<c:pt idx="' + $i + '"><c:v>' + (XmlEscape $v) + '</c:v></c:pt>'
    }
    return '<c:numCache><c:formatCode>0.0000</c:formatCode><c:ptCount val="6"/>' + $pts + '</c:numCache>'
}

function Build-NewCharts($templateChart, $dataCells) {
    $charts = @{}
    for ($i = 0; $i -lt 12; $i++) {
        $chartNum = 40 + $i
        $valueRow = 73 + ($i * 2)
        $srcRow = 114 + $i
        $formula = '辅助数据!$B$' + $valueRow + ':$G$' + $valueRow
        $chart = [regex]::Replace($templateChart, '<c:f>[^<]+</c:f>', '<c:f>' + (XmlEscape $formula) + '</c:f>', 1)
        $cache = Build-ChartCache (Get-LastSixValues $dataCells $srcRow)
        $chart = [regex]::Replace($chart, '<c:numCache>[\s\S]*?</c:numCache>', $cache, 1)
        $charts["xl/charts/chart$chartNum.xml"] = $chart
    }
    return $charts
}

function Build-NewAnchors() {
    $anchors = ""
    $x = 6902450
    $yStart = 11277600
    $rowStep = 504825
    for ($i = 0; $i -lt 12; $i++) {
        $frontRow = 24 + $i
        $fromRow = $frontRow - 1
        $toRow = $frontRow
        $chartName = 23 + $i
        $cNvId = 39 + $i
        $rId = 23 + $i
        $y = $yStart + ($rowStep * $i)
        $anchors += '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + $fromRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>10</xdr:col><xdr:colOff>559080</xdr:colOff><xdr:row>' + $toRow + '</xdr:row><xdr:rowOff>286560</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="' + $cNvId + '" name="Chart ' + $chartName + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="' + $x + '" y="' + $y + '"/><a:ext cx="1827530" cy="791210"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' + $rId + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
    }
    return $anchors
}

function Add-NewRelationships([string]$xml) {
    for ($i = 0; $i -lt 12; $i++) {
        $rId = 23 + $i
        $chartNum = 40 + $i
        if ($xml -notmatch ('Id="rId' + $rId + '"')) {
            $rel = '<Relationship Id="rId' + $rId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart' + $chartNum + '.xml"/>'
            $xml = $xml -replace '</Relationships>', ($rel + '</Relationships>')
        }
    }
    return $xml
}

function Add-ContentTypes([string]$xml) {
    for ($i = 40; $i -le 51; $i++) {
        $part = '/xl/charts/chart' + $i + '.xml'
        if ($xml -notmatch [regex]::Escape($part)) {
            $override = '<Override PartName="' + $part + '" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
            $xml = $xml -replace '</Types>', ($override + '</Types>')
        }
    }
    return $xml
}

function Write-TextEntry($zip, [string]$name, [string]$text) {
    $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        try { $writer.Write($text) } finally { $writer.Dispose() }
    } finally { $stream.Dispose() }
}

$inFs = [System.IO.File]::Open($source.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$inZip = New-Object System.IO.Compression.ZipArchive($inFs, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    $sheet3 = Read-EntryText $inZip "xl/worksheets/sheet3.xml"
    $sheet7 = Read-EntryText $inZip "xl/worksheets/sheet7.xml"
    $sheet4 = Read-EntryText $inZip "xl/worksheets/sheet4.xml"
    $drawing3 = Read-EntryText $inZip "xl/drawings/drawing3.xml"
    $drawing3Rels = Read-EntryText $inZip "xl/drawings/_rels/drawing3.xml.rels"
    $contentTypes = Read-EntryText $inZip "[Content_Types].xml"
    $templateChart = Read-EntryText $inZip "xl/charts/chart39.xml"

    $dataCells = Parse-Cells $sheet4
    $newAuxRows = Build-AuxRows $dataCells
    $sheet7 = Replace-AuxRows $sheet7 $newAuxRows
    $sheet3 = [regex]::Replace($sheet3, '<extLst>[\s\S]*?sparklineGroups[\s\S]*?</extLst>', "")
    $drawing3 = $drawing3 -replace '</xdr:wsDr>', ((Build-NewAnchors) + '</xdr:wsDr>')
    $drawing3Rels = Add-NewRelationships $drawing3Rels
    $contentTypes = Add-ContentTypes $contentTypes
    $newCharts = Build-NewCharts $templateChart $dataCells

    $tempOutput = $output + ".tmp"
    if (Test-Path -LiteralPath $tempOutput) { Remove-Item -LiteralPath $tempOutput -Force }
    $outFs = [System.IO.File]::Open($tempOutput, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    $outZip = New-Object System.IO.Compression.ZipArchive($outFs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        $replace = @{
            "xl/worksheets/sheet3.xml" = $sheet3
            "xl/worksheets/sheet7.xml" = $sheet7
            "xl/drawings/drawing3.xml" = $drawing3
            "xl/drawings/_rels/drawing3.xml.rels" = $drawing3Rels
            "[Content_Types].xml" = $contentTypes
        }
        foreach ($entry in $inZip.Entries) {
            if ($replace.ContainsKey($entry.FullName) -or $newCharts.ContainsKey($entry.FullName)) { continue }
            $newEntry = $outZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $src = $entry.Open()
            $dst = $newEntry.Open()
            try { $src.CopyTo($dst) } finally { $dst.Dispose(); $src.Dispose() }
        }
        foreach ($key in $replace.Keys) { Write-TextEntry $outZip $key $replace[$key] }
        foreach ($key in $newCharts.Keys) { Write-TextEntry $outZip $key $newCharts[$key] }
    } finally {
        $outZip.Dispose()
        $outFs.Dispose()
    }
    Move-Item -LiteralPath $tempOutput -Destination $output -Force
} finally {
    $inZip.Dispose()
    $inFs.Dispose()
}

Write-Output "FixedPath=$output"
Write-Output "AddedEmbeddedCharts=12"
Write-Output "RemovedSparklineGroups=12"
Write-Output "AuxRows=72-95"
