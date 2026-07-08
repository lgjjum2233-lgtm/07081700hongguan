Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Open-ZipRead($path) {
    return [System.IO.Compression.ZipFile]::OpenRead($path)
}

function Open-ZipUpdate($path) {
    return [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Update)
}

function Get-ZipEntryText($zip, [string]$entryName) {
    $entry = $zip.GetEntry($entryName)
    if ($null -eq $entry) { return $null }
    $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Set-ZipEntryText($zip, [string]$entryName, [string]$text) {
    $entry = $zip.GetEntry($entryName)
    if ($null -ne $entry) { $entry.Delete() }
    $entry = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $writer = New-Object System.IO.StreamWriter($entry.Open(), (New-Object System.Text.UTF8Encoding($false)))
    try { $writer.Write($text) } finally { $writer.Dispose() }
}

function Xml-Doc([string]$text) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($text)
    return $doc
}

function Get-SharedStrings($zip) {
    $text = Get-ZipEntryText $zip 'xl/sharedStrings.xml'
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    $doc = Xml-Doc $text
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $items = @()
    foreach ($si in $doc.SelectNodes('//x:si', $ns)) {
        $parts = @()
        foreach ($t in $si.SelectNodes('.//x:t', $ns)) { $parts += $t.InnerText }
        $items += ($parts -join '')
    }
    return $items
}

function Get-PercentStyleIds($zip) {
    $text = Get-ZipEntryText $zip 'xl/styles.xml'
    $set = New-Object 'System.Collections.Generic.HashSet[int]'
    if ([string]::IsNullOrWhiteSpace($text)) { return $set }
    $doc = Xml-Doc $text
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $numFmts = @{}
    foreach ($fmt in $doc.SelectNodes('//x:numFmts/x:numFmt', $ns)) {
        $numFmts[[int]$fmt.GetAttribute('numFmtId')] = $fmt.GetAttribute('formatCode')
    }
    $xfIndex = 0
    foreach ($xf in $doc.SelectNodes('//x:cellXfs/x:xf', $ns)) {
        $numFmtId = 0
        if ($xf.HasAttribute('numFmtId')) { $numFmtId = [int]$xf.GetAttribute('numFmtId') }
        $isPct = ($numFmtId -eq 9 -or $numFmtId -eq 10)
        if (-not $isPct -and $numFmts.ContainsKey($numFmtId)) {
            $isPct = ($numFmts[$numFmtId] -like '*%*')
        }
        if ($isPct) { [void]$set.Add($xfIndex) }
        $xfIndex++
    }
    return $set
}

function Get-WorkbookSheets($zip) {
    $wbText = Get-ZipEntryText $zip 'xl/workbook.xml'
    $relsText = Get-ZipEntryText $zip 'xl/_rels/workbook.xml.rels'
    $wb = Xml-Doc $wbText
    $rels = Xml-Doc $relsText
    $ns = New-Object System.Xml.XmlNamespaceManager($wb.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $relNs = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
    $relNs.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relMap = @{}
    foreach ($rel in $rels.SelectNodes('//rel:Relationship', $relNs)) {
        $relMap[$rel.Id] = $rel.Target
    }
    $sheets = @()
    foreach ($s in $wb.SelectNodes('//x:sheets/x:sheet', $ns)) {
        $rid = $s.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $target = $relMap[$rid]
        if ($target -notmatch '^xl/') { $target = 'xl/' + $target.TrimStart('/') }
        $sheets += [pscustomobject]@{
            Name = $s.GetAttribute('name')
            SheetId = $s.GetAttribute('sheetId')
            Rid = $rid
            Path = $target
        }
    }
    return $sheets
}

function Get-ColumnName([int]$n) {
    $name = ''
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $name = [char](65 + $rem) + $name
        $n = [math]::Floor(($n - 1) / 26)
    }
    return $name
}

function Get-ColumnNumber([string]$name) {
    $n = 0
    foreach ($ch in $name.ToUpperInvariant().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { continue }
        $n = $n * 26 + ([int][char]$ch - [int][char]'A' + 1)
    }
    return $n
}

function Get-SheetCells($zip, [string]$sheetPath, $sharedStrings) {
    $text = Get-ZipEntryText $zip $sheetPath
    $doc = Xml-Doc $text
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('x', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $cells = @{}
    foreach ($c in $doc.SelectNodes('//x:c', $ns)) {
        $ref = $c.GetAttribute('r')
        if ($ref -notmatch '^([A-Z]+)([0-9]+)$') { continue }
        $col = $matches[1]
        $row = [int]$matches[2]
        $t = $c.GetAttribute('t')
        $vNode = $c.SelectSingleNode('x:v', $ns)
        $value = ''
        if ($t -eq 's' -and $null -ne $vNode) {
            $idx = [int]$vNode.InnerText
            if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) { $value = $sharedStrings[$idx] }
        } elseif ($t -eq 'inlineStr') {
            $texts = @()
            foreach ($tn in $c.SelectNodes('.//x:t', $ns)) { $texts += $tn.InnerText }
            $value = ($texts -join '')
        } elseif ($null -ne $vNode) {
            $value = $vNode.InnerText
        }
        $cells["$col$row"] = $value
    }
    return $cells
}

function Get-Cell($cells, [string]$col, [int]$row) {
    $key = "$col$row"
    if ($cells.ContainsKey($key)) { return $cells[$key] }
    return ''
}

function Is-NumberText([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $false }
    return $s.Trim() -match '^-?[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?$'
}

function To-DoubleOrNull([string]$s) {
    if (-not (Is-NumberText $s)) { return $null }
    return [double]::Parse($s.Trim(), [System.Globalization.CultureInfo]::InvariantCulture)
}

function HtmlEncode([string]$s) {
    return [System.Security.SecurityElement]::Escape($s)
}

function Get-StyleAttr([string]$sheetText, [string]$ref, [string]$fallbackRef) {
    $m = [regex]::Match($sheetText, "<c\b[^>]*\br=`"$ref`"[^>]*>")
    if (-not $m.Success -and -not [string]::IsNullOrWhiteSpace($fallbackRef)) {
        $m = [regex]::Match($sheetText, "<c\b[^>]*\br=`"$fallbackRef`"[^>]*>")
    }
    if ($m.Success -and $m.Value -match '\bs="([^"]+)"') {
        return ' s="' + $matches[1] + '"'
    }
    return ''
}

function Make-CellXml([string]$ref, $value, [string]$styleAttr = '') {
    if ($null -eq $value) { $value = '' }
    $s = [string]$value
    if (Is-NumberText $s) {
        return "<c r=`"$ref`"$styleAttr><v>$s</v></c>"
    }
    return "<c r=`"$ref`"$styleAttr t=`"inlineStr`"><is><t>$(HtmlEncode $s)</t></is></c>"
}

function Convert-ValueForStyle($value, [string]$styleAttr, $percentStyleIds) {
    if ($null -eq $value) { return $value }
    $s = [string]$value
    if (-not (Is-NumberText $s)) { return $value }
    if ($styleAttr -match 's="([0-9]+)"') {
        $styleId = [int]$matches[1]
        if ($percentStyleIds.Contains($styleId)) {
            $num = [double]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture) / 100.0
            return $num.ToString('G15', [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return $value
}

function Escape-FormulaSheet([string]$name) {
    if ($name -match '[ !''\[\]]') {
        return "'" + ($name -replace "'", "''") + "'"
    }
    return $name
}

function Get-DrawingPathForSheet($zip, [string]$sheetPath) {
    $dir = [System.IO.Path]::GetDirectoryName($sheetPath).Replace('\', '/')
    $file = [System.IO.Path]::GetFileName($sheetPath)
    $relsPath = "$dir/_rels/$file.rels"
    $relsText = Get-ZipEntryText $zip $relsPath
    if ([string]::IsNullOrWhiteSpace($relsText)) { return $null }
    $rels = Xml-Doc $relsText
    $ns = New-Object System.Xml.XmlNamespaceManager($rels.NameTable)
    $ns.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
    foreach ($rel in $rels.SelectNodes('//rel:Relationship', $ns)) {
        if ($rel.Type -like '*drawing') {
            $target = $rel.Target
            if ($target.StartsWith('../')) {
                $base = ($dir -replace '/worksheets$', '')
                return ($base + '/' + $target.Substring(3)).Replace('//', '/')
            }
            if ($target -notmatch '^xl/') { return ($dir + '/' + $target).Replace('//', '/') }
            return $target
        }
    }
    return $null
}

function Resolve-RelTarget([string]$relsPath, [string]$target) {
    $base = [System.IO.Path]::GetDirectoryName(($relsPath -replace '/_rels/[^/]+\.rels$', '')).Replace('\', '/')
    $ownerDir = ($relsPath -replace '/_rels/[^/]+\.rels$', '')
    if ($target.StartsWith('../')) {
        $cur = $ownerDir
        $rest = $target
        while ($rest.StartsWith('../')) {
            $cur = [System.IO.Path]::GetDirectoryName($cur).Replace('\', '/')
            $rest = $rest.Substring(3)
        }
        return ($cur + '/' + $rest).Replace('//', '/')
    }
    if ($target -match '^xl/') { return $target }
    return ($ownerDir + '/' + $target).Replace('//', '/')
}

function Get-ChartValues($zip, [string]$chartPath) {
    $text = Get-ZipEntryText $zip $chartPath
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    $doc = Xml-Doc $text
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart')
    $cache = $doc.SelectSingleNode('//c:ser/c:val/c:numRef/c:numCache', $ns)
    if ($null -eq $cache) { $cache = $doc.SelectSingleNode('//c:ser/c:val/c:numLit', $ns) }
    if ($null -eq $cache) { return @() }
    $vals = @()
    foreach ($pt in $cache.SelectNodes('c:pt', $ns)) {
        $v = $pt.SelectSingleNode('c:v', $ns)
        if ($null -ne $v -and (Is-NumberText $v.InnerText)) {
            $vals += $v.InnerText
        }
    }
    return $vals
}

function Get-ChartMapForSheet($zip, [string]$sheetPath) {
    $drawingPath = Get-DrawingPathForSheet $zip $sheetPath
    $map = @{}
    if ($null -eq $drawingPath) { return $map }
    $drawingText = Get-ZipEntryText $zip $drawingPath
    if ([string]::IsNullOrWhiteSpace($drawingText)) { return $map }
    $relsPath = ([System.IO.Path]::GetDirectoryName($drawingPath).Replace('\', '/')) + '/_rels/' + ([System.IO.Path]::GetFileName($drawingPath)) + '.rels'
    $relsText = Get-ZipEntryText $zip $relsPath
    if ([string]::IsNullOrWhiteSpace($relsText)) { return $map }
    $relsDoc = Xml-Doc $relsText
    $relsNs = New-Object System.Xml.XmlNamespaceManager($relsDoc.NameTable)
    $relsNs.AddNamespace('rel', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relMap = @{}
    foreach ($rel in $relsDoc.SelectNodes('//rel:Relationship', $relsNs)) {
        if ($rel.Type -like '*chart') { $relMap[$rel.Id] = (Resolve-RelTarget $relsPath $rel.Target) }
    }
    $doc = Xml-Doc $drawingText
    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
    $ns.AddNamespace('c', 'http://schemas.openxmlformats.org/drawingml/2006/chart')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    foreach ($anchor in $doc.SelectNodes('/xdr:wsDr/xdr:twoCellAnchor|/xdr:wsDr/xdr:oneCellAnchor', $ns)) {
        $rowNode = $anchor.SelectSingleNode('xdr:from/xdr:row', $ns)
        $chart = $anchor.SelectSingleNode('.//c:chart', $ns)
        if ($null -eq $rowNode -or $null -eq $chart) { continue }
        $rid = $chart.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if (-not $relMap.ContainsKey($rid)) { continue }
        $row = [int]$rowNode.InnerText + 1
        $vals = Get-ChartValues $zip $relMap[$rid]
        if ($vals.Count -gt 0) { $map[$row] = $vals }
    }
    return $map
}

function Get-RowHeightMap([string]$sheetText) {
    $map = @{}
    foreach ($m in [regex]::Matches($sheetText, '<row\b[^>]*\br="([0-9]+)"[^>]*>')) {
        $tag = $m.Value
        if ($tag -match '\bht="([^"]+)"') {
            $map[[int]$m.Groups[1].Value] = [double]::Parse($matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return $map
}

function Get-SourceSeries($cells, [int]$row, [int]$startCol, [int]$endCol) {
    $items = @()
    for ($c = $startCol; $c -le $endCol; $c++) {
        $col = Get-ColumnName $c
        $val = Get-Cell $cells $col $row
        if (-not (Is-NumberText $val)) { continue }
        $date = Get-Cell $cells $col 1
        $items += [pscustomobject]@{ Col = $c; Date = $date; Value = $val; Number = (To-DoubleOrNull $val) }
    }
    return $items
}

function Make-SourceRecord($cells, [int]$row, [string]$freq) {
    $series = Get-SourceSeries $cells $row 6 520
    if ($series.Count -eq 0) { return $null }
    $last = $series[-1]
    $prev = if ($series.Count -ge 2) { $series[-2] } else { $null }
    $change = ''
    if ($null -ne $prev -and $null -ne $last.Number -and $null -ne $prev.Number) {
        $change = ([math]::Round(($last.Number - $prev.Number), 6)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    $take = 6
    if ($freq -like '*day*' -or $freq -like '*日*') { $take = 60 }
    $trend = @($series | Select-Object -Last $take | ForEach-Object { $_.Value })
    return [pscustomobject]@{
        Current = $last.Value
        Date = $last.Date
        Change = $change
        Trend = $trend
    }
}

function Make-ComputedDiffRecord($cells, [int]$rowA, [int]$rowB, [double]$factor, [string]$freq) {
    $a = Get-SourceSeries $cells $rowA 6 520
    $b = Get-SourceSeries $cells $rowB 6 520
    if ($a.Count -eq 0 -or $b.Count -eq 0) { return $null }
    $bByCol = @{}
    foreach ($x in $b) { $bByCol[$x.Col] = $x }
    $series = @()
    foreach ($x in $a) {
        if (-not $bByCol.ContainsKey($x.Col)) { continue }
        $y = $bByCol[$x.Col]
        if ($null -eq $x.Number -or $null -eq $y.Number) { continue }
        $v = ($x.Number - $y.Number) * $factor
        $series += [pscustomobject]@{
            Col = $x.Col
            Date = $x.Date
            Value = ([math]::Round($v, 6)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
            Number = $v
        }
    }
    if ($series.Count -eq 0) { return $null }
    $last = $series[-1]
    $prev = if ($series.Count -ge 2) { $series[-2] } else { $null }
    $change = ''
    if ($null -ne $prev) {
        $change = ([math]::Round(($last.Number - $prev.Number), 6)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    $take = 6
    if ($freq -like '*day*' -or $freq -like '*日*') { $take = 60 }
    $trend = @($series | Select-Object -Last $take | ForEach-Object { $_.Value })
    return [pscustomobject]@{
        Current = $last.Value
        Date = $last.Date
        Change = $change
        Trend = $trend
    }
}

function Make-FrontRecord($frontSheets, [int]$sheetIndex, [int]$row) {
    $sheet = $frontSheets[$sheetIndex]
    $trend = @()
    if ($sheet.ChartMap.ContainsKey($row)) { $trend = @($sheet.ChartMap[$row]) }
    return [pscustomobject]@{
        Current = Get-Cell $sheet.Cells 'E' $row
        Date = Get-Cell $sheet.Cells 'F' $row
        Change = Get-Cell $sheet.Cells 'G' $row
        Trend = $trend
    }
}

function Make-ChartXml([string]$formula, [string[]]$values) {
    $pts = ''
    for ($i = 0; $i -lt $values.Count; $i++) {
        $pts += "<c:pt idx=`"$i`"><c:v>$($values[$i])</c:v></c:pt>"
    }
    $count = $values.Count
    return @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="zh-CN"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    <c:autoTitleDeleted val="1"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:spPr>
            <a:ln w="19050">
              <a:solidFill><a:srgbClr val="E60000"/></a:solidFill>
              <a:round/>
            </a:ln>
          </c:spPr>
          <c:marker><c:symbol val="none"/></c:marker>
          <c:val>
            <c:numRef>
              <c:f>$formula</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="$count"/>
                $pts
              </c:numCache>
            </c:numRef>
          </c:val>
          <c:smooth val="0"/>
        </c:ser>
        <c:dLbls>
          <c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>
        </c:dLbls>
        <c:axId val="230315520"/>
        <c:axId val="230317568"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="230315520"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:tickLblPos val="none"/><c:crossAx val="230317568"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="230317568"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="none"/><c:crossAx val="230315520"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:delete val="1"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </c:spPr>
</c:chartSpace>
"@
}

function Make-ChartAnchor([int]$row, [string]$rid, [int]$cx, [int]$cy) {
    $fromRow = $row - 1
    return @"
<xdr:oneCellAnchor>
  <xdr:from><xdr:col>17</xdr:col><xdr:colOff>50000</xdr:colOff><xdr:row>$fromRow</xdr:row><xdr:rowOff>30000</xdr:rowOff></xdr:from>
  <xdr:ext cx="$cx" cy="$cy"/>
  <xdr:graphicFrame macro="">
    <xdr:nvGraphicFramePr><xdr:cNvPr id="$($row + 1000)" name="Chart $row"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="$rid"/>
      </a:graphicData>
    </a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:oneCellAnchor>
"@
}

function Replace-Or-Insert-RowCells([string]$sheetText, [int]$row, [hashtable]$cellsToAdd) {
    $rowPattern = "(?s)<row\b[^>]*\br=`"$row`"[^>]*>.*?</row>"
    $m = [regex]::Match($sheetText, $rowPattern)
    if (-not $m.Success) { return $sheetText }
    $rowXml = $m.Value
    $cellMap = @{}
    foreach ($cm in [regex]::Matches($rowXml, '<c\b[^>]*\br="([A-Z]+[0-9]+)"[^>]*(?:/>|>.*?</c>)')) {
        $cellMap[$cm.Groups[1].Value] = $cm.Value
    }
    foreach ($ref in $cellsToAdd.Keys) {
        if ($cellMap.ContainsKey($ref)) { $cellMap.Remove($ref) }
    }
    for ($c = 27; $c -le 80; $c++) {
        $ref = (Get-ColumnName $c) + $row
        if ($cellMap.ContainsKey($ref)) { $cellMap.Remove($ref) }
    }
    foreach ($ref in $cellsToAdd.Keys) {
        $cellMap[$ref] = $cellsToAdd[$ref]
    }
    $rowXml = [regex]::Replace($rowXml, '<c\b[^>]*\br="[A-Z]+[0-9]+"[^>]*(?:/>|>.*?</c>)', '')
    $orderedRefs = @($cellMap.Keys | Sort-Object { 
        if ($_ -match '^([A-Z]+)([0-9]+)$') { return Get-ColumnNumber $matches[1] }
        return 9999
    })
    $insert = ''
    foreach ($ref in $orderedRefs) { $insert += $cellMap[$ref] }
    $rowXml = $rowXml -replace '</row>$', ($insert + '</row>')
    return $sheetText.Substring(0, $m.Index) + $rowXml + $sheetText.Substring($m.Index + $m.Length)
}

function Get-VersionedPath([string]$dir, [string]$baseName, [string]$ext) {
    $candidate = Join-Path $dir ($baseName + '_V1' + $ext)
    if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    for ($i = 1; $i -le 99; $i++) {
        $candidate = Join-Path $dir ($baseName + '_V1.' + $i + $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    throw 'No available versioned output path.'
}

$daily = Get-ChildItem -File -Filter '*add12macro*.xlsx' |
    Where-Object { $_.Length -ge 490000 } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $daily) { throw 'Daily workbook not found.' }

$summary = Get-ChildItem -File -Filter '0521*.xlsx' |
    Where-Object { $_.Length -lt 300000 -and $_.Name -notlike '*~$*' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $summary) { throw 'Summary workbook not found.' }

$datePrefix = Get-Date -Format 'yyyyMMdd'
$baseName = $datePrefix + ($summary.BaseName -replace '^[0-9]{4,8}', '') + '_matched'
$output = Get-VersionedPath $summary.DirectoryName $baseName $summary.Extension
Copy-Item -LiteralPath $summary.FullName -Destination $output -Force

$dailyZip = Open-ZipRead $daily.FullName
$summaryZip = $null
try {
    $dailyShared = Get-SharedStrings $dailyZip
    $dailySheets = Get-WorkbookSheets $dailyZip
    $frontSheets = @{}
    for ($i = 0; $i -lt 3; $i++) {
        $frontSheets[$i + 1] = [pscustomobject]@{
            Cells = Get-SheetCells $dailyZip $dailySheets[$i].Path $dailyShared
            ChartMap = Get-ChartMapForSheet $dailyZip $dailySheets[$i].Path
        }
    }
    $macroCells = Get-SheetCells $dailyZip $dailySheets[3].Path $dailyShared
    $indexCells = Get-SheetCells $dailyZip $dailySheets[4].Path $dailyShared

    $summaryZip = Open-ZipUpdate $output
    $summaryShared = Get-SharedStrings $summaryZip
    $percentStyleIds = Get-PercentStyleIds $summaryZip
    $summarySheets = Get-WorkbookSheets $summaryZip
    $summarySheet = $summarySheets[0]
    $summaryCells = Get-SheetCells $summaryZip $summarySheet.Path $summaryShared
    $sheetText = Get-ZipEntryText $summaryZip $summarySheet.Path
    $rowHeights = Get-RowHeightMap $sheetText
    $sheetNameFormula = Escape-FormulaSheet $summarySheet.Name

    $dash = [string][char]0x2014

    $rowMap = @{
        6=@{kind='front';sheet=3;row=2}; 7=@{kind='source';row=27}; 8=@{kind='source';row=79}; 9=@{kind='source';row=80}; 10=@{kind='front';sheet=3;row=5};
        11=@{kind='source';row=28}; 12=@{kind='source';row=29}; 13=@{kind='source';row=30}; 14=@{kind='source';row=31};
        15=@{kind='source';row=81}; 16=@{kind='source';row=82}; 17=@{kind='source';row=83}; 18=@{kind='source';row=84};
        19=@{kind='front';sheet=3;row=24}; 20=@{kind='front';sheet=3;row=25}; 21=@{kind='front';sheet=3;row=26}; 22=@{kind='front';sheet=3;row=28}; 23=@{kind='front';sheet=3;row=27};
        24=@{kind='source';row=39}; 25=@{kind='source';row=41};
        28=@{kind='front';sheet=3;row=30}; 29=@{kind='source';row=9}; 30=@{kind='source';row=15}; 31=@{kind='source';row=105}; 32=@{kind='front';sheet=3;row=31};
        33=@{kind='source';row=21}; 34=@{kind='source';row=3}; 35=@{kind='front';sheet=3;row=32}; 36=@{kind='front';sheet=3;row=33}; 37=@{kind='source';row=8};
        38=@{kind='source';row=20}; 39=@{kind='source';row=23}; 40=@{kind='source';row=25}; 41=@{kind='source';row=107}; 42=@{kind='front';sheet=3;row=35}; 43=@{kind='front';sheet=3;row=34};
        50=@{kind='front';sheet=1;row=2}; 51=@{kind='front';sheet=1;row=3}; 52=@{kind='front';sheet=1;row=4}; 53=@{kind='front';sheet=1;row=5}; 54=@{kind='front';sheet=1;row=6};
        56=@{kind='front';sheet=1;row=7}; 57=@{kind='front';sheet=1;row=8}; 58=@{kind='front';sheet=1;row=9}; 59=@{kind='source';row=86};
        61=@{kind='front';sheet=1;row=11}; 62=@{kind='front';sheet=1;row=12}; 63=@{kind='front';sheet=1;row=14}; 64=@{kind='front';sheet=1;row=13};
        66=@{kind='source';row=44}; 67=@{kind='source';row=45}; 69=@{kind='diff';rowA=45;rowB=43;factor=100}; 71=@{kind='diff';rowA=50;rowB=46;factor=100};
        73=@{kind='source';row=53}; 75=@{kind='front';sheet=2;row=4}; 76=@{kind='front';sheet=2;row=5}; 77=@{kind='diff';rowA=56;rowB=54;factor=100};
        96=@{kind='source';row=87}; 98=@{kind='source';row=98}; 108=@{kind='source';row=63}; 116=@{kind='source';row=90}; 118=@{kind='source';row=91};
        119=@{kind='source';row=57}; 120=@{kind='source';row=58}; 121=@{kind='source';row=92}; 122=@{kind='source';row=94}; 124=@{kind='source';row=99}; 125=@{kind='source';row=60};
        127=@{kind='source';row=110}; 129=@{kind='source';row=59}; 130=@{kind='source';row=62}; 131=@{kind='index';row=2}; 132=@{kind='source';row=61}
    }

    $matched = 0
    $unmatched = New-Object System.Collections.Generic.List[string]
    $anchors = ''
    $chartRelEntries = ''
    $chartOverrides = ''
    $chartIndex = 1

    for ($r = 6; $r -le 133; $r++) {
        $indicator = (Get-Cell $summaryCells 'I' $r).Trim()
        if ([string]::IsNullOrWhiteSpace($indicator)) { continue }
        $freq = Get-Cell $summaryCells 'L' $r
        $record = $null
        if ($rowMap.ContainsKey($r)) {
            $m = $rowMap[$r]
            switch ($m.kind) {
                'front' { $record = Make-FrontRecord $frontSheets ([int]$m.sheet) ([int]$m.row) }
                'source' { $record = Make-SourceRecord $macroCells ([int]$m.row) $freq }
                'index' { $record = Make-SourceRecord $indexCells ([int]$m.row) $freq }
                'diff' { $record = Make-ComputedDiffRecord $macroCells ([int]$m.rowA) ([int]$m.rowB) ([double]$m.factor) $freq }
            }
        }
        $cellsToAdd = @{}
        if ($null -eq $record -or [string]::IsNullOrWhiteSpace([string]$record.Current)) {
            foreach ($col in @('O','P','Q','R')) {
                $ref = "$col$r"
                $style = Get-StyleAttr $sheetText $ref ($col + '6')
                $cellsToAdd[$ref] = Make-CellXml $ref $dash $style
            }
            $unmatched.Add("$r $indicator") | Out-Null
        } else {
            $matched++
            foreach ($pair in @(@('O',$record.Current), @('P',$record.Date), @('Q',$record.Change))) {
                $ref = "$($pair[0])$r"
                $style = Get-StyleAttr $sheetText $ref ($pair[0] + '6')
                $valueForStyle = Convert-ValueForStyle $pair[1] $style $percentStyleIds
                $cellsToAdd[$ref] = Make-CellXml $ref $valueForStyle $style
            }
            $trend = @($record.Trend | Where-Object { Is-NumberText ([string]$_) } | Select-Object -Last 60)
            if ($trend.Count -gt 1) {
                $startCol = 27
                for ($i = 0; $i -lt $trend.Count; $i++) {
                    $col = Get-ColumnName ($startCol + $i)
                    $ref = "$col$r"
                    $cellsToAdd[$ref] = Make-CellXml $ref $trend[$i]
                }
                $endCol = Get-ColumnName ($startCol + $trend.Count - 1)
                $formula = $sheetNameFormula + '!$AA$' + $r + ':$' + $endCol + '$' + $r
                $chartPath = "xl/charts/chart$chartIndex.xml"
                Set-ZipEntryText $summaryZip $chartPath (Make-ChartXml $formula $trend)
                $rid = "rId" + ($chartIndex + 2)
                $chartRelEntries += "<Relationship Id=`"$rid`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart`" Target=`"../charts/chart$chartIndex.xml`"/>"
                $chartOverrides += "<Override PartName=`"/xl/charts/chart$chartIndex.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.drawingml.chart+xml`"/>"
                $height = if ($rowHeights.ContainsKey($r)) { $rowHeights[$r] } else { 36.0 }
                $cy = [int]([math]::Max(380000, [math]::Min(1050000, $height * 11500)))
                $anchors += Make-ChartAnchor $r $rid 2500000 $cy
                $chartIndex++
            } else {
                $ref = "R$r"
                $style = Get-StyleAttr $sheetText $ref 'R6'
                $cellsToAdd[$ref] = Make-CellXml $ref $dash $style
            }
        }
        $sheetText = Replace-Or-Insert-RowCells $sheetText $r $cellsToAdd
    }

    $sheetText = [regex]::Replace($sheetText, '<dimension ref="[^"]+"', '<dimension ref="A1:CB134"')
    if ($sheetText -match '<cols>') {
        $sheetText = $sheetText -replace '</cols>', '<col min="27" max="80" width="0" hidden="1" customWidth="1"/></cols>'
    } else {
        $sheetText = $sheetText -replace '<sheetData>', '<cols><col min="27" max="80" width="0" hidden="1" customWidth="1"/></cols><sheetData>'
    }
    Set-ZipEntryText $summaryZip $summarySheet.Path $sheetText

    $drawingPath = Get-DrawingPathForSheet $summaryZip $summarySheet.Path
    if ($null -eq $drawingPath) { throw 'Summary drawing not found.' }
    $drawingText = Get-ZipEntryText $summaryZip $drawingPath
    $drawingText = $drawingText -replace '</xdr:wsDr>\s*$', ($anchors + '</xdr:wsDr>')
    Set-ZipEntryText $summaryZip $drawingPath $drawingText

    $drawingRelsPath = ([System.IO.Path]::GetDirectoryName($drawingPath).Replace('\', '/')) + '/_rels/' + ([System.IO.Path]::GetFileName($drawingPath)) + '.rels'
    $drawingRelsText = Get-ZipEntryText $summaryZip $drawingRelsPath
    $drawingRelsText = $drawingRelsText -replace '</Relationships>\s*$', ($chartRelEntries + '</Relationships>')
    Set-ZipEntryText $summaryZip $drawingRelsPath $drawingRelsText

    $ctText = Get-ZipEntryText $summaryZip '[Content_Types].xml'
    $ctText = $ctText -replace '</Types>\s*$', ($chartOverrides + '</Types>')
    Set-ZipEntryText $summaryZip '[Content_Types].xml' $ctText

    $report = [pscustomobject]@{
        Daily = $daily.FullName
        Summary = $summary.FullName
        Output = $output
        Matched = $matched
        UnmatchedCount = $unmatched.Count
        Unmatched = @($unmatched)
        Charts = ($chartIndex - 1)
    }
    $report | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $summaryZip) { $summaryZip.Dispose() }
    if ($null -ne $dailyZip) { $dailyZip.Dispose() }
}
