Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

function U([int[]]$codes) {
    return -join ($codes | ForEach-Object { [char]$_ })
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

function SheetRef([string]$name) {
    return "'" + $name.Replace("'", "''") + "'!"
}

function HasAny([string]$text, [string[]]$words) {
    foreach ($w in $words) {
        if (-not [string]::IsNullOrWhiteSpace($w) -and $text.Contains($w)) { return $true }
    }
    return $false
}

function Infer-Frequency([string]$name) {
    $day = U @(0x65E5,0x5EA6)
    $week = U @(0x5468,0x5EA6)
    $month = U @(0x6708,0x5EA6)
    $quarter = U @(0x5B63,0x5EA6)
    $irregular = U @(0x4E0D,0x5B9A,0x671F)

    $weekly = @((U @(0x5F53,0x5468)), 'SOFR', 'ONRRP')
    $quarterly = @('GDP')
    $daily = @(
        (U @(0x6536,0x76CA,0x7387)), (U @(0x5229,0x7387)), (U @(0x6C47,0x7387)),
        (U @(0x7F8E,0x5143,0x6307,0x6570)), (U @(0x6CE2,0x52A8,0x7387)),
        (U @(0x539F,0x6CB9)), (U @(0x9006,0x56DE,0x8D2D)), 'DR001', 'DR007', 'R001', 'R007',
        'VIX', 'ETF'
    )
    if (HasAny $name $quarterly) { return $quarter }
    if (HasAny $name $weekly) { return $week }
    if (HasAny $name $daily) { return $day }
    if ($name.Contains((U @(0x62CD,0x5356)))) { return $irregular }
    return $month
}

function Infer-Unit([string]$name, [string]$category) {
    $point = U @(0x70B9)
    $yiYuan = U @(0x4EBF,0x5143)
    $usd100m = U @(0x4EBF,0x7F8E,0x5143)
    $pctWords = @(
        (U @(0x540C,0x6BD4)), (U @(0x73AF,0x6BD4)), (U @(0x5229,0x7387)),
        (U @(0x6536,0x76CA,0x7387)), (U @(0x5931,0x4E1A,0x7387)),
        (U @(0x5229,0x5DEE)), 'PMI', 'PCE', 'CPI', 'PPI', 'IORB', 'SOFR'
    )
    $amountWords = @(
        (U @(0x6295,0x653E)), (U @(0x878D,0x8D44)), (U @(0x8D37,0x6B3E)),
        (U @(0x5B58,0x6B3E)), (U @(0x9500,0x552E,0x989D)),
        (U @(0x6210,0x4EA4,0x989D)), (U @(0x8D64,0x5B57)),
        (U @(0x8D44,0x4EA7)), (U @(0x8D1F,0x503A)), (U @(0x50A8,0x5907)),
        (U @(0x91D1,0x989D))
    )
    if (HasAny $name $pctWords) { return '%' }
    if ($name.Contains((U @(0x6307,0x6570)))) { return $point }
    if (HasAny $name $amountWords) {
        if ($category -eq 'overseas') { return $usd100m }
        return $yiYuan
    }
    return ''
}

function Get-LastUsedRow($sheet, [int]$maxRow, [int[]]$cols) {
    for ($r = $maxRow; $r -ge 1; $r--) {
        foreach ($c in $cols) {
            $txt = [string]$sheet.Cells.Item($r, $c).Text
            $formula = [string]$sheet.Cells.Item($r, $c).Formula
            if (-not [string]::IsNullOrWhiteSpace($txt) -or -not [string]::IsNullOrWhiteSpace($formula)) {
                return $r
            }
        }
    }
    return 1
}

function Get-ExistingCodes($sheet) {
    $codes = @{}
    $last = Get-LastUsedRow $sheet 2000 @(1,3,5)
    for ($r = 2; $r -le $last; $r++) {
        $code = ([string]$sheet.Cells.Item($r, 3).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($code)) { $codes[$code] = $true }
    }
    return $codes
}

function Category-For([int]$row, [string]$name, [string]$code) {
    $us = U @(0x7F8E,0x56FD)
    $usd = U @(0x7F8E,0x5143)
    $sp = U @(0x6807,0x51C6,0x666E,0x5C14)
    $global = U @(0x5168,0x7403)
    $brent = U @(0x5E03,0x4F26,0x7279)
    $gold = U @(0x9EC4,0x91D1)
    $hkex = U @(0x6E2F,0x4EA4,0x6240)
    $stockAccount = U @(0x80A1,0x7968,0x8D26,0x6237)
    $bondWords = @(
        (U @(0x56FD,0x503A)), (U @(0x56FD,0x5F00,0x503A)), (U @(0x4F01,0x4E1A,0x503A)),
        (U @(0x57CE,0x6295,0x503A)), (U @(0x540C,0x4E1A,0x5B58,0x5355)),
        'DR001', 'DR007', 'R001', 'R007'
    )
    if ($name.Contains($us) -or $name.Contains($usd) -or $name.Contains($sp) -or $name.Contains($global) -or $name.Contains($brent) -or $name.Contains($gold)) { return 'overseas' }
    if ($name.Contains($hkex) -or $name.Contains($stockAccount)) { return 'stock' }
    if (HasAny $name $bondWords) { return 'bond' }
    return 'china'
}

function Add-TrendChart($sheet, [int]$frontRow, $dataSheet, [int]$srcRow) {
    try {
        $cell = $sheet.Cells.Item($frontRow, 10)
        $left = [double]$cell.Left + 2
        $top = [double]$cell.Top + 2
        $width = [math]::Max(40, [double]$cell.Width - 4)
        $height = [math]::Max(18, [double]$cell.Height - 4)
        $co = $sheet.ChartObjects().Add($left, $top, $width, $height)
        $chart = $co.Chart
        $chart.ChartType = 4
        $chart.SetSourceData($dataSheet.Range("F$srcRow:SZ$srcRow"))
        try { $chart.HasLegend = $false } catch {}
        try { $chart.HasTitle = $false } catch {}
        try { $chart.Axes(1).Delete() } catch {}
        try { $chart.Axes(2).Delete() } catch {}
        try { $chart.ChartArea.Format.Line.Visible = 0 } catch {}
        try { $chart.ChartArea.Format.Fill.Visible = 0 } catch {}
        try { $chart.PlotArea.Format.Line.Visible = 0 } catch {}
        try { $chart.PlotArea.Format.Fill.Visible = 0 } catch {}
        try {
            $ser = $chart.SeriesCollection(1)
            $ser.MarkerStyle = -4142
            $ser.Format.Line.ForeColor.RGB = 255
            $ser.Format.Line.Weight = 1.25
        } catch {}
        return $true
    } catch {
        return $false
    }
}

function Add-DisplayRows($sheet, [object[]]$items, $dataSheet, [string]$dataRef) {
    if ($items.Count -eq 0) {
        return [pscustomobject]@{ Added=0; Charts=0; FirstRow=0; LastRow=0 }
    }
    $lastRow = Get-LastUsedRow $sheet 5000 @(1,3,5)
    $startRow = $lastRow + 1
    $endRow = $lastRow + $items.Count
    $templateRow = if ($lastRow -ge 2) { $lastRow } else { 2 }
    try {
        $sheet.Range("A$templateRow:J$templateRow").Copy()
        $sheet.Range("A$startRow:J$endRow").PasteSpecial(-4122)
    } catch {}
    try { $sheet.Application.CutCopyMode = $false } catch {}
    try { $sheet.Range("A$startRow:J$endRow").RowHeight = 39.75 } catch {}
    $charts = 0
    for ($i = 0; $i -lt $items.Count; $i++) {
        $row = $startRow + $i
        $item = $items[$i]
        $srcRow = [int]$item.Row
        $sheet.Cells.Item($row, 1).Value2 = $item.Name
        $sheet.Cells.Item($row, 2).Value2 = $item.Frequency
        $sheet.Cells.Item($row, 3).Value2 = $item.Code
        $sheet.Cells.Item($row, 4).Value2 = $item.Unit
        $series = $dataRef + '$F$' + $srcRow + ':$SZ$' + $srcRow
        $dates = $dataRef + '$F$1:$SZ$1'
        $firstCell = $dataRef + '$F$' + $srcRow
        $sheet.Cells.Item($row, 5).Formula = '=IFERROR(LOOKUP(9.99E+307,' + $series + '),"")'
        $sheet.Cells.Item($row, 6).Formula = '=IFERROR(LOOKUP(2,1/(' + $series + '<>""),' + $dates + '),"")'
        $sheet.Cells.Item($row, 7).Formula = '=IFERROR(E' + $row + '-H' + $row + ',"")'
        $sheet.Cells.Item($row, 8).Formula = '=IFERROR(LOOKUP(9.99E+307,OFFSET(' + $firstCell + ',0,0,1,LOOKUP(2,1/(' + $series + '<>""),COLUMN(' + $series + '))-COLUMN(' + $firstCell + '))),"")'
        $sheet.Cells.Item($row, 9).Formula = '=IFERROR(LOOKUP(2,1/(' + $series + '<>"")/(' + $dates + '<F' + $row + '),' + $dates + '),"")'
        try { $sheet.Range("E$row").NumberFormat = 'General' } catch {}
        try { $sheet.Range("G$row:H$row").NumberFormat = 'General' } catch {}
        try { $sheet.Range("F$row").NumberFormat = 'yyyy-mm-dd' } catch {}
        try { $sheet.Range("I$row").NumberFormat = 'yyyy-mm-dd' } catch {}
        try { $sheet.Range("J$row").ClearContents() } catch {}
        if (Add-TrendChart $sheet $row $dataSheet $srcRow) { $charts++ }
    }
    return [pscustomobject]@{ Added=$items.Count; Charts=$charts; FirstRow=$startRow; LastRow=$endRow }
}

$source = Get-ChildItem -File -Filter '*add12macro*.xlsx' |
    Where-Object { $_.Name -notlike '~$*' -and $_.Length -ge 490000 } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $source) { throw 'Daily source workbook not found.' }

$v30 = Get-ChildItem -File -Filter '*V30*.xlsx' |
    Where-Object { $_.Name -notlike '~$*' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $v30) { throw 'V30 workbook not found.' }

$datePrefix = Get-Date -Format 'yyyyMMdd'
$outputBase = $datePrefix + '_daily_macro_expanded'
$outputPath = Get-VersionedPath $source.DirectoryName $outputBase $source.Extension
Copy-Item -LiteralPath $source.FullName -Destination $outputPath -Force

$overseasName = U @(0x6D77,0x5916,0x6570,0x636E)

$app = New-Object -ComObject ket.Application
$app.Visible = $false
$app.DisplayAlerts = $false

$v30Wb = $null
$wb = $null
try {
    $v30Wb = $app.Workbooks.Open($v30.FullName)
    $v30Sheet = $v30Wb.Worksheets.Item(1)
    $v30Items = @()
    for ($r = 2; $r -le 500; $r++) {
        $code = ([string]$v30Sheet.Cells.Item($r, 1).Text).Trim()
        $name = ([string]$v30Sheet.Cells.Item($r, 2).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($code) -and [string]::IsNullOrWhiteSpace($name)) { break }
        if (-not [string]::IsNullOrWhiteSpace($code) -and -not [string]::IsNullOrWhiteSpace($name)) {
            $v30Items += [pscustomobject]@{ Code=$code; Name=$name }
        }
    }
    $v30Wb.Close($false)
    $v30Wb = $null

    $wb = $app.Workbooks.Open($outputPath)
    $stockSheet = $wb.Worksheets.Item(1)
    $bondSheet = $wb.Worksheets.Item(2)
    $chinaSheet = $wb.Worksheets.Item(3)
    $dataSheet = $wb.Worksheets.Item(4)
    $dataRef = SheetRef ([string]$dataSheet.Name)

    $lastDataRow = Get-LastUsedRow $dataSheet 5000 @(4,5)
    $existingCodes = @{}
    for ($r = 2; $r -le $lastDataRow; $r++) {
        $code = ([string]$dataSheet.Cells.Item($r, 5).Text).Trim()
        if (-not [string]::IsNullOrWhiteSpace($code)) { $existingCodes[$code] = $true }
    }

    $newV30 = @($v30Items | Where-Object { -not $existingCodes.ContainsKey($_.Code) })
    if ($newV30.Count -gt 0) {
        $firstNew = $lastDataRow + 1
        $lastNew = $lastDataRow + $newV30.Count
        try {
            $dataSheet.Range("A$lastDataRow:SZ$lastDataRow").Copy()
            $dataSheet.Range("A$firstNew:SZ$lastNew").PasteSpecial(-4122)
        } catch {}
        try { $app.CutCopyMode = $false } catch {}
        try { $dataSheet.Range("A$firstNew:C$lastNew").ClearContents() } catch {}
        try { $dataSheet.Range("F$firstNew:SZ$lastNew").ClearContents() } catch {}
        for ($i = 0; $i -lt $newV30.Count; $i++) {
            $row = $firstNew + $i
            $dataSheet.Cells.Item($row, 4).Value2 = $newV30[$i].Name
            $dataSheet.Cells.Item($row, 5).Value2 = $newV30[$i].Code
        }
        $lastDataRow = $lastNew
    }

    $f2 = [string]$dataSheet.Range('F2').Formula
    if ([string]::IsNullOrWhiteSpace($f2)) {
        $f2 = '=wsd(E2:E' + $lastDataRow + ',C2,A2,B2,"TradingCalendar=SSE","PriceAdj=","rptType=1","Direction=V","Version=1","ShowParams=Y","cols=520;rows=' + ($lastDataRow - 1) + '")'
    } else {
        $f2 = $f2 -replace 'E2:E\d+', ('E2:E' + $lastDataRow)
        $f2 = $f2 -replace 'rows=\d+', ('rows=' + ($lastDataRow - 1))
    }
    $dataSheet.Range('F2').Formula = $f2

    for ($s = $wb.Worksheets.Count; $s -ge 1; $s--) {
        if ([string]$wb.Worksheets.Item($s).Name -eq $overseasName) {
            $wb.Worksheets.Item($s).Delete()
        }
    }
    $overSheet = $wb.Worksheets.Add($null, $chinaSheet)
    $overSheet.Name = $overseasName
    try {
        $chinaSheet.Range('A1:J1').Copy()
        $overSheet.Range('A1').PasteSpecial(-4104)
        $app.CutCopyMode = $false
    } catch {}
    for ($c = 1; $c -le 10; $c++) {
        try { $overSheet.Columns.Item($c).ColumnWidth = $chinaSheet.Columns.Item($c).ColumnWidth } catch {}
    }

    $stockCodes = Get-ExistingCodes $stockSheet
    $bondCodes = Get-ExistingCodes $bondSheet
    $chinaCodes = Get-ExistingCodes $chinaSheet
    $overCodes = @{}

    $targets = @{
        stock = New-Object System.Collections.Generic.List[object]
        bond = New-Object System.Collections.Generic.List[object]
        china = New-Object System.Collections.Generic.List[object]
        overseas = New-Object System.Collections.Generic.List[object]
    }

    for ($r = 2; $r -le $lastDataRow; $r++) {
        $name = ([string]$dataSheet.Cells.Item($r, 4).Text).Trim()
        $code = ([string]$dataSheet.Cells.Item($r, 5).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($code)) { continue }
        $cat = Category-For $r $name $code
        $freq = Infer-Frequency $name
        $unit = Infer-Unit $name $cat
        $item = [pscustomobject]@{ Row=$r; Name=$name; Code=$code; Frequency=$freq; Unit=$unit; Category=$cat }
        switch ($cat) {
            'stock' { if (-not $stockCodes.ContainsKey($code)) { $targets.stock.Add($item) | Out-Null; $stockCodes[$code] = $true } }
            'bond' { if (-not $bondCodes.ContainsKey($code)) { $targets.bond.Add($item) | Out-Null; $bondCodes[$code] = $true } }
            'overseas' { if (-not $overCodes.ContainsKey($code)) { $targets.overseas.Add($item) | Out-Null; $overCodes[$code] = $true } }
            default { if (-not $chinaCodes.ContainsKey($code)) { $targets.china.Add($item) | Out-Null; $chinaCodes[$code] = $true } }
        }
    }

    $stockResult = @(Add-DisplayRows $stockSheet $targets.stock.ToArray() $dataSheet $dataRef)[-1]
    $bondResult = @(Add-DisplayRows $bondSheet $targets.bond.ToArray() $dataSheet $dataRef)[-1]
    $chinaResult = @(Add-DisplayRows $chinaSheet $targets.china.ToArray() $dataSheet $dataRef)[-1]
    $overResult = @(Add-DisplayRows $overSheet $targets.overseas.ToArray() $dataSheet $dataRef)[-1]

    try { $app.Calculate() } catch {}
    $wb.Save()
    $wb.Close($true)
    $wb = $null

    [pscustomobject]@{
        Source = $source.FullName
        V30 = $v30.FullName
        Output = $outputPath
        V30Rows = $v30Items.Count
        NewMacroRows = $newV30.Count
        MacroLastRow = $lastDataRow
        StockAdded = $stockResult.Added
        BondAdded = $bondResult.Added
        ChinaAdded = $chinaResult.Added
        OverseasAdded = $overResult.Added
        ChartsAdded = $stockResult.Charts + $bondResult.Charts + $chinaResult.Charts + $overResult.Charts
    } | ConvertTo-Json -Depth 5
} finally {
    if ($null -ne $v30Wb) { try { $v30Wb.Close($false) } catch {} }
    if ($null -ne $wb) { try { $wb.Close($true) } catch {} }
    try { $app.DisplayAlerts = $true } catch {}
    try { $app.Quit() } catch {}
}
