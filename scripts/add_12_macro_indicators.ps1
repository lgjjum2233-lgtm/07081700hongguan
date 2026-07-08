$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$source = Get-ChildItem -File -Filter "20260526*.xlsx" |
    Where-Object { $_.Name -notlike "*add12macro*" } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $source) { throw "Source daily workbook not found." }

$indicatorFile = Get-ChildItem -File -Filter "*.xlsx" |
    Where-Object { $_.Length -lt 20000 } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $indicatorFile) { throw "Indicator workbook not found." }

$outputName = [System.IO.Path]::GetFileNameWithoutExtension($source.Name) + "_add12macro.xlsx"
$outputPath = Join-Path $source.DirectoryName $outputName
if (Test-Path -LiteralPath $outputPath) {
    $stamp = Get-Date -Format "HHmmss"
    $outputName = [System.IO.Path]::GetFileNameWithoutExtension($source.Name) + "_add12macro_$stamp.xlsx"
    $outputPath = Join-Path $source.DirectoryName $outputName
}
Copy-Item -LiteralPath $source.FullName -Destination $outputPath

$monthly = [string]([char]0x6708) + [string]([char]0x5EA6)
$yiYuan = [string]([char]0x4EBF) + [string]([char]0x5143)

$app = New-Object -ComObject ket.Application
$app.Visible = $false
$app.DisplayAlerts = $false

$newWb = $null
$wb = $null
try {
    $newWb = $app.Workbooks.Open($indicatorFile.FullName)
    $newSheet = $newWb.Worksheets.Item(1)
    $items = @()
    for ($r = 2; $r -le 13; $r++) {
        $name = [string]$newSheet.Cells.Item($r, 2).Text
        $code = [string]$newSheet.Cells.Item($r, 3).Text
        if (-not [string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($code)) {
            $items += [pscustomobject]@{ Name = $name; Code = $code }
        }
    }
    if ($items.Count -ne 12) { throw "Expected 12 indicators, got $($items.Count)." }
    $newWb.Close($false)
    $newWb = $null

    $wb = $app.Workbooks.Open($outputPath)
    $front = $wb.Worksheets.Item(3)
    $data = $wb.Worksheets.Item(4)
    $aux = $wb.Worksheets.Item(7)

    $dataName = [string]$data.Name
    $auxName = [string]$aux.Name

    $lastDataRow = 0
    for ($r = 1000; $r -ge 1; $r--) {
        if (-not [string]::IsNullOrWhiteSpace([string]$data.Cells.Item($r, 5).Text) -or
            -not [string]::IsNullOrWhiteSpace([string]$data.Cells.Item($r, 4).Text)) {
            $lastDataRow = $r
            break
        }
    }
    if ($lastDataRow -lt 2) { throw "Cannot detect source data rows." }
    $firstDataRow = $lastDataRow + 1
    $lastNewDataRow = $lastDataRow + $items.Count

    $data.Range("A$lastDataRow:SZ$lastDataRow").Copy()
    $data.Range("A$firstDataRow:SZ$lastNewDataRow").PasteSpecial(-4122)
    $app.CutCopyMode = $false
    $data.Range("A$firstDataRow:C$lastNewDataRow").ClearContents()
    $data.Range("F$firstDataRow:SZ$lastNewDataRow").ClearContents()

    for ($i = 0; $i -lt $items.Count; $i++) {
        $row = $firstDataRow + $i
        $data.Cells.Item($row, 4).Value2 = $items[$i].Name
        $data.Cells.Item($row, 5).Value2 = $items[$i].Code
    }

    $f2 = [string]$data.Range("F2").Formula
    if ([string]::IsNullOrWhiteSpace($f2)) {
        $f2 = '=wsd(E2:E' + $lastNewDataRow + ',C2,A2,B2,"TradingCalendar=SSE","PriceAdj=","rptType=1","Direction=V","Version=1","ShowParams=Y","cols=520;rows=' + ($lastNewDataRow - 1) + '")'
    } else {
        $f2 = $f2 -replace 'E2:E\d+', ('E2:E' + $lastNewDataRow)
        $f2 = $f2 -replace 'rows=\d+', ('rows=' + ($lastNewDataRow - 1))
    }
    $data.Range("F2").Formula = $f2

    $lastAuxRow = 0
    for ($r = 1000; $r -ge 1; $r--) {
        if (-not [string]::IsNullOrWhiteSpace([string]$aux.Cells.Item($r, 2).Formula) -or
            -not [string]::IsNullOrWhiteSpace([string]$aux.Cells.Item($r, 2).Text)) {
            $lastAuxRow = $r
            break
        }
    }
    if ($lastAuxRow -lt 1) { throw "Cannot detect auxiliary rows." }
    $firstAuxRow = $lastAuxRow + 1
    $lastNewAuxRow = $lastAuxRow + $items.Count

    $aux.Range("A$lastAuxRow:BI$lastAuxRow").Copy()
    $aux.Range("A$firstAuxRow:BI$lastNewAuxRow").PasteSpecial(-4122)
    $app.CutCopyMode = $false

    function SheetRef([string]$name) {
        return "'" + $name.Replace("'", "''") + "'!"
    }
    $dataRef = SheetRef $dataName
    $auxRef = SheetRef $auxName

    for ($i = 0; $i -lt $items.Count; $i++) {
        $srcRow = $firstDataRow + $i
        $auxRow = $firstAuxRow + $i
        $range = $dataRef + '$F$' + $srcRow + ':$SZ$' + $srcRow
        $firstCell = $dataRef + '$F$' + $srcRow
        for ($c = 2; $c -le 61; $c++) {
            $offset = $c - 1
            $idx = 'INDEX(' + $range + ',1,LOOKUP(2,1/(' + $range + '<>""),COLUMN(' + $range + '))-COLUMN(' + $firstCell + ')+1-60+' + $offset + ')'
            $aux.Cells.Item($auxRow, $c).Formula = '=IF(ISERROR(' + $idx + '),NA(),IF(' + $idx + '=0,NA(),' + $idx + '))'
        }
    }

    $lastFrontRow = 0
    for ($r = 300; $r -ge 1; $r--) {
        if (-not [string]::IsNullOrWhiteSpace([string]$front.Cells.Item($r, 1).Text) -or
            -not [string]::IsNullOrWhiteSpace([string]$front.Cells.Item($r, 5).Formula) -or
            -not [string]::IsNullOrWhiteSpace([string]$front.Cells.Item($r, 5).Text)) {
            $lastFrontRow = $r
            break
        }
    }
    if ($lastFrontRow -lt 2) { throw "Cannot detect front page rows." }
    $firstFrontRow = $lastFrontRow + 1
    $lastNewFrontRow = $lastFrontRow + $items.Count

    $front.Range("A$lastFrontRow:J$lastFrontRow").Copy()
    $front.Range("A$firstFrontRow:J$lastNewFrontRow").PasteSpecial(-4122)
    $app.CutCopyMode = $false
    $front.Range("A$firstFrontRow:J$lastNewFrontRow").RowHeight = $front.Range("A$lastFrontRow:J$lastFrontRow").RowHeight

    for ($i = 0; $i -lt $items.Count; $i++) {
        $frontRow = $firstFrontRow + $i
        $srcRow = $firstDataRow + $i
        $auxRow = $firstAuxRow + $i
        $unit = "%"
        if ($i -le 6) { $unit = $yiYuan }

        $front.Cells.Item($frontRow, 1).Value2 = $items[$i].Name
        $front.Cells.Item($frontRow, 2).Value2 = $monthly
        $front.Cells.Item($frontRow, 3).Value2 = $items[$i].Code
        $front.Cells.Item($frontRow, 4).Value2 = $unit

        $series = $dataRef + '$F$' + $srcRow + ':$SZ$' + $srcRow
        $dates = $dataRef + '$F$1:$SZ$1'
        $firstCell = $dataRef + '$F$' + $srcRow
        $front.Cells.Item($frontRow, 5).Formula = '=IFERROR(LOOKUP(9.99E+307,' + $series + '),"")'
        $front.Cells.Item($frontRow, 6).Formula = '=IFERROR(LOOKUP(2,1/(' + $series + '<>""),' + $dates + '),"")'
        $front.Cells.Item($frontRow, 7).Formula = '=IFERROR(E' + $frontRow + '-H' + $frontRow + ',"")'
        $front.Cells.Item($frontRow, 8).Formula = '=IFERROR(LOOKUP(9.99E+307,OFFSET(' + $firstCell + ',0,0,1,LOOKUP(2,1/(' + $series + '<>""),COLUMN(' + $series + '))-COLUMN(' + $firstCell + '))),"")'
        $front.Cells.Item($frontRow, 9).Formula = '=IFERROR(LOOKUP(2,1/(' + $series + '<>"")/(' + $dates + '<F' + $frontRow + '),' + $dates + '),"")'

        try { $front.Range("J$frontRow").ClearContents() } catch {}
        try { $front.Range("J$frontRow").SparklineGroups.Clear() } catch {}
        try {
            $sg = $front.Range("J$frontRow").SparklineGroups.Add(1, $auxRef + 'B' + $auxRow + ':BI' + $auxRow)
            try { $sg.SeriesColor.Color = 255 } catch {}
        } catch {
            $front.Cells.Item($frontRow, 10).Formula = '=HYPERLINK("#' + $auxName + '!B' + $auxRow + ':BI' + $auxRow + '","trend")'
        }
    }

    try { $wb.RefreshAll() } catch {}
    try { $app.CalculateFullRebuild() } catch { try { $app.CalculateFull() } catch { try { $app.Calculate() } catch {} } }
    $wb.Save()
    $wb.Close($true)
    $wb = $null
} finally {
    if ($null -ne $newWb) { try { $newWb.Close($false) } catch {} }
    if ($null -ne $wb) { try { $wb.Close($true) } catch {} }
    try { $app.DisplayAlerts = $true } catch {}
    try { $app.Quit() } catch {}
}

Write-Output "Source=$($source.Name)"
Write-Output "IndicatorFile=$($indicatorFile.Name)"
Write-Output "OutputPath=$outputPath"
Write-Output "AddedIndicators=12"
Write-Output "DataRows=$firstDataRow-$lastNewDataRow"
Write-Output "FrontRows=$firstFrontRow-$lastNewFrontRow"
Write-Output "AuxRows=$firstAuxRow-$lastNewAuxRow"
