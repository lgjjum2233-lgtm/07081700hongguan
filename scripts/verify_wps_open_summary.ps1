Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$target = Get-ChildItem -File -Filter '*_matched*.xlsx' |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $target) { throw 'Matched workbook not found.' }
$path = $target.FullName
$app = $null
$progId = $null
foreach ($id in @('KET.Application', 'Excel.Application')) {
    try {
        $app = New-Object -ComObject $id
        $progId = $id
        break
    } catch {
        $app = $null
    }
}
if ($null -eq $app) { throw 'No spreadsheet COM application is available.' }

$wb = $null
try {
    $app.Visible = $false
    try { $app.DisplayAlerts = $false } catch {}
    $wb = $app.Workbooks.Open($path)
    $ws = $wb.Worksheets.Item(1)
    $chartCount = $ws.ChartObjects().Count
    $o6 = $ws.Range('O6').Text
    $p6 = $ws.Range('P6').Text
    $q6 = $ws.Range('Q6').Text
    $r78 = $ws.Range('R78').Text
    [pscustomobject]@{
        Program = $progId
        Workbook = $path
        Sheet = $ws.Name
        ChartObjects = $chartCount
        O6 = $o6
        P6 = $p6
        Q6 = $q6
        R78 = $r78
    } | ConvertTo-Json
    $wb.Close($false)
} finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $app) { try { $app.Quit() } catch {} }
}
