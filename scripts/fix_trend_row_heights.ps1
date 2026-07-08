$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$source = Get-ChildItem -File -Filter "*add12macro_trendfixed.xlsx" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $source) { throw "trendfixed workbook not found." }

$output = Join-Path $source.DirectoryName (($source.BaseName + "_rowfixed") + $source.Extension)

function Read-EntryText($zip, [string]$name) {
    $entry = $zip.GetEntry($name)
    if ($null -eq $entry) { throw "Missing zip entry: $name" }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Write-TextEntry($zip, [string]$name, [string]$text) {
    $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        try { $writer.Write($text) } finally { $writer.Dispose() }
    } finally { $stream.Dispose() }
}

function Fix-RowHeight([string]$xml, [int]$rowNum) {
    $pattern = '<row\b(?=[^>]*\sr="' + $rowNum + '")[^>]*>'
    $m = [regex]::Match($xml, $pattern)
    if (-not $m.Success) { return $xml }
    $tag = $m.Value
    $tag = [regex]::Replace($tag, '\sht="[^"]*"', '')
    $tag = [regex]::Replace($tag, '\scustomHeight="[^"]*"', '')
    $tag = $tag.Substring(0, $tag.Length - 1) + ' ht="39.75" customHeight="1">'
    return $xml.Substring(0, $m.Index) + $tag + $xml.Substring($m.Index + $m.Length)
}

$inFs = [System.IO.File]::Open($source.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
$inZip = New-Object System.IO.Compression.ZipArchive($inFs, [System.IO.Compression.ZipArchiveMode]::Read)
$temp = $output + ".tmp"
if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
$outFs = [System.IO.File]::Open($temp, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
$outZip = New-Object System.IO.Compression.ZipArchive($outFs, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    $sheet3 = Read-EntryText $inZip "xl/worksheets/sheet3.xml"
    foreach ($r in 24..35) { $sheet3 = Fix-RowHeight $sheet3 $r }

    foreach ($entry in $inZip.Entries) {
        if ($entry.FullName -eq "xl/worksheets/sheet3.xml") {
            Write-TextEntry $outZip $entry.FullName $sheet3
        } else {
            $newEntry = $outZip.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $src = $entry.Open()
            $dst = $newEntry.Open()
            try { $src.CopyTo($dst) } finally { $dst.Dispose(); $src.Dispose() }
        }
    }
} finally {
    $outZip.Dispose()
    $outFs.Dispose()
    $inZip.Dispose()
    $inFs.Dispose()
}

Move-Item -LiteralPath $temp -Destination $output -Force
Write-Output "RowFixedPath=$output"
