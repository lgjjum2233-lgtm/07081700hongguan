param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,
    [string]$OutputPath = "",
    [int]$WaitSeconds = 120
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$target = [System.IO.Path]::GetFullPath($WorkbookPath)
if (-not (Test-Path -LiteralPath $target)) {
    throw "Workbook not found: $target"
}

try {
    $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("ket.Application")
} catch {
    $app = New-Object -ComObject ket.Application
}

$app.Visible = $true
$app.DisplayAlerts = $false

$wb = $null
foreach ($openWb in @($app.Workbooks)) {
    if ([System.IO.Path]::GetFullPath($openWb.FullName) -eq $target) {
        $wb = $openWb
        break
    }
}
if ($null -eq $wb) {
    $wb = $app.Workbooks.Open($target)
}

$wb.Activate()
try { $wb.RefreshAll() } catch { Write-Output "RefreshAllError=$($_.Exception.Message)" }
try { $app.CalculateFullRebuild() } catch {
    try { $app.CalculateFull() } catch {
        try { $app.Calculate() } catch { Write-Output "CalculateError=$($_.Exception.Message)" }
    }
}

Start-Sleep -Seconds $WaitSeconds

try { $app.CalculateFullRebuild() } catch {
    try { $app.CalculateFull() } catch {
        try { $app.Calculate() } catch {}
    }
}

if ($OutputPath -and $OutputPath.Trim().Length -gt 0) {
    $out = [System.IO.Path]::GetFullPath($OutputPath)
    $wb.SaveCopyAs($out)
    Write-Output "SavedCopy=$out"
} else {
    $wb.Save()
    Write-Output "Saved=$($wb.FullName)"
}
