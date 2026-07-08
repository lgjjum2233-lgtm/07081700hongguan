param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath,
    [int]$WaitSeconds = 90
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$inputFull = (Resolve-Path -LiteralPath $InputPath).Path
$outputFull = [System.IO.Path]::GetFullPath($OutputPath)

if (Test-Path -LiteralPath $outputFull) {
    Remove-Item -LiteralPath $outputFull -Force
}

$app = $null
$wb = $null
try {
    try {
        $app = New-Object -ComObject ket.Application
        $appName = "ket.Application"
    } catch {
        $app = New-Object -ComObject Excel.Application
        $appName = "Excel.Application"
    }

    $app.Visible = $false
    $app.DisplayAlerts = $false

    $wb = $app.Workbooks.Open($inputFull)
    try { $wb.RefreshAll() } catch {}

    $deadline = (Get-Date).AddSeconds($WaitSeconds)
    do {
        try { $app.CalculateFullRebuild() } catch {
            try { $app.CalculateFull() } catch {
                try { $app.Calculate() } catch {}
            }
        }
        Start-Sleep -Seconds 5
    } while ((Get-Date) -lt $deadline)

    try { $wb.SaveAs($outputFull) } catch {
        $wb.Save()
        Copy-Item -LiteralPath $inputFull -Destination $outputFull -Force
    }
    $wb.Close($true)
    $wb = $null
} finally {
    if ($null -ne $wb) { try { $wb.Close($true) } catch {} }
    if ($null -ne $app) {
        try { $app.DisplayAlerts = $true } catch {}
        try { $app.Quit() } catch {}
    }
}

Write-Output "App=$appName"
Write-Output "Input=$inputFull"
Write-Output "Output=$outputFull"
Write-Output "WaitSeconds=$WaitSeconds"
