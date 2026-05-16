$ErrorActionPreference = "Stop"
$source = "C:\Users\user.EYAL-NAOR-L\Documents\claude\working_area\Newfive\output\20260516-0951-fleet-report-yoy-q1-2026-v1.1.docx"
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$target = "C:\Users\user.EYAL-NAOR-L\Documents\claude\working_area\Newfive\output\$timestamp-fleet-report-yoy-q1-2026-v1.2.pdf"

Write-Output "Source: $source"
Write-Output "Target: $target"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone
try {
    $doc = $word.Documents.Open($source, $false, $true)  # ConfirmConversions=false, ReadOnly=true
    $totalPages = $doc.ComputeStatistics(2)  # wdStatisticPages = 2
    Write-Output "Total pages: $totalPages"
    # wdExportFormatPDF = 17
    $doc.ExportAsFixedFormat($target, 17, $false, 0, 0, 1, $totalPages, 0, $true, $true, 0, $true, $true, $false)
    $doc.Close($false)  # don't save changes
    Write-Output "PDF created: $target"
    if (Test-Path $target) {
        $size = (Get-Item $target).Length
        Write-Output "File size: $size bytes"
    }
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
