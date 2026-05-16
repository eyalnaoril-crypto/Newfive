# Barak — PDF generator via Microsoft Word COM (PowerShell)
# Usage:
#   powershell.exe -ExecutionPolicy Bypass -File scripts/barak-pdf-via-word-com.ps1
#   powershell.exe -ExecutionPolicy Bypass -File scripts/barak-pdf-via-word-com.ps1 -Source "output/myfile.docx" -Target "output/myfile.pdf"
#
# Paths are resolved relative to the project root (one level above this script).
# Word COM requires absolute paths — they are computed automatically.

param(
    [string]$Source = "",
    [string]$Target = ""
)

$ErrorActionPreference = "Stop"

# Resolve project root (parent of the scripts/ directory)
$projectRoot = Split-Path -Parent $PSScriptRoot

# Default source: latest v1.1 docx in output/ (matches naming convention)
if (-not $Source) {
    $defaultPattern = Join-Path $projectRoot "output\*-fleet-report-yoy-q1-2026-v1.1.docx"
    $candidate = Get-ChildItem -Path $defaultPattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($candidate) {
        $Source = $candidate.FullName
    } else {
        throw "No default source found. Specify -Source <path-to-docx>."
    }
} elseif (-not [System.IO.Path]::IsPathRooted($Source)) {
    $Source = Join-Path $projectRoot $Source
}

# Default target: output/<timestamp>-fleet-report-yoy-q1-2026-v1.2.pdf
if (-not $Target) {
    $timestamp = Get-Date -Format "yyyyMMdd-HHmm"
    $Target = Join-Path $projectRoot "output\$timestamp-fleet-report-yoy-q1-2026-v1.2.pdf"
} elseif (-not [System.IO.Path]::IsPathRooted($Target)) {
    $Target = Join-Path $projectRoot $Target
}

Write-Output "Project root: $projectRoot"
Write-Output "Source:       $Source"
Write-Output "Target:       $Target"

if (-not (Test-Path $Source)) {
    throw "Source file not found: $Source"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0  # wdAlertsNone
$exportedOk = $false
try {
    $doc = $word.Documents.Open($Source, $false, $true)  # ConfirmConversions=false, ReadOnly=true
    $totalPages = $doc.ComputeStatistics(2)  # wdStatisticPages = 2
    Write-Output "Total pages: $totalPages"
    # wdExportFormatPDF = 17
    $doc.ExportAsFixedFormat($Target, 17, $false, 0, 0, 1, $totalPages, 0, $true, $true, 0, $true, $true, $false)
    $exportedOk = $true
    Write-Output "PDF created: $Target"
    # $doc.Close() is wrapped — Word COM occasionally disconnects after a clean export
    try { $doc.Close($false) } catch { Write-Verbose "Doc.Close cleanup error suppressed: $_" }
    if (Test-Path $Target) {
        $size = (Get-Item $Target).Length
        Write-Output "File size: $size bytes"
    }
} finally {
    try { $word.Quit() } catch { Write-Verbose "Word.Quit cleanup error suppressed: $_" }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# Exit cleanly if export succeeded (don't propagate cosmetic COM cleanup errors)
if ($exportedOk -and (Test-Path $Target)) { exit 0 } else { exit 1 }
