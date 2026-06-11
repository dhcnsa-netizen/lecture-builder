param([Parameter(Mandatory=$true)][string]$Path)
# Bake the table of contents & page-number fields into a .docx using Microsoft Word
# COM automation. Safe to skip if only WPS is installed (TOC still updates on open).
$ErrorActionPreference = "SilentlyContinue"
$full = (Resolve-Path $Path).Path
try {
    $w = New-Object -ComObject Word.Application
} catch {
    Write-Output "Microsoft Word COM not available; skipping (TOC will update on first open in WPS)."
    exit 0
}
$w.Visible = $false; $w.DisplayAlerts = 0
$doc = $w.Documents.Open($full)
if ($doc.TablesOfContents.Count -gt 0) { $doc.TablesOfContents.Item(1).Update() }
$doc.Fields.Update() | Out-Null
$doc.Repaginate()
$pages = $doc.ComputeStatistics(2)   # wdStatisticPages
$doc.Save()
$doc.Close($true)
$w.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($w) | Out-Null
Write-Output "Finalized $full ($pages pages)."
