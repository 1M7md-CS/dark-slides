$ErrorActionPreference = "Stop"
$folder = $PSScriptRoot
$pdfDir = Join-Path $folder "pdf"
$pptDir = Join-Path $folder "ppt"
$darkDir = Join-Path $folder "pdf-dark"
$convertScript = Join-Path $folder "convert-dark.js"

New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
New-Item -ItemType Directory -Path $pptDir -Force | Out-Null
New-Item -ItemType Directory -Path $darkDir -Force | Out-Null

$files = Get-ChildItem -Path "$folder\*" -Include *.ppt,*.pptx -File
Write-Host "Found $($files.Count) PPT files"

if ($files.Count -eq 0) {
    Write-Host "No PPT files found!"
    exit
}

$office = New-Object -ComObject PowerPoint.Application

foreach ($file in $files) {
    $outFile = Join-Path $pdfDir "$($file.BaseName).pdf"
    Write-Host "Converting: $($file.Name)"

    try {
        $pres = $office.Presentations.Open($file.FullName)
        $pres.SaveAs($outFile, 32)
        $pres.Close()
        Move-Item -LiteralPath $file.FullName -Destination $pptDir -Force
        Write-Host "  -> OK"
    } catch {
        Write-Host "  -> ERROR: $_"
    }
}

$office.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($office) | Out-Null

Write-Host "`n--- Dark Mode Conversion (Claude Warm) ---"

$pdfs = Get-ChildItem -Path $pdfDir -Filter *.pdf -File
foreach ($pdf in $pdfs) {
    $darkFile = Join-Path $darkDir "$($pdf.BaseName)_dark.pdf"
    Write-Host "Dark mode: $($pdf.Name)"

    try {
        & node $convertScript $pdf.FullName $darkFile
    } catch {
        Write-Host "  -> ERROR: $_"
    }
}

Write-Host "`nDone. Dark PDFs saved to: $darkDir"
