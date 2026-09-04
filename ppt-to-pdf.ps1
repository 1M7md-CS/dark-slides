$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$pdfDir = Join-Path $root "pdf"
$pptDir = Join-Path $root "ppt"
$darkDir = Join-Path $root "pdf-dark"
$convertScript = Join-Path $root "convert-dark.js"

New-Item -ItemType Directory -Path $pdfDir, $pptDir, $darkDir -Force | Out-Null

$files = Get-ChildItem -Path $root -File |
    Where-Object { $_.Extension -in ".ppt", ".pptx" }

if ($files.Count -eq 0) {
    Write-Host "No PowerPoint files found."
    exit
}

Write-Host ""
Write-Host "PPT -> PDF"
Write-Host "----------"

$office = New-Object -ComObject PowerPoint.Application
$success = 0
$failed = 0
$total = $files.Count
$count = 0

foreach ($file in $files) {
    $count++
    $outFile = Join-Path $pdfDir "$($file.BaseName).pdf"

    try {
        $pres = $office.Presentations.Open($file.FullName)
        $pres.SaveAs($outFile, 32)
        $pres.Close()

        Move-Item -LiteralPath $file.FullName -Destination $pptDir -Force

        $success++
        Write-Host "[$count/$total] $($file.Name)  OK"
    }
    catch {
        $failed++
        Write-Host "[$count/$total] $($file.Name)  FAILED"

        if ($pres) {
            try {
                $pres.Close()
            }
            catch {
            }
        }
    }
}

$office.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($office) | Out-Null

Write-Host ""
Write-Host "$success converted, $failed failed"

$pdfs = Get-ChildItem -Path $pdfDir -Filter "*.pdf" -File

if ($pdfs.Count -eq 0) {
    Write-Host ""
    Write-Host "No PDFs found."
    exit
}

Write-Host ""
Write-Host "Dark Mode"
Write-Host "---------"

$darkSuccess = 0
$darkFailed = 0
$total = $pdfs.Count
$count = 0

foreach ($pdf in $pdfs) {
    $count++
    $darkFile = Join-Path $darkDir "$($pdf.BaseName)_dark.pdf"

    try {
        & node $convertScript $pdf.FullName $darkFile 2>$null | Out-Null

        if ($LASTEXITCODE -ne 0) {
            throw "Node exited with code $LASTEXITCODE"
        }

        $darkSuccess++
        Write-Host "[$count/$total] $($pdf.Name)  OK"
    }
    catch {
        $darkFailed++
        Write-Host "[$count/$total] $($pdf.Name)  FAILED"
    }
}

Write-Host ""
Write-Host "$darkSuccess converted, $darkFailed failed"

Write-Host ""
Write-Host "Output"
Write-Host "------"
Write-Host "PDF:      $pdfDir"
Write-Host "Dark PDF: $darkDir"
Write-Host "Original: $pptDir"

Write-Host ""
Write-Host "Done."