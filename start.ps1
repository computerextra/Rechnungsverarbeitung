$host.UI.RawUI.WindowTitle = "Rechnungsverarbeitung"

$RECHNUNGEN = "V:\SERVICE\KYO_SCANS\RECHNUNGEN"
$DESTINATION = "\\compex4\SCANS\unbearbeitet\"
# Erst Scoop + Ghostscript installieren, dann das hier ausführen

while ($true) {
    Clear-Host
    # Finde alle PDF Dateien in Ordner:
    $pdfs = Get-ChildItem $RECHNUNGEN -Recurse | Where-Object { $_.Extension -match "pdf" }

    foreach ($pdf in $pdfs) {
        "PDF: " + $pdf
        $pdfArray = $pdf.FullName.Split(".")
        $tiff = ""
        if ($pdfArray.Length -gt 2) {
            $i = 0
            $end = $pdfArray.Length
            $end -= 1
            while ($i -lt $end) {
                $tiff += $pdfArray[$i]
                $tiff += "_"
                $i++
            }
        }
        else {
            $tiff = $pdfArray[0]
        }
        $tiff = $tiff -replace "\s+", "_"
        $tiff += ".tiff"
        if (Test-Path $tiff) {
            "tiff file already exists " + $tiff
        }
        else {
            "Processing " + $pdf.Name
            $param = "-sOutputFile=$tiff"
            gs -q  -dNOPAUSE -sDEVICE=tiffgray $param -r200x200 $pdf.FullName -c quit
        }
    }
    Clear-Host
    Write-Host "Lösche alle pdf Dateien"
    # Remove-Item $RECHNUNGEN/*.pdf
    Write-Host "Verschiebe alle Tiff Dateien"
    # Move-Item /y $RECHNUNGEN/*.tiff $DESTINATION
    $timeout = 1 # Timeout in Stunden
    $timeout = $timeout * 60 * 60 # Timeout in Sekunden
    Timeout /T $timeout
}