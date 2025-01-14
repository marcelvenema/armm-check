param (
    [string]$new,
    [string]$old
)

Write-Output "Check ARMM wijzigingen in lijst..."
Write-Output "(p) MarcelVenema.com, 2024 for MedCor Pharmaceuticals B.V."
Write-Output "usage: check_cbg_lijst.ps1 -new CBG_LIJST_20250114.json -old CBG_CHECK_20250101.json"
Write-Output "."

# CBG_CHECK_<date>.json
$old_registraties = @() 
$old_registraties = Get-Content -Raw -Path $old | ConvertFrom-Json # CBG_CHECK_<date>.json

# CBG_LIJST_<date>.json
$registraties = @()
$registraties = Get-Content -Raw -Path $new | ConvertFrom-Json # CBG_LIJST_<date>.json
$datum = Get-Date -Format "dd-MM-yyyy" # NL date format for detection date
$date = Get-Date -Format "yyyyMMdd" # US date format for filename
Write-Output "Totaal $($registraties.Count) registraties gevonden..."

############################################################################################################
# Zoek de verschillen
############################################################################################################

ForEach ($registratie in $registraties) {

    Write-Output "Controle wijzigingen product $($registratie.PRODUCTNAAM)..."
    $reg_nummer = $registratie.REG_NUMMER_HANDELSVERGUNNINGHOUDER

    # $old_registratie = $old_registraties | Where-Object { $_.REG_ID -eq $reg_id }
    $old_registratie = $old_registraties | Where-Object { $_.REG_NUMMER_HANDELSVERGUNNINGHOUDER -eq $reg_nummer }
    $datum_gewijzigd = $false

    # ARMM CHECKSUM MERKHOUDER check
    If ($old_registratie.ARMM_CHECKSUM_MAH -ne $registratie.ARMM_CHECKSUM_MAH) {
        Write-Output "!!! ARMM_CHECKSUM_MERKHOUDER is veranderd voor $($registratie.PRODUCTNAAM)..."
        $registratie.ARMM_CHECKSUM_MAH_WIJZIGING = "Yes"
        $registratie.DATUM_DETECTIE = $datum
        $datum_gewijzigd = $true
    } else { If ($datum_gewijzigd -eq $false) { $registratie.DATUM_DETECTIE = $old_registratie.DATUM_DETECTIE } }

    # BIJSLUITER_CHECKSUM MERKHOUDER check
    If ($old_registratie.BIJSLUITER_CHECKSUM_MAH -ne $registratie.BIJSLUITER_CHECKSUM_MAH) {
        Write-Output "!!! BIJSLUITER_CHECKSUM MERKHOUDER is veranderd voor $($registratie.PRODUCTNAAM)..."
        $registratie.BIJSLUITER_CHECKSUM_MAH_WIJZIGING = "Yes"
        $registratie.DATUM_DETECTIE = $datum
        $datum_gewijzigd = $true
    } else { If ($datum_gewijzigd -eq $false) { $registratie.DATUM_DETECTIE = $old_registratie.DATUM_DETECTIE } }



    # SMPC_CHECKSUM MERKHOUDER check
    If ($old_registratie.SMPC_CHECKSUM_MAH -ne $registratie.SMPC_CHECKSUM_MAH) {
        Write-Output "!!! SMPC_CHECKSUM MERKHOUDER is veranderd voor $($registratie.PRODUCTNAAM)..."
        $registratie.SMPC_CHECKSUM_MAH_WIJZIGING = "Yes"
        $registratie.DATUM_DETECTIE = $datum
        $datum_gewijzigd = $true
    } else { If ($datum_gewijzigd -eq $false) { $registratie.DATUM_DETECTIE = $old_registratie.DATUM_DETECTIE } }

    # REG check
    if ($old_registratie -eq $null) {
        Write-Output "!!! Nieuwe registratie gevonden: $($registratie.PRODUCTNAAM)"
        $registratie.RVG_WIJZIGING = "Yes"
        $registratie.DATUM_DETECTIE = $datum
    }
}

# Save $registraties to a file
$registraties | ConvertTo-Json | Out-File -FilePath "CBG_CHECK_$date.json"
Write-Output "Controle bestand opgeslagen als CBG_CHECK_$date.json"

############################################################################################################
# Converteer naar Excel
############################################################################################################

Write-Output "Lijst wordt geconverteerd naar Excel. Een moment geduld aub..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Worksheets.Item(1)

# Get the property names from the first object to use as headers
$headers = $registraties[0].PSObject.Properties.Name
$row = 1

# Add headers to the first row
for ($col = 0; $col -lt $headers.Count; $col++) {
    $sheet.Cells.Item($row, $col + 1) = $headers[$col]
}

# Add the JSON data to the sheet
foreach ($item in $registraties) {
    $row++
    for ($col = 0; $col -lt $headers.Count; $col++) {
        $sheet.Cells.Item($row, $col + 1) = $item.($headers[$col])
    }
}

# Autofit the columns
$sheet.Columns.AutoFit()

# Add comments to sheet
$comment = "Alle geregistreerde RVG nummers van Medcor Pharmaceuticals.`nBron: CBG website.`nDit bestand is gegenereerd door het CBG_CHECK script op $($datum).`nVergelijking met bestand: $($registratie_file)."
$sheet.Cells.Item(1, 1).AddComment($comment) | Out-Null
$comment = "Het deel achter de dubbele schuine streep uit kolom A wordt in deze kolom gezet."
$sheet.Cells.Item(1, 2).AddComment($comment) | Out-Null
$comment = "De productnaam van het RVG nummer uit kolom A.`nBron: CBG website."
$sheet.Cells.Item(1, 3).AddComment($comment) | Out-Null
$comment = "URL naar product van kolom A op CBG website.`nBron: CBG website."
$sheet.Cells.Item(1, 4).AddComment($comment) | Out-Null
$comment = "URL naar product van kolom B op CBG website.(merkhouder)`nBron: CBG website."
$sheet.Cells.Item(1, 5).AddComment($comment) | Out-Null
$comment = "Toont verdere informatie op de CBG-site over datum intrekking van RVG nr uit kolom B.`nBron: CBG website."
$sheet.Cells.Item(1, 6).AddComment($comment) | Out-Null
$comment = "Geeft aan of er een wijziging is gevonden in kolommen A-F t.o.v. de vorige run."
$sheet.Cells.Item(1, 7).AddComment($comment) | Out-Null
$comment = "De URL die hoort bij het aRMM van het referentieproduct.`nBron: CBG website."
$sheet.Cells.Item(1, 8).AddComment($comment) | Out-Null
$comment = "Het bestand die hoort bij het aRMM van het referentieproduct.`nBron: CBG website."
$sheet.Cells.Item(1, 9).AddComment($comment) | Out-Null
$comment = "Unieke checksum welke het script genereerd bij het aRMM bestand van het referentieproduct."
$sheet.Cells.Item(1, 10).AddComment($comment) | Out-Null
$comment = "Geeft aan of er een wijziging is geweest in het aRMM bestand ten opzichte van bestand $($registratie_file)."
$sheet.Cells.Item(1, 11).AddComment($comment) | Out-Null
$comment = "De URL die hoort bij de patiëntenbijsluiter van het referentieproduct.`nBron: CBG website."
$sheet.Cells.Item(1, 12).AddComment($comment) | Out-Null
$comment = "Het bestand die hoort bij de patiëntenbijsluiter van het referentieproduct.`nBron: CBG website."
$sheet.Cells.Item(1, 13).AddComment($comment) | Out-Null
$comment = "Unieke checksum welke het script genereerd bij de patiëntenbijsluiter van het referentieproduct."
$sheet.Cells.Item(1, 14).AddComment($comment) | Out-Null
$comment = "Geeft aan of er een wijziging is geweest in de patiëntenbijsluiter van het referentieproduct ten opzichte van bestand $($registratie_file)."
$sheet.Cells.Item(1, 15).AddComment($comment) | Out-Null
$comment = "De URL die hoort bij de SmPC van het referentieproduct.`nBron: CBG website."
$sheet.Cells.Item(1, 16).AddComment($comment) | Out-Null
$comment = "Het bestand die hoort bij de SmPC van het referentieproduct.`nBron: CBG website."
$sheet.Cells.Item(1, 17).AddComment($comment) | Out-Null
$comment = "Unieke checksum welke het script genereerd bij de SmOC bestand van het referentieproduct."
$sheet.Cells.Item(1, 18).AddComment($comment) | Out-Null 
$comment = "Geeft aan of er een wijziging is geweest in de SmPC van het referentieproduct ten opzichte van bestand $($registratie_file)."
$sheet.Cells.Item(1, 19).AddComment($comment) | Out-Null
$comment = "Datum waarop het script de wijzigingen van het CBG heeft ontdekt."
$sheet.Cells.Item(1, 20).AddComment($comment) | Out-Null

# Get current script folder, not the file
$scriptFolder = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
# Save workbook
$workbook.SaveAs( $ScriptFolder + "\CBG_CHECK_$date.xlsx")
$excel.Quit()

# Destroy excel object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Output "Bestand CBG_CHECK_$date.xlsx is aangemaakt."
