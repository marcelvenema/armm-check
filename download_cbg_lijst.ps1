param (
    [switch]$download
)

Write-Output "Download ARMM wijzigingen CBG website..."
Write-Output "(p) MarcelVenema.com, 2024 for MedCor Pharmaceuticals B.V."
Write-Output "."

# Define the URL of handelsvergunninghouder. In this case this is MedCor Pharmaceuticals B.V.
$url = "https://www.geneesmiddeleninformatiebank.nl/ords/f?p=111:2::SEARCH:::P0_DOMAIN,P0_LANG,P2_QS,P2_AS_PROD,P2_AS_RVGNR,P2_AS_EU1,P2_AS_EU2,P2_AS_ACTSUB,P2_AS_INACTSUB,P2_AS_NOTINACT,P2_AS_ADDM,P2_AS_ARMM,P2_AS_APPDATE,P2_AS_APPDATS,P2_AS_ATC,P2_AS_PHARM,P2_AS_MAH,P2_AS_ROUTE,P2_AS_AUTHS,P2_AS_TGTSP,P2_AS_INDIC,P2_AS_TXTF,P2_AS_TXTI,P2_AS_TXTC,P2_RESPAGE,P2_SORT,P2_RESPPG,P2_OPTIONS:H,NL,%5C%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,N,N,N,,,%5C%5C,%5C%5C,%5Cmedcor%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,%5C%5C,1,PRODA,558,N"

$datum = Get-Date -Format "dd-MM-yyyy" # Dutch day format, used in column DATUM_DETECTIE
$date = Get-Date -Format "yyyyMMdd" # US date format, used in file names

# $download = $true # true = keep download files, false = delete download files

# Create folder for downloads, use current script path + date
# only if $download is true
If ($download) { $download_folder = $date; If (-Not (Test-Path $download_folder)) { New-Item -ItemType Directory -Path $download_folder | Out-Null }}

# Download the HTML content
$htmlContent = curl --silent $url
# Extract the table content using regex
$regex = '<table summary="Search Results">(.*?)</table>'
$tableContent = [regex]::Match($htmlContent, $regex, [System.Text.RegularExpressions.RegexOptions]::Singleline)
# Save tableContent to a file for debug purposes
# $tableContent.Value | Out-File -FilePath "tableContent.xml"
# Extract all rows
$regex = '<tr>(.*?)</tr>'
$rows = [regex]::Matches($tableContent.Value, $regex, [System.Text.RegularExpressions.RegexOptions]::Singleline)
Write-Output "Totaal $($rows.Count-1) registraties gevonden..."

# Cycle through all $matches.value, extract the columns and save it to variable
$registraties = @()
for ($i=1; $i -lt $rows.Count; $i++) {
    $regex = '<td.*?>(.*?)</td>'
    $col = [regex]::Matches($rows[$i].Value, $regex, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $registratie = [ordered]@{}

    # KOLOM REG_NUMMER HANDELSVERGUNNINGHOUDER
    $reg_nummer = $col[0].Groups[1].Value
    $reg_nummer = $col[0].Groups[1].Value -replace '.*RVG', 'RVG'   # delete everything before 'RVG'
    $registratie["REG_NUMMER_HANDELSVERGUNNINGHOUDER"] = $reg_nummer

    # KOLOM REG_NUMMER_MAH (MERKHOUDER)
    $regex_id = '//(.+)'
    if ($registratie["REG_NUMMER_HANDELSVERGUNNINGHOUDER"] -match $regex_id) { $registratie["REG_NUMMER_MAH"] = $matches[1] }
    else { $registratie["REG_NUMMER_MAH"] = ""}
  
    # PRODUCTNAAM
    $productnaam = $col[1].Groups[1].Value
    $regex_naam = '<a[^>]*>(.*?)<\/a>'
    if ($productnaam -match $regex_naam) { $registratie["PRODUCTNAAM"] = $matches[1] } else { $registratie["PRODUCTNAAM"] = "" }

    # REG_URL_HANDELSVERGUNNINGHOUDER
    $regex_url = 'href="([^"]+)"'
    If ($col[1].Groups[1].Value -match $regex_url) { $registratie_url = [System.Net.WebUtility]::HtmlDecode($matches[1]) } else { $registratie_url = ""}
    If ($registratie_url) { $registratie["REG_URL_HANDELSVERGUNNINGHOUDER"] = "https://www.geneesmiddeleninformatiebank.nl/ords/" + $registratie_url }
    else { $registratie["REG_URL_HANDELSVERGUNNINGHOUDER"] = "" }

    # REG_URL_MERKHOUDER
    $reg_nummer_mh = ($reg_nummer -split "//")[0] -replace "RVG\s+"
    $registratie_url_mh = $registratie["REG_URL_HANDELSVERGUNNINGHOUDER"] -replace $reg_nummer_MH, $registratie["REG_NUMMER_MAH"]
    $registratie["REG_URL_MAH"] = $registratie_url_mh
    # get content from merkhouder url
    $product_content_mh = curl --silent $registratie["REG_URL_MAH"]
    # Save content to a file for debug purposes
    # $product_content_mh | Out-File -FilePath "$($reg_id)_mh_content.txt"
    # Get <!-- downloads --> content
    $regex_downloads = '(?s)<!-- downloads -->(.*?)<!-- end downloads -->'
    $download_content_mh = [regex]::Match($product_content_mh, $regex_downloads)
    # Get urls from download_content
    $regex_urls = '<a\s+href="([^"]+)"'
    $urls_matches_mh = [regex]::Matches($download_content_mh, $regex_urls)
    # Save urls to a file for debug purposes
    # $urls_matches_mh | Out-File -FilePath "url_mh_$($reg_id)_urls.txt"

    ############################################################################################################
    # Zoek voor "Niet geregistreed geneesmiddel"
    ############################################################################################################
    
    $geen_rvg = $product_content_mh -match "Niet geregistreerd geneesmiddel"
    If ($geen_rvg) {
        $regex_comment = "(?<=</h2>)(.*?)(?=<br>)"
        $rvg_comment = [regex]::Match($product_content_mh, $regex_comment).Value.Trim()
        If ($rvg_comment) {
            $registratie["RVG_COMMENT"] = $rvg_comment
        } else { $registratie["RVG_COMMENT"] = "" }
        Write-Output "Niet geregistreerd geneesmiddel. $rvg_comment..."
        # Save content to a file for debug purposes
        # $product_content_mh | Out-File -FilePath "$($reg_id)_mh_content.txt"    
    } else { $registratie["RVG_COMMENT"] = ""}
    $registratie["RVG_WIJZIGING"] = ""

    ############################################################################################################
    # Zoek voor "Additioneel risicominimalisatie materiaal"
    ############################################################################################################
 
    $registratie["ARMM_URL_MAH"] = ""
    $registratie["ARMM_FILE_MAH"] = ""
    $registratie["ARMM_CHECKSUM_MAH"] = ""
    $registratie["ARMM_CHECKSUM_MAH_WIJZIGING"] = ""
       
    # ARMM MERKHOUDER  
    $armm = $product_content_mh -match "Additioneel risicominimalisatie materiaal"
    If ($armm) {
       # Get ARMM download
       $regex_armm_url = "https://www.geneesmiddeleninformatiebank.nl/arms/h[0-9]+_armm.pdf"
       $armm_url = ($urls_matches_mh | Select-String -Pattern $regex_armm_url).Matches.Value
       $registratie["ARMM_URL_MAH"] = $armm_url
       # Download file from armm_url
       Write-Output "Download ARMM merkhouder document voor registratie $($registratie["PRODUCTNAAM"])..."
       # extract file name from url
       $armm_file = "ARMM_MAH_" + "$($date)_" + [System.IO.Path]::GetFileName($armm_url)         
       $registratie["ARMM_FILE_MAH"] = $armm_file
       curl --silent $armm_url -o $armm_file 
       # Get checksum of the file
       $checksum = Get-FileHash -Path $armm_file -Algorithm MD5
       $registratie["ARMM_CHECKSUM_MAH"] = $checksum.Hash
       $registratie["ARMM_CHECKSUM_MAH_WIJZIGING"] = ""
       # Save file to download folder
       If ($download) {Move-Item $armm_file -Destination $download_folder -ErrorAction SilentlyContinue}
       # Delete downloaded file
       Remove-Item $armm_file -ErrorAction SilentlyContinue     
    }
   
    ############################################################################################################
    # Zoek voor bijsluiter
    ############################################################################################################
   
    $registratie["BIJSLUITER_URL_MAH"] = ""
    $registratie["BIJSLUITER_FILE_MAH"] = ""
    $registratie["BIJSLUITER_CHECKSUM_MAH"] = ""
    $registratie["BIJSLUITER_CHECKSUM_MAH_WIJZIGING"] = ""
    
    # BIJSLUITER MERKHOUDER  
    If (-not $geen_rvg) {
        $regex_bijsluiter_url = "https://www.geneesmiddeleninformatiebank.nl/bijsluiters/h[0-9]+.pdf"
        # Find bijsluiter url in url_matches
        $bijsluiter_url = ($urls_matches_mh | Select-String -Pattern $regex_bijsluiter_url).Matches.Value
        $registratie["BIJSLUITER_URL_MAH"] = $bijsluiter_url
        If ($bijsluiter_url) {
            # Download file from bijsluiter_url
            Write-Output "Download bijsluiter merkhouder document voor registratie $($registratie["PRODUCTNAAM"])..."
            $bijsluiter_file = "BIJSLUITER_MAH_" + "$($date)_" + [System.IO.Path]::GetFileName($bijsluiter_url) 
            $registratie["BIJSLUITER_FILE_MAH"] = $bijsluiter_file
            curl --silent $bijsluiter_url -o $bijsluiter_file 
            # Get checksum of the file
            $checksum = Get-FileHash -Path $bijsluiter_file -Algorithm MD5
            $registratie["BIJSLUITER_CHECKSUM_MAH"] = $checksum.Hash
            $registratie["BIJSLUITER_CHECKSUM_MAH_WIJZIGING"] = ""
            # Save file to download folder
            If ($download) {Move-Item $bijsluiter_file -Destination $download_folder -ErrorAction SilentlyContinue}
            # Delete downloaded file
            Remove-Item $bijsluiter_file -ErrorAction SilentlyContinue
        }
    }

    ############################################################################################################
    # Zoeken naar SMPC
    ############################################################################################################
    
    $registratie["SMPC_URL_MAH"] = ""
    $registratie["SMPC_FILE_MAH"] = ""
    $registratie["SMPC_CHECKSUM_MAH"] = ""
    $registratie["SMPC_CHECKSUM_MAH_WIJZIGING"] = ""
    $regex_smpc_url = "https://www.geneesmiddeleninformatiebank.nl/smpc/h[0-9]+_smpc.pdf"

    # SMPC MERKHOUDER
    # Find smpc url in url_matches
    $smpc_url = ($urls_matches_mh | Select-String -Pattern $regex_smpc_url).Matches.Value
    $registratie["SMPC_URL_MAH"] = $smpc_url
    If ($smpc_url) {
        # Download file from smpc_url
        Write-Output "Download SMPC merkhouder document voor registratie $($registratie["PRODUCTNAAM"])..."
        $smpc_file = "SMPC_MAH_" + "$($date)_" + [System.IO.Path]::GetFileName($smpc_url)
        $registratie["SMPC_FILE_MAH"] = $smpc_file
        curl --silent $smpc_url -o $smpc_file
        # Get checksum of the file
        $checksum = Get-FileHash -Path $smpc_file -Algorithm MD5
        $registratie["SMPC_CHECKSUM_MAH"] = $checksum.Hash
        $registratie["SMPC_CHECKSUM_MAH_WIJZIGING"] = ""
        # Save file to download folder
        If ($download) {Move-Item $smpc_file -Destination $download_folder -ErrorAction SilentlyContinue}
        # Delete downloaded file
        Remove-Item $smpc_file -ErrorAction SilentlyContinue
    }

    # DATUM DETECTIE
    $registratie["DATUM_DETECTIE"] = $datum

    # Add $registratie to $registraties
    Write-Output "--------------------------------------------------------------------------------"
    Write-Output "Registratie $($i) van $($rows.Count) wordt verwerkt..."
    $registraties += $registratie
}

# Save $registraties to a file
$registraties | ConvertTo-Json | Out-File -FilePath "CBG_LIJST_$date.json"
Write-Output "Registraties opgeslagen in CBG_LIJST_$date.json"
