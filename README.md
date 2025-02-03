# armm-check
Powershell scripts voor controle of ARMM documenten zijn gewijzigd op CBG website.

## Download bestanden
Klik op een scriptbestand (bijvoorbeeld `download_cbg_lijst.ps1`) dat je wilt downloaden. Klik op de "Raw" knop rechts bovenaan de codeweergave.
Druk op Ctrl + S (of gebruik Opslaan als in het browsermenu) om het bestand op te slaan. Kies een locatie naar keuze.
Kies een locatie, zoals je Bureaublad of een map naar keuze.
Zorg ervoor dat de bestandsnaam eindigt op .ps1.
1. Download Powershell scripts van deze website. Selecteer bestand en klik op download knop rechtsboven (download raw file).
2. Accepteer waarschuwing en download file.
3. Unblock scripts via rechtermuisknop - properties - Unblock.


## Powershell permissies
Wanneer het starten van Powershell scripts wordt geblokkeerd, dienen eventueel permissies te worden aangepast. Dit is eenmalig per gebruiker.
Dit kan via `Set-Executionpoliy -Executionpoliy Unrestricted -Scope CurrentUser`


## Download CBG lijst
Selecteer via verkenner de juiste folder waar het script is opgeslagen.
Open Windows Verkenner en navigeer naar de map waar het script is opgeslagen.
Houd Shift ingedrukt en klik met de rechtermuisknop op een lege ruimte in de map.
Klik op "PowerShell-venster hier openen".
Type `download` en vul aan met de <TAB> toets. `./download_cbg_lijst.ps1` verschijnt. Druk op <ENTER> om het script te starten.

## Check CBG

