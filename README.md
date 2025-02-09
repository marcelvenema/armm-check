# armm-check

PowerShell scripts voor controle of ARMM documenten zijn gewijzigd op de CBG website.

## Download script bestanden vanaf GitHub

1. Klik op een scriptbestand (bijvoorbeeld `download_cbg_lijst.ps1`) dat je wilt downloaden.
2. Klik op de "Raw" knop rechts bovenaan de codeweergave.
3. Druk op Ctrl + S (of gebruik "Opslaan als" in het browsermenu) om het bestand op te slaan. Kies een locatie naar keuze.
4. Zorg ervoor dat de bestandsnaam eindigt op .ps1.
5. Download PowerShell scripts van deze website. Selecteer het bestand en klik op de download knop rechtsboven (download raw file).
6. Accepteer de waarschuwing en download het bestand.
7. Deblokkeer scripts via rechtermuisknop - Eigenschappen - Deblokkeren.

## PowerShell permissies om een script te starten

Wanneer het starten van PowerShell scripts wordt geblokkeerd, dienen eventueel permissies te worden aangepast. Dit is eenmalig per gebruiker. Dit kan via:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
```

# ARMM-check

Het maken van de ARMM check bestaat uit twee stappen: het `download_CBG_lijst.ps1` script om bijsluiters-, smpc- en armm-bestanden vanaf de CBG website te controleren en het `CHECK_CBG_lijst.ps1` script om deze te vergelijken met de voorgaande output en deze te plaatsen in een Excel bestand.

## Download CBG lijst

In deze stap wordt op de CBG website alle benodigde informatie binnengehaald waarmee de controle en vergelijking gemaakt moet worden. De uitkomst van deze stap is een bestand `CBG_LIJST_<datum>.json`.

1. Open Windows Verkenner en navigeer naar de map waar het script is opgeslagen.
2. Klik in de verkenner balk op de folder, druk op ALT+F en klik op "PowerShell-venster hier openen".

Een PowerShell venster wordt geopend in de folder waar ook het script is geplaatst. Type `download` en vul aan met de <TAB> toets. `./download_cbg_lijst.ps1` verschijnt. Druk op <ENTER> om het script te starten.

Er volgt mogelijk een beveiligingswaarschuwing; toets 'r' (van run once) om het script te starten en druk op <ENTER>. Het script start en documenten worden vanaf de CBG website gedownload en verwerkt. Dit kan enige tijd duren. Op het scherm is zichtbaar welk product wordt verwerkt.

Het resultaat van het script is een bestand genaamd `CBG_LIJST_<huidige datum>.json`.

## Check CBG

Dit is de tweede stap. Het script vergelijkt het `CBG_LIJST_<datum>.json` bestand met het vorige `CBG_CHECK_<datum>.json` bestand en maakt twee nieuwe bestanden aan: `CBG_CHECK_<huidige datum>.json` als JSON bestand en `CBG_CHECK_<huidige datum>.xlsx` als Excel bestand.

De volledige syntax is:

```powershell
./check_cbg_lijst.ps1 -new ./CBG_LIJST_<datum>.json -old ./CBG_CHECK_<datum>.json
```

1. Type `check` en vul aan met de <TAB> toets. `./check_cbg_lijst.ps1` verschijnt.
2. Vul het aan met `-new cbg_lijst`. Druk op <Tab> om het bestand aan te vullen. Met <TAB> wordt een volgend bestand gekozen.
3. Vul het aan met `-old cbg_check`. Druk op <Tab> om het bestand aan te vullen. Let erop dat dit een JSON bestand is. Met <TAB> wordt een volgend bestand gekozen.
4. Controleer de volledige syntax `./check_cbg_lijst.ps1 -new ./CBG_LIJST_<datum>.json -old ./CBG_CHECK_<datum>.json` op juistheid en klik op <Enter> om het script te starten.

Er volgt mogelijk een beveiligingswaarschuwing; toets 'r' (van run once) om het script te starten en druk op <ENTER>. Het script start nu met het vergelijken van beide lijsten, geeft de verschillen aan en maakt twee bestanden aan: een `CBG_CHECK_<huidige datum>.json` bestand en een `CBG_CHECK_<huidige datum>.xlsx` Excel bestand.

Het `CBG_CHECK_<huidige datum>.json` bestand kan gebruikt worden als vergelijkingsbestand voor de komende check.
