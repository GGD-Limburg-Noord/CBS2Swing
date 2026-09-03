#  CBS2Swing: CBS-data naar Swing

- **Doel:** Gegevens ophalen van CBS Statline en klaarzetten voor Swing
- **Auteur/Contactpersoon:** Kris de Klerk, GGD Limburg-Noord
- **Gebruik:** Zie handleiding hieronder
- **Databronnen:** CBS Statline
- **Output:** Een data en metadata CSV-bestand, klaar om in te laden in Swing Studio

Met dit script haal je data op via de CBS Statline API en zet je die om naar een bestand dat je direct kunt importeren in Swing Studio. Het script is zo gemaakt dat het voor elke GGD bruikbaar is, zonder dat je de code zelf hoeft aan te passen. Je stelt alles in via een los Excel-bestand.

Bij de ontwikkeling van dit script is gebruikgemaakt van AI-ondersteuning (Claude, Anthropic).

___

## Handleiding

### Stap 1: Code downloaden en R script openen
1. Klik rechtsboven op deze Github pagina op de groene knop `<> Code`.
2. Kies onderaan het dropdown-menu `Download ZIP`.
3. Pak het ZIP-bestand uit op de plek waar je wilt werken.
4. Open in de uitgepakte map het bestand `CBS2Swing.Rproj`. Dit opent het project in RStudio.
5. Open vanuit RStudio (of vanuit de map zelf) het hoofdscript: `main.R`.

### Stap 2: Configuratie invullen in input.xlsx

Het script haalt zijn instellingen uit `data/openbaar/input.xlsx`. Verander de bestandsnaam niet, of als je dat toch graag wilt, verander dan ook de bestandsnaam in het script.

`input.xlsx` heeft drie tabbladen: **GGD**, **Downloaden** en **Indicatoren**.

<br>

#### Tabblad "GGD"

Selecteer je GGD uit de dropdown opties.


#### Tabblad "Downloaden"

Hier geef je aan welke CBS-tabellen je wilt ophalen.

| Kolom | Uitleg |
| :--- | :--- |
| **Onderwerp** | Een zelfgekozen naam voor het onderwerp van de CBS-tabel. Deze naam is de koppelsleutel tussen dit tabblad en het tabblad Indicatoren. Zorg ervoor dat de naam op beide plekken exact hetzelfde is, en dat er geen spaties in de cel staan. |
| **Tabelcode** | De CBS-tabelcode, bijvoorbeeld `86131NED`. |
| **Periode** | De periode(s) waarvan je data wilt, in hetzelfde format als CBS gebruikt (bijv. `2025JJ00`). Wil je meerdere losse periodes? Scheid ze met een komma. Wil je een reeks van begin- tot eindperiode? Gebruik een streepje, bijv. `2018JJ00 - 2024JJ00`. Spaties voor en na een periode worden genegeerd. |
| **Lege_cellen** | De waarde die het script in de output zet als er voor een gekozen combinatie geen data beschikbaar is. Zie de tabel hieronder voor de standaardwaarden in Swing. |

Standaardwaarden voor lege cellen in Swing:

| Waarde | Betekenis |
| :--- | :--- |
| -99996 | Empty |
| -99997 | Hidden |
| -99998 | Not applicable |
| -99999 | Missing |

**Waar vind je de beschikbare periodes en tabelcodes?** Kijk op StatLine bij de tabelinformatie. Voor tabel 86131NED ga je bijvoorbeeld naar `https://opendata.cbs.nl/#/CBS/nl/dataset/86131NED/table` en klik je rechtsboven op het i-tje. Je kunt dit ook opzoeken via de metadata (zie `help.R`).

<br>

#### Tabblad "Indicatoren"

Hier geef je aan welke indicatoren (variabelen) je uit elke CBS-tabel wilt, en hoe ze in Swing moeten heten.

| Kolom | Uitleg |
| :--- | :--- |
| **Onderwerp** | Moet exact overeenkomen met een Onderwerp uit het tabblad Downloaden. |
| **CBS_indicatornaam** | De naam van de indicator zoals CBS die gebruikt. |
| **CBS_omschrijving_kort** | Optioneel: eigen (korte) aantekening over wat de indicator inhoudt. Wordt niet gebruikt door het script. |
| **CBS_omschrijving_lang** | Optioneel: eigen (lange) aantekening over wat de indicator inhoudt. Wordt niet gebruikt door het script. |
| **override_provisional_period** | Optioneel. Normaal bepaalt het script zelf vanaf welke periode de cijfers "voorlopig" zijn (op basis van de CBS-metadata). Wijkt dit voor één indicator af (te vinden in de tabelinformatie op StatLine)? Vul dan hier de juiste periode in, in hetzelfde format als CBS gebruikt (bijv. `2025JJ00`). Vanaf deze periode wordt alles automatisch als voorlopig gemarkeerd in Swing. Spaties voor en na de periode worden genegeerd. |
| **indicator_includeren** | Kies uit de dropdown opties 'ja' of 'nee' om de indicator mee te nemen in de analyse. |
| **Swing_indicator_code** | Een unieke, zelfgekozen naam voor deze indicator in Swing (het veld "indicator code" in Swing Studio). Koppelsleutel -> Moet exact overeenkomen met Swing_indicator_code in tabblad Kruisingen, als deze indicator daar voorkomt. |
| **Swing_name** | Een zelfgekozen korte beschrijving voor Swing (het veld "name" in Swing Studio). Wordt onder andere gebruikt voor titels van grafieken. Maximaal 100 tekens, inclusief spaties. |
| **Swing_description** | Een zelfgekozen beschrijving voor Swing (het veld "description" in Swing Studio). Wordt onder andere gebruikt voor hover-over labels. Geen tekenlimiet. |


#### Tabblad "Kruisingen"

Hier geef je aan welke dimensies (kruisingen) van de variabelen je wilt selecteren, per indicator.

Heeft de CBS tabel geen dimensies? Dan hoeft deze indicator helemaal niet in dit tabblad te staan. 

Heeft de CBS tabel meerdere dimensies (bijvoorbeeld zowel Geslacht als Leeftijd)? Voeg dan je selectie van alle dimensies voor deze indicator in meerdere rijen toe: één rij per dimensie, allemaal met dezelfde Swing_indicator_code maar een eigen dimcat_kolomnaam en dimcat_bewaren.

**Let op**: Als je niet voor alle dimensies van de tabel een selectie maakt, worden van de overgebleven dimensies alle categorieën meegenomen in de analyse, zonder dat ze herkenbaar zijn aan een unieke naam of beschrijving. Zorg er dus voor dat je voor elke dimensie die de CBS-tabel heeft, ook een rij toevoegt.

| Kolom | Uitleg |
| :--- | :--- |
| **Swing_indicator_code** | Moet exact overeenkomen met de Swing_indicator_code van de bijbehorende indicator in tabblad Indicatoren. Koppelsleutel. |
| **dimcat_kolomnaam** | De CBS-naam van de dimensie waarop je wilt filteren, bijvoorbeeld `Geslacht`. |
| **dimcat_bewaren** | De CBS-code van de dimensie categorie die je uit de dimensie wilt behouden, bijvoorbeeld `T001038` voor zowel mannen als vrouwen. |
| **dimcat_omschrijving** | Optioneel: eigen aantekening over wat de gekozen dimensie categorie inhoudt. |


### Stap 3: Script draaien

1. Sluit `input.xlsx` voordat je het script draait.
2. Run het script.
3. De resultaten worden opgeslagen in de map `output`, als `data.csv` en `metadata.csv`. Uiteraard kun je de bestandsnaam en locatie aanpassen.
4. Controleer je output altijd, bijvoorbeeld steekproefsgewijs.

---

## Help

Bekijk `help.R` voor uitleg over:

- Hoe je relevante CBS-tabellen vindt
- Hoe je opzoekt welke dimensie-combinaties beschikbaar zijn binnen een tabel

## Problemen of vragen?

Loop je tegen een probleem aan? Maak dan een Issue aan op deze GitHub-pagina (hiervoor heb je een GitHub-account nodig).

### Disclaimer

De maker van dit script is niet verantwoordelijk voor eventuele fouten in de data door gebruik van dit script.
"CBS2Swing" 
