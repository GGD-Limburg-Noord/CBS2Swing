
### Help file bij CBS-tabellen naar Swing.R

library(cbsodataR)
library(openxlsx)


## Downloaden van een overzicht van alle beschikbare tabellen

CBS_catalogus <- cbs_get_datasets(catalog = "CBS")
write.xlsx(CBS_catalogus, paste0("output/tabellijst CBS ", Sys.Date(),".xlsx"), overwrite = T)

totaal_catalogus <- cbs_get_datasets(NULL)
write.xlsx(totaal_catalogus, paste0("output/tabellijst totaal ", Sys.Date(),".xlsx"), overwrite = T)

# Om tabellen te zoeken die voor de GGD van toepassing zijn, kan bijvoorbeeld gezocht worden op RegioS en WijkenenBuurten:
# Open het Excel bestand
# Breng een filter aan op de 1e rij (kolomtitels)
# Filter kolom 'DefaultSelection' op vrije tekst 'RegioS' (voor tabellen met gemeentecijfers) of 'WijkenEnBuurten' (voor tabellen met wijkcijfers).


## Metadata downloaden en bekijken

# Om je input file te vullen is het handig om gebruik te maken van de CBS metadata. Hierin staat bijvoorbeeld welke kruisingen er mogelijk zijn.

metadata_86131NED <- cbs_get_meta("86131NED")

# De metadata bestaat uit een list van dataframes
View(metadata_86131NED)

# Bijna alle metadata-lists bevatten in ieder geval TableInfos, DataProperties, CategoryGroups.
# Veel metadata-lists bevatten ook Perioden en een gebiedsindeling zoals RegioS of WijkenEnBuurten.
# De kruisingsvariabele heet bijvoorbeeld Persoonskenmerken of PersoonsEnHuishoudenskenmerken.

dimcat_kolomnaam <- unique(metadata_86131NED$CategoryGroups$DimensionKey)
dimcat_kolomnaam

# In deze tabel kun je zien welke dimensie-categorieen beschikbaar zijn per kruisingsvariabele (voor dimcat_bewaren)
View(metadata_86131NED$PersoonsEnHuishoudenskenmerken)
