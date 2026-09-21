#
# Dit script vormt CBS-data om naar een bestand wat herkend wordt bij import in Swing. 
# De configuratie gaat middels een Excelsheet, zie bijgevoegde documentatie.
# Problemen en verzoeken kunnen worden ingediend op
# https://github.com/GGD-Limburg-Noord/CBS2Swing
#
# Versie: 21 september 2026
#



# Environment leegmaken
rm(list = ls())

# Laden van benodigde packages
benodigde_packages <- c("cbsodataR", "data.table", "tidyverse", "readxl", "openxlsx", "purrr", "ISOweek")

for (package in benodigde_packages) {
  if (!requireNamespace(package, quietly = TRUE)) {
    install.packages(package)
  }
}

lapply(benodigde_packages, library, character.only = TRUE)


# Aanmaken lijsten voor gebruik in de loop
metadata_list = list()
data_list = list()

# Inladen configuratie
CBS_tabellen <- read.xlsx('data/openbaar/input.xlsx', sheet = 'Downloaden')
CBS_tabellen_alle <- CBS_tabellen                    # ongefilterde versie, nodig om ook de bron te bepalen van onderwerpen die volledig uit indicator_includeren = "nee" bestaan
indicatoren <- readxl::read_excel('data/openbaar/input.xlsx', sheet = 'Indicatoren', col_types = "text") %>% 
  filter(!is.na(CBS_indicatornaam))
indicatoren_alle <- indicatoren                      # ongefilterde versie (incl. indicator_includeren = "nee"), nodig voor automatische 'bron'-bepaling van formule-indicatoren
kruisingen <- readxl::read_excel('data/openbaar/input.xlsx', sheet = 'Kruisingen', col_types = "text") %>%
  mutate(across(everything(), str_trim))

check_onderwerp_leeg <- function(reden) {
  ontbrekend <- CBS_tabellen$Onderwerp[!(CBS_tabellen$Onderwerp %in% unique(indicatoren$Onderwerp))]
  if (length(ontbrekend) > 0) {
    message("Voor de volgende onderwerp(en) zijn geen indicatoren meer over (reden: ", reden, "): ",
            paste(ontbrekend, collapse = ", "), ".")
  }
}

indicatoren <- indicatoren %>%
  mutate(indicator_includeren = str_trim(indicator_includeren)) %>%
  filter(indicator_includeren == "ja")

check_onderwerp_leeg("indicator_includeren = 'nee' voor alle indicatoren van dit onderwerp")

if(any(is.na(indicatoren$Swing_indicator_code))) {
  message("Let op: Er zijn enkele indicatoren zonder Swing_indicator_code. Deze worden niet meegenomen in de analyse.")
}

indicatoren <- indicatoren %>%
  filter(!is.na(Swing_indicator_code))

check_onderwerp_leeg("Swing_indicator_code ontbreekt voor alle indicatoren van dit onderwerp")

CBS_tabellen <- CBS_tabellen %>%
  filter(Onderwerp %in% unique(indicatoren$Onderwerp))

kruisingen_check <- kruisingen %>%
  mutate(aantal_ingevuld = (!is.na(dimcat_bewaren)) + (!is.na(dimcat_CategoryGroupID_bewaren)))

if (any(kruisingen_check$aantal_ingevuld != 1)) {
  fout_rijen <- kruisingen_check %>% filter(aantal_ingevuld != 1)
  stop("Vul voor elke rij in Kruisingen precies één van dimcat_bewaren of dimcat_CategoryGroupID_bewaren in. Foutieve rijen: ",
       paste(fout_rijen$Swing_indicator_code, fout_rijen$dimcat_kolomnaam, sep = "/", collapse = ", "))
}


# Inladen bestand gemeentecodes en niveau voor Swing

GGD <- read.xlsx('data/openbaar/input.xlsx', sheet = 'GGD')

GGD <- str_remove(GGD[1,1], "^\\d{1,2}\\.\\s")

CBS_toc <- cbs_get_datasets(catalog = "CBS")   # eenmalig ophalen

get_gebieden_tabelcode <- function(jaar) {
  match <- CBS_toc %>% filter(Title == paste("Gebieden in Nederland", jaar))
  if (nrow(match) != 1) stop("Geen unieke 'Gebieden in Nederland'-tabel gevonden voor ", jaar)
  match$Identifier
}


get_topic_key <- function(metadata_gebieden_CBS, groep_titel, veld_titel) {
  groep_id <- metadata_gebieden_CBS$DataProperties$ID[
    metadata_gebieden_CBS$DataProperties$Type == "TopicGroup" &
      metadata_gebieden_CBS$DataProperties$Title == groep_titel]          # groep_titel is bijvoorbeeld "Codes en namen van gemeenten" of "GGD-regio's"
  metadata_gebieden_CBS$DataProperties$Key[
    metadata_gebieden_CBS$DataProperties$ParentID %in% groep_id &
      metadata_gebieden_CBS$DataProperties$Title == veld_titel]
}


gemeentes_GGD_cache <- list()

get_gemeentes_GGD <- function(
    GGD_naam, 
    jaar,
    geolevelcode_NL = "nederland",
    geoitemcode_NL  = "1",
    geolevelcode_PV = "provincie",
    geoitemcode_PV  = NULL,
    #geoitemcode_PV  = 12,      # Voorbeeld voor Limburg bij GGD Limburg-Noord
    geolevelcode_GM = "gemeente") {
  
  key <- paste(GGD_naam, jaar, sep = "_")
  
  if (key %in% names(gemeentes_GGD_cache)) {
    return(gemeentes_GGD_cache[[key]])
  }
  
  tabelcode <- get_gebieden_tabelcode(jaar)
  metadata_gebieden_CBS <- cbs_get_meta(catalog = "CBS", id = tabelcode)
  
  gemeentenaam_key    <- get_topic_key(metadata_gebieden_CBS, "Codes en namen van gemeenten", "Naam")
  gemeentecode_key    <- get_topic_key(metadata_gebieden_CBS, "Codes en namen van gemeenten", "Code")
  ggd_key             <- get_topic_key(metadata_gebieden_CBS, "GGD-regio's", "Naam")
  provincienaam_key   <- get_topic_key(metadata_gebieden_CBS, "Provincies", "Naam")
  provinciecode_key   <- get_topic_key(metadata_gebieden_CBS, "Provincies", "Code")
  
  # cbsodataR heeft zelf de GeoDimension-kolom nodig om zijn interne post-processing
  # uit te kunnen voeren, ook al gebruiken wij die kolom hier verder niet inhoudelijk
  geodimension_key <- metadata_gebieden_CBS$DataProperties$Key[metadata_gebieden_CBS$DataProperties$Type == "GeoDimension"]
  
  data <- cbs_get_data(
    catalog = "CBS",
    id = tabelcode,
    select = unique(c(gemeentenaam_key, gemeentecode_key, ggd_key, provincienaam_key, provinciecode_key, geodimension_key))
  ) %>%
    mutate(across(where(is.character), str_trim)) %>%
    filter(.data[[ggd_key]] == GGD_naam)
  
  gemeentes <- data %>%
    mutate(
      jaar = jaar,
      geoitemcode = as.character(as.integer(str_remove(.data[[gemeentecode_key]], "^GM"))),
      geoitemcode_long = .data[[gemeentecode_key]],
      geolevelcode = geolevelcode_GM
    ) %>%
    rename(BESCHRIJVING = all_of(gemeentenaam_key)) %>%
    select(BESCHRIJVING, jaar, geoitemcode, geoitemcode_long, geolevelcode)
  
  provincies <- data %>%
    filter(
      !is.na(.data[[provincienaam_key]]),
      !is.na(.data[[provinciecode_key]])
    ) %>%
    distinct(
      BESCHRIJVING = paste0(.data[[provincienaam_key]], " (PV)"),
      geoitemcode = as.character(as.integer(str_remove(.data[[provinciecode_key]], "^PV"))),
      geoitemcode_long = .data[[provinciecode_key]]
    ) 
  
  # Provincie geoitemcode kan worden overschreven als ie niet nul is (bijv. geoitemcode_PV = "12")
  if (!is.null(geoitemcode_PV)) {
    provincies$geoitemcode <- as.character(geoitemcode_PV)
  }
  
  provincies <- provincies %>%
    mutate(
      jaar = jaar,
      geolevelcode = geolevelcode_PV
    )
  
  result <- bind_rows(
    gemeentes,
    tibble(
      BESCHRIJVING = "Nederland",
      jaar = jaar,
      geoitemcode = geoitemcode_NL,
      geoitemcode_long = "NL01",
      geolevelcode = geolevelcode_NL
    ),
    provincies
  )
  
  gemeentes_GGD_cache[[key]] <<- result
  result
}

# Inladen omzettabel die CBS-units omzet naar Swing-units
omzettabel_unit <- read.csv2('data/openbaar/omzettabel_unit.csv')

# Functie om de uitgeschreven periodes uit metadata om te zetten naar periodes CBS-format (nodig bij datatabellen zonder perioden kolom)
extract_snapshot_periodcode <- function(period_string) {
  maanden <- c(
    "januari" = "01", "februari" = "02", "maart" = "03", "april" = "04",
    "mei" = "05", "juni" = "06", "juli" = "07", "augustus" = "08",
    "september" = "09", "oktober" = "10", "november" = "11", "december" = "12"
  )
  
  jaar <- str_extract(period_string, "\\d{4}")
  maand_naam <- str_extract(tolower(period_string), paste(names(maanden), collapse = "|"))
  
  if (!is.na(maand_naam)) {
    paste0(jaar, "MM", maanden[[maand_naam]])
  } else {
    paste0(jaar, "JJ00")
  }
}

# Functie om periodes in CBS-format om te zetten naar Swing-format
CBS_to_Swing_periodcodes <- function(CBS_periode) {
  case_when(
    str_detect(CBS_periode, "JJ") ~ str_extract(CBS_periode, "\\d{4}"),
    str_detect(CBS_periode, "KW") ~ paste0("q", str_remove(str_extract(CBS_periode, "KW\\d{2}"), "KW"), "y", str_extract(CBS_periode, "\\d{4}")),
    str_detect(CBS_periode, "MM") ~ paste0("m", as.numeric(str_remove(str_extract(CBS_periode, "MM\\d{2}"), "MM")), "y", str_extract(CBS_periode, "\\d{4}")),
    str_detect(CBS_periode, "W1") ~ paste0("w", str_remove(str_extract(CBS_periode, "W1\\d{2}"), "W1"), "y", str_extract(CBS_periode, "\\d{4}"))
  )
}

# Functie om een CategoryGroupID te vertalen naar de losse dimensie-codes eronder
resolve_dimcat_codes <- function(kolomnaam, dimcode_bewaren, dimgroep_bewaren) {
  if (!is.na(dimcode_bewaren)) {
    return(str_trim(str_split_1(dimcode_bewaren, ",")))
  }
  
  metadata_onderwerp_CBS[[kolomnaam]] %>%
    filter(as.character(CategoryGroupID) == dimgroep_bewaren) %>%
    pull(Key)
}

# Functie om ongevraagde dimensie-categorieen uit het databestand te filteren
build_var_table <- function(i) {
  code <- indicator_subset$Swing_indicator_code[i]
  filters_i <- kruisingen %>% filter(Swing_indicator_code == code)
  
  enkelvoudig  <- filters_i %>% filter(!is.na(dimcat_bewaren))
  gegroepeerd  <- filters_i %>% filter(!is.na(dimcat_CategoryGroupID_bewaren))
  
  gefilterd <- reduce(
    seq_len(nrow(filters_i)),
    function(data, j) {
      codes <- resolve_dimcat_codes(
        filters_i$dimcat_kolomnaam[j],
        filters_i$dimcat_bewaren[j],
        filters_i$dimcat_CategoryGroupID_bewaren[j]
      )
      filter(data, .data[[filters_i$dimcat_kolomnaam[j]]] %in% codes)
    },
    .init = onderwerp_GGD
  )
  
  plak_of_NA <- function(x) {
    if (length(x) == 0) NA_character_ else paste(x, collapse = "; ")
  }
  
  gefilterd %>%
    group_by(Perioden, BESCHRIJVING, geoitemcode, geolevelcode) %>%
    summarize(values = sum(.data[[indicator_subset$CBS_indicatornaam[i]]]), .groups = "drop") %>%
    mutate(
      variablecode                   = code,
      dimcat_kolomnaam                = plak_of_NA(enkelvoudig$dimcat_kolomnaam),
      dimcat_bewaren                  = plak_of_NA(enkelvoudig$dimcat_bewaren),
      dimcat_kolomnaam_CategoryGroup  = plak_of_NA(gegroepeerd$dimcat_kolomnaam),
      CategoryGroupID_bewaren         = plak_of_NA(gegroepeerd$dimcat_CategoryGroupID_bewaren)
    )
}


# Bij het trouble-shooten kun je het script hiermee testen voor één tabel
#CBS_tabel <- length(1)

for (CBS_tabel in 1:nrow(CBS_tabellen)){
  
  # Selecteren indicatoren per CBS tabel
  indicator_subset <- indicatoren %>%
    filter(Onderwerp %in% CBS_tabellen$Onderwerp[CBS_tabel])
  
  # Metadata downloaden vanaf CBS
  metadata_onderwerp_CBS <- cbs_get_meta(catalog = "CBS", 
                                         id = CBS_tabellen$Tabelcode[CBS_tabel])
  
  # Check of er van alle beschikbare dimensies een categorie is geselecteerd
  alle_dimensies <- metadata_onderwerp_CBS$DataProperties %>%
    filter(Type == "Dimension") %>%
    pull(Key)
  
  ontbrekende_dims <- indicator_subset %>%
    distinct(Swing_indicator_code) %>%
    rowwise() %>%
    mutate(dims_ontbreken = list(setdiff(
      alle_dimensies,
      kruisingen$dimcat_kolomnaam[kruisingen$Swing_indicator_code == Swing_indicator_code]
    ))) %>%
    ungroup() %>%
    filter(lengths(dims_ontbreken) > 0)
  
  if (nrow(ontbrekende_dims) > 0) {
    stop("Voor de volgende indicator(en) ontbreekt in het tabblad Kruisingen van het input bestand minstens één dimcat_kolomnaam: ",
         paste0(ontbrekende_dims$Swing_indicator_code, " (mist: ",
                map_chr(ontbrekende_dims$dims_ontbreken, paste, collapse = ", "), ")",
                collapse = "; "))
  }
  # De gewenste periodes definieren
  heeft_perioden <- !is.null(metadata_onderwerp_CBS$Perioden)
  
  if (heeft_perioden) {
    perioden <- str_split(CBS_tabellen$Periode[CBS_tabel], ",")[[1]] %>%
      str_trim()
    
    perioden <- map(perioden, function(chunk) {
      if (str_detect(chunk, "-")) {
        grenzen <- str_split(chunk, "-")[[1]] %>% str_trim()
        pos <- which(metadata_onderwerp_CBS$Perioden$Key %in% grenzen)
        metadata_onderwerp_CBS$Perioden$Key[min(pos):max(pos)]
      } else {
        chunk
      }
    }) %>% unlist()
    
  } else {
    perioden <- extract_snapshot_periodcode(metadata_onderwerp_CBS$TableInfos$Period)
  }
  
  jaren_nodig <- unique(as.numeric(str_extract(perioden, "\\d{4}")))
  
  gemeentes_GGD_jaren <- map(jaren_nodig, get_gemeentes_GGD, GGD_naam = GGD) %>% 
    bind_rows()
  
  # Geografisch niveau definieren
  geo_col <- metadata_onderwerp_CBS$DataProperties %>%
    filter(Type %in% c("GeoDetail", "GeoDimension"), 
           Key %in% c("RegioS", "WijkenEnBuurten")) %>%
    pull(Key)
  
  # Server-side filter opbouwen om alleen de benodigde regiocodes op te kunnen vragen (kan alleen bij RegioS, dit scheelt tijd)
  if (length(geo_col) == 1 && geo_col == "RegioS") {
    regio_codes <- gemeentes_GGD_jaren %>%
      distinct(geoitemcode_long) %>%
      mutate(geoitemcode_long = str_pad(geoitemcode_long, width = 6, side = "right")) %>%
      pull(geoitemcode_long)
    
    filter_arg <- setNames(list(regio_codes), geo_col)
  } else {
    filter_arg <- list()
  }
  
  
  select_geo <- if (length(geo_col) == 1 && geo_col == "WijkenEnBuurten") {
    "WijkenEnBuurten"
  } else if (length(geo_col) == 1) {
    geo_col
  } else {
    NULL
  }
  
  # Data downloaden via CBS' open data API
  onderwerp <- do.call(cbs_get_data, c(
    list(catalog = "CBS",
         id = CBS_tabellen$Tabelcode[CBS_tabel],
         select = c(if (heeft_perioden) 'Perioden',
                    select_geo,
                    unique(indicator_subset$CBS_indicatornaam[!is.na(indicator_subset$CBS_indicatornaam)]),
                    unique(kruisingen$dimcat_kolomnaam[kruisingen$Swing_indicator_code %in% indicator_subset$Swing_indicator_code]))),
    if (heeft_perioden) list(Perioden = perioden),
    filter_arg
  )) %>%
    cbs_add_label_columns() %>%
    mutate(across(where(is.character), str_trim))
  
  
  # Periode toevoegen (uit metadata als er geen periode-kolom in data staat)
  onderwerp <- onderwerp %>%
    { if (!heeft_perioden) mutate(., Perioden = perioden) else . } %>%
    mutate(jaar = as.numeric(str_extract(Perioden, "\\d{4}")))
  
  
  # Toevoegen selectie wijken en geoitemcode/geolevelcode
  if (length(geo_col) == 1 && geo_col == "RegioS") {
    
    onderwerp_GGD <- onderwerp %>%
      mutate(!!geo_col := str_trim(.data[[geo_col]])) %>%
      inner_join(gemeentes_GGD_jaren,
                 by = c("jaar", setNames("geoitemcode_long", geo_col))) %>%
      select(-jaar)
    
  } else if (length(geo_col) == 1 && geo_col == "WijkenEnBuurten") {
    
    gemeentecodes_GGD <- gemeentes_GGD_jaren %>% 
      filter(geolevelcode == "gemeente") %>% 
      pull(geoitemcode)
    
    onderwerp_GGD <- onderwerp %>%
      mutate(WijkenEnBuurten = str_trim(WijkenEnBuurten)) %>%
      mutate(
        soort        = str_sub(WijkenEnBuurten, 1, 2),
        gemeentecode = as.character(as.integer(str_sub(WijkenEnBuurten, 3, 6)))
      ) %>%
      filter(
        WijkenEnBuurten == "NL01" |
          soort %in% c("GM", "WK") & gemeentecode %in% gemeentecodes_GGD
      ) %>%
      mutate(
        BESCHRIJVING = WijkenEnBuurten_label,
        geoitemcode  = case_when(
          soort == "WK" ~ str_pad(str_remove(WijkenEnBuurten, "^[A-Z]+"),
                                  width = 6, side = "left", pad = "0"),
          TRUE          ~ as.character(as.integer(str_remove(WijkenEnBuurten, "^[A-Z]+")))
        ),
        geolevelcode = case_when(
          soort == "WK" ~ paste0("wijk", str_sub(as.character(jaar), -2)),
          soort == "GM" ~ "gemeente",
          TRUE ~ "nederland"
        )
      ) %>%
      select(-jaar, -WijkenEnBuurten, -WijkenEnBuurten_label, -soort, -gemeentecode)
    
  } else {
    onderwerp_GGD <- onderwerp %>%
      mutate(BESCHRIJVING = "Nederland",
             geoitemcode = "1",
             geolevelcode = "nederland") %>%
      select(-jaar)
  }
  
  
  # Swing-gebiedscodes toevoegen en Weggooien de kolommen die we niet hoeven
  swingdata <- map(1:nrow(indicator_subset), build_var_table) %>% 
    bind_rows() %>% 
    mutate(periodcode = CBS_to_Swing_periodcodes(Perioden)) %>%              # Periodcode toevoegen
    select(-Perioden) %>%
    arrange(variablecode, periodcode)
  
  swingdata$values[is.na(swingdata$values)] <- CBS_tabellen$Lege_cellen[CBS_tabel]
  
  data_list[[CBS_tabel]] = swingdata
  
  
  ## Aanmaken metadata
  
  # Afronding opmaken uit de data
  roundoff = swingdata %>%
    group_by(variablecode) %>%
    summarize(roundoff = max(nchar(str_split_fixed(as.character(values), "\\.", 2)[, 2]))) %>%
    mutate(roundoff = 10^-roundoff) 
  
  unit_lookup <- metadata_onderwerp_CBS$DataProperties %>%
    filter(!is.na(Key), Key != "") %>%
    distinct(Key, Unit)
  
  metaonderwerp <- indicator_subset %>%
    filter(!is.na(Swing_indicator_code)) %>%
    left_join(unit_lookup, by = c("CBS_indicatornaam" = "Key")) %>%
    transmute(
      'indicator code' = Swing_indicator_code,
      'name' = Swing_name,
      'description' = Swing_description,
      'unit' = Unit,
      'bron' = str_remove(metadata_onderwerp_CBS$TableInfos$Source[1], "\\.$"),
      'provisional period' = if_else(
        !is.na(override_provisional_period),
        str_trim(override_provisional_period),
        if("Voorlopig" %in% metadata_onderwerp_CBS$Perioden$Status) {
          CBS_to_Swing_periodcodes(min(metadata_onderwerp_CBS$Perioden$Key[metadata_onderwerp_CBS$Perioden$Status == "Voorlopig"]))
        } else {
          NA_character_
        }
      )
    ) %>%
    left_join(omzettabel_unit, by = c("unit" = "unit_CBS")) %>%
    mutate(unit = if_else(!is.na(unit_Swing), unit_Swing, unit)) %>%
    select(-unit_Swing) %>%
    rename('data type' = data_type_Swing) %>%
    left_join(roundoff, by = c("indicator code" = "variablecode"))
  
  
  metadata_list[[CBS_tabel]] = metaonderwerp
  
  
  # Verwijderen tussen-dataframes
  rm(gemeentes_GGD_jaren, indicator_subset, jaren_nodig, metadata_onderwerp_CBS, metaonderwerp, onderwerp,
     onderwerp_GGD, roundoff, swingdata, geo_col, heeft_perioden, perioden, select_geo)
  
}

# Data/metadata lijsten aan elkaar verbinden tot een dataframe
#data_troubleshooting = bind_rows(data_list)
data = bind_rows(data_list) %>%
  select(-dimcat_kolomnaam, -dimcat_bewaren, -dimcat_kolomnaam_CategoryGroup, -CategoryGroupID_bewaren)
metadata = bind_rows(metadata_list) %>%
  mutate(formula = "")

# Toevoegen berekende indicatoren aan metadata
formula <- readxl::read_excel('data/openbaar/input.xlsx', sheet = 'Formula') %>%
  mutate(across(everything(), str_trim))

# Bron van formule-indicatoren automatisch afleiden uit de bron(nen) van de indicatoren waaruit de formule is opgebouwd. 
# De bron van elke basisindicator is de "Source" uit de CBS-tabelmetadata (TableInfos), opgehaald per onderwerp vanuit de ongefilterde Downloaden-tabel.
bron_per_onderwerp <- CBS_tabellen_alle %>%
  rowwise() %>%
  mutate(bron = str_remove(cbs_get_meta(catalog = "CBS", id = Tabelcode)$TableInfos$Source[1], "\\.$")) %>%
  ungroup() %>%
  select(Onderwerp, bron)

bron_lookup_CBS <- indicatoren_alle %>%
  filter(!is.na(Swing_indicator_code)) %>%
  distinct(Swing_indicator_code, Onderwerp) %>%
  left_join(bron_per_onderwerp, by = "Onderwerp") %>%
  {setNames(.$bron, .$Swing_indicator_code)}

formule_lookup  <- setNames(formula$formula, formula$Swing_indicator_code)

bepaal_bron_formula <- function(code, in_behandeling = character()) {
  
  if (code %in% names(bron_lookup_CBS)) {
    return(bron_lookup_CBS[[code]])
  }
  
  if (!code %in% names(formule_lookup)) {
    return(NA_character_)                            # onbekende term (geen indicator, bijv. een functienaam)
  }
  
  if (code %in% in_behandeling) {
    stop("Circulaire verwijzing gevonden in tabblad Formula bij indicator: ", code)
  }
  
  genoemde_indicatoren <- str_extract_all(formule_lookup[[code]], "[A-Za-z_][A-Za-z0-9_]*")[[1]]
  bronnen <- map_chr(genoemde_indicatoren, bepaal_bron_formula, in_behandeling = c(in_behandeling, code))
  
  if (length(bronnen) > 0 && all(!is.na(bronnen)) && length(unique(bronnen)) == 1) {
    unique(bronnen)
  } else {
    NA_character_
  }
}

formula <- formula %>%
  mutate(bron = map_chr(Swing_indicator_code, bepaal_bron_formula))

if (any(is.na(formula$bron))) {
  message("Let op: voor de volgende formule-indicatoren kon de bron niet automatisch (eenduidig) worden bepaald: ",
          paste(formula$Swing_indicator_code[is.na(formula$bron)], collapse = ", "))
}

metadata <-  omzettabel_unit %>% 
  select(-unit_CBS) %>%
  filter(unit_Swing != "") %>%
  distinct() %>%
  right_join(formula, by = join_by("unit_Swing")) %>%
  rename(unit             = unit_Swing,
         'data type'      = data_type_Swing,
         'indicator code' = Swing_indicator_code,
         name             = Swing_name,
         description      = Swing_description) %>%
  mutate(roundoff = as.double(roundoff),
         'provisional period' = NA_character_) %>%           # ontwikkelpunt: provisional period uit Indicatoren sheet halen
  select('indicator code', name, description, unit, 'provisional period', 'data type', roundoff, formula, bron) %>%
  bind_rows(metadata) %>%
  rename(source           = bron) %>%
  mutate(visible = 1)


rm(CBS_tabellen, data_list, formula, indicatoren, metadata_list, omzettabel_unit, CBS_tabel, GGD, 
   build_var_table, CBS_to_Swing_periodcodes, bron_lookup_CBS, formule_lookup, bepaal_bron_formula,
   CBS_tabellen_alle, indicatoren_alle, bron_per_onderwerp)


# Opslaan outputbestanden
write.csv2(data, "output/data.csv", row.names = FALSE)
write.csv2(metadata, "output/metadata.csv", row.names = FALSE)
