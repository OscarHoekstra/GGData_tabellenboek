# Functie voor het maken van platte kubusdata voor Swing (ABF)
# Gebaseerd op de logica uit swing_kubusdata_wizard/global.R

MaakKubusData <- function(
    data,
    configuraties,
    variabelen,
    output_bestandsnaam_prefix = "kubusdata",
    output_folder = "swing_output"
) {
  
  # # Instellingen (overgenomen uit global.R of default voor platte data)
  # min_observaties <- 0
  # min_observaties_per_cel <- 0
  # missing_voor_privacy <- -99996
  # max_char_labels <- 100
  # bron <- "GGD" # Default bronvermelding
  # is_kubus <- 0 # 0 want platte data (geen crossings)
  dummy_crossing_var <- "dummy_crossing"
  
  # Check of output map bestaat
  if (!dir.exists(output_folder)) dir.create(output_folder, recursive = TRUE)
  
  
  vars <- unique(variabelen$variabelen)
  crossings <- unique(variabelen$crossings)
  
  gebiedsniveaus <- unique(configuraties$gebiedsniveau)
  # gebiedsindelingen <- unique(configuraties$gebiedsindeling_kolom)
  jaren <- unique(configuraties$jaar_voor_analyse)
  # TODO: Misschien moeten we controleren of variabelen/crossings uniek zijn (ipv uniek maken)
  
  # NA verwijderen
  vars <- vars[!is.na(vars)] 
  crossings <- crossings[!is.na(crossings)] 
  jaren <- jaren[!is.na(jaren)]
  
  
  # Check of variabelen bestaan
  vars_niet_in_data <- vars[!vars %in% names(data)]
  if (length(vars_niet_in_data) > 0) {
    msg("Variabele %s niet gevonden in data voor kubus export.\r\n", paste(vars_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  
  # Check of crossing bestaan
  crossings_niet_in_data <- vars[!vars %in% names(data)]
  if (length(crossings_niet_in_data) > 0 |
      (length(crossings_niet_in_data) == 1 && crossings_niet_in_data == dummy_crossing_var)) {
    msg("Crossings %s niet gevonden in data voor kubus export.\r\n", paste(crossings_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  # TODO: Controleren of variabelen niet in crossings voorkomen etc.
  
  
  
  # Controleer voor elke waarde in swing_configuraties$gebiedsindeling_kolom 
  # of deze kolomnaam voorkomt in data. Vervang kolomnamen die niet bestaan 
  # door NA en geef een waarschuwing. Houd de oorspronkelijke vectorstructuur 
  # (inclusief bestaande NA's) exact intact
  gebiedsindelingen_kolom <- sapply(configuraties$gebiedsindeling_kolom, function(kolomnaam) {
    if (is.na(kolomnaam)) {
      return(NA)
    } else if (!(kolomnaam %in% names(data))) {
      msg("Gebiedsindeling_kolom `%s` niet gevonden in data voor kubus export. Deze wordt als leeg beschouwd.\r\n", 
          kolomnaam, level = WARN)
      return(NA)
    } else {
      return(kolomnaam)
    }
  })
  
  for (jaar in jaren) {
    for (gebiedsniveau in gebiedsniveaus) {
      if (!dir.exists(file.path(output_folder, jaar, gebiedsniveau))){
        dir.create(file.path(output_folder, jaar, gebiedsniveau), recursive = TRUE)
        msg("Output folder gemaakt: %s", file.path(output_folder, jaar, gebiedsniveau), level = MSG)
      }
    }
  }
  
  
  # dummy_crossing toevoegen (als die er nog niet in zat) om lege crossings te vermijden
  crossings <- unique(c(crossings, dummy_crossing_var)) 
  
  # Combinatie maken tussen elke configuratie regel, variabele en crossing variabelen
  configs <- expand_grid(
    configuraties,
    vars,
    crossings
  )
  
  # Bij geen_crossings alles behalve dummy var weghalen
  configs <- configs |> 
    filter(!(isTRUE(geen_crossings) & crossings != dummy_crossing_var))
  
  
  # Over elke regel van de config lopen om de platte kubusdata te maken
  configs |> 
    pwalk(function(...){
      args <- list(...)
      
      # Kopie maken om orginele data niet aan te tasten
      kubusdata <- data |> 
        filter(tbl_dataset == args$tbl_dataset) |> 
        mutate(geolevel = args$gebiedsniveau)
      
      
      if (is.na(args$jaarvariabele)) {
        kubusdata <- kubusdata |> 
          mutate(Jaar = args$jaar_voor_analyse)
        msg("jaarvariabele is NA, dus jaar wordt ingesteld op %s", args$jaar_voor_analyse)
      } else {
        kubusdata <- kubusdata |> 
          mutate(Jaar = .data[[args$jaarvariabele]]) |> 
          filter(Jaar == as.integer(args$jaar_voor_analyse))
      }
      
      if (isTRUE(args$geen_crossings)) {
        # TODO: geen dummy_crossing en Dimensies tabbladen aanmaken
      }
      
      # Als Gebiedsindeling_kolom is NA, dan omzetten naar een 1
      if (is.na(args$gebiedsindeling_kolom)) {
        kubusdata <- kubusdata |> 
          mutate(geoitem = 1)
      } else {
        kubusdata <- kubusdata |> 
          mutate(geoitem = .data[[args$gebiedsindeling_kolom]])
      }
      
      # necessary_cols <- c(args$jaarvariabele, args$gebiedsindeling, args$vars, args$crossings, args$weegfactor)
      # necessary_cols <- necessary_cols[!is.na(necessary_cols)]
      # 
      # kubusdata <- kubusdata |> 
      #   select(all_of(necessary_cols))
      
      
      grouping_cols <- c("Jaar", "geolevel", "geoitem")
      if (!args$geen_crossings) {
        grouping_cols <- c(grouping_cols, args$crossings)
      }
      
      # if (!is.na(args$gebiedsindeling_kolom)) {
      #   kubusdata <- kubusdata |> 
      #     filter(!is.na(.data[[args$gebiedsindeling]]))
      # }
      
      kubusdata <- kubusdata |>
        filter(!is.na(geoitem),
               !is.na(.data[[args$vars]]),
               !is.na(.data[[args$crossings]]),
               !is.na(.data[[args$weegfactor]]))
      
      browser()
      
      kubusdata <- kubusdata |> 
        group_by(across(all_of(c(grouping_cols, args$vars)))) |> 
        summarise(
          n_gewogen = sum(!!sym(args$weegfactor), na.rm = TRUE),
          n_ongewogen = n(),
          .groups = "drop"
        )
      
      if (nrow(kubusdata) == 0) {
        msg("Geen data over voor kubus export van variabele %s (mogelijk alles missing).", variabele, level = WARN)
        return()
      }
      
      # Pivot naar breed formaat (platte structuur: 1 rij per gebied/jaar, kolommen per antwoord)
      kubusdata <- kubusdata |>
        group_by(across(all_of(grouping_cols))) |> 
        mutate(!!paste0(args$vars, "_ONG") := sum(n_ongewogen, na.rm = TRUE)) |> 
        ungroup() |> 
        pivot_wider(
          # Keep Year, Group, and the new ONG column fixed
          id_cols = c(grouping_cols, ends_with("_ONG")), 
          names_from = all_of(args$vars),
          values_from = n_gewogen,
          # Add the variable name as a prefix (AGETS411_0, AGETS411_1)
          names_prefix = paste0(args$vars, "_"), 
          values_fill = 0
        )
      
      
      
    },
    .progress = TRUE
    )
  
  
  
  
  # # Data voorbereiden en aggregatie
  # # We groeperen op jaar, gebied en het antwoord op de vraag (variabele)
  # kubus_df <- data |>
  #   filter(!is.na(.[[variabele]]), 
  #          !is.na(.[[gebiedsindeling]]), 
  #          !is.na(tbl_weegfactor)) |>
  #   mutate(var = factor(.[[variabele]], 
  #                       levels = val_labels(.[[variabele]]), 
  #                       labels = names(val_labels(.[[variabele]])))) |>
  #   group_by(tbl_jaar, .[[gebiedsindeling]], var) |>
  #   summarise(
  #     n_gewogen = sum(tbl_weegfactor, na.rm = TRUE),
  #     n_ongewogen = n(),
  #     .groups = "drop"
  #   )
  

  
  # Pivot naar breed formaat (platte structuur: 1 rij per gebied/jaar, kolommen per antwoord)
  # Logica overgenomen uit global.R (geen_crossings == TRUE tak)
  kubus_df <- kubus_df |>
    mutate(across(.cols = n_ongewogen, min, .names = "n_cel")) |> # min gebruiken als placeholder
    pivot_wider(names_from = var, values_from = c(n_gewogen, n_ongewogen), values_fill = 0) |>
    rowwise() |>
    # Totaal ongewogen berekenen door sommeren van de unpivoted kolommen
    mutate(n_ongewogen = rowSums(across(starts_with("n_ongewogen")))) |>
    select(-starts_with("n_ongewogen_")) |>
    ungroup() |>
    # Gewogen kolommen hernoemen (prefix 'n_gewogen_' verwijderen)
    rename_with(~str_sub(.x, start=11), starts_with("n_gewogen")) |>
    mutate(geolevel = "gemeente", .after = 1) # Aanname: niveau is gemeente (of regio), Swing vereist geolevel
  
  # Kolomnamen en volgorde herstellen
  # Volgorde forceren op basis van labels
  volgorde_labels <- names(val_labels(data[[variabele]]))
  
  # Huidige kolomindexen zoeken
  volgorde_labels_in_df <- sapply(volgorde_labels, function(x) which(names(kubus_df) == x))
  
  # Correcte volgorde samenstellen: Jaar, Geolevel, Geoitem, [Antwoorden], Totaal Ongewogen
  # Let op: afhankelijk van de pivot kan de volgorde variëren, dus we bouwen hem expliciet op
  cols_base <- c("tbl_jaar", "geolevel", gebiedsindeling)
  cols_answers <- names(val_labels(data[[variabele]])) # De antwoord categorieën
  cols_total <- c("n_ongewogen")
  
  # Check of alle kolommen bestaan (soms vallen categorieën weg als er geen data is, pivot vult aan met 0 door values_fill maar levels moeten matchen)
  # De pivot_wider maakt kolommen aan voor alle FACTOR levels die in de data zitten. Als levels empty zijn in de subset, maakt pivot ze niet aan tenzij drop=FALSE
  # We moeten zeker weten dat de kolommen matchen met de labels
  
  missing_cols <- setdiff(cols_answers, names(kubus_df))
  if(length(missing_cols) > 0){
    for(mc in missing_cols) kubus_df[[mc]] <- 0
  }
  
  kubus_df <- kubus_df |> select(all_of(c(cols_base, cols_answers, cols_total)))
  
  # Kolomnamen conform Swing formaat: Variabele_Label
  nieuwe_namen_vars <- glue("{variabele}_{cols_answers}")
  names(kubus_df) <- c("Period", "geolevel", "geoitem", nieuwe_namen_vars, glue("{variabele}_ONG"))
  
  # Privacy checks (kleine aantallen verwijderen)
  verwijder_kleine_aantallen <- function(x, ongewogen){
    if(ongewogen < min_observaties & ongewogen != missing_voor_privacy){missing_voor_privacy}else{x}
  }
  
  # Toepassen privacy regels (in dit geval op basis van totaal ongewogen per rij)
  # In global.R wordt dit complexer gedaan, hier vereenvoudigd voor platte data
  # Als n_ongewogen < threshold, dan alle waarden op missing
  kubus_df <- kubus_df |>
    mutate(across(.cols = all_of(c(nieuwe_namen_vars, glue("{variabele}_ONG"))),
                  .fns = ~ifelse(n_ongewogen < 0, missing_voor_privacy, .x))) # Placeholder logica, pas aan indien nodig
  
  
  # Excel bestand opbouwen
  wb <- createWorkbook()
  
  # 1. Sheet Data
  addWorksheet(wb, sheetName = "Data")
  writeData(wb, "Data", kubus_df)
  
  # 2. Sheet Data_def
  addWorksheet(wb, sheetName = "Data_def")
  data_def <- data.frame(
    col = names(kubus_df),
    type = c("period", "geolevel", "geoitem", rep("var", length(cols_answers) + 1))
  )
  writeData(wb, "Data_def", data_def)
  
  # Variabele metadata ophalen
  variabel_labels <- val_labels(data[[variabele]])
  is_dichotoom <- all(unname(variabel_labels) %in% c(0,1))
  variabel_naam <- var_label(data[[variabele]])
  if(is.null(variabel_naam)) variabel_naam <- variabele
  
  n_labels <- length(variabel_labels)
  
  # 3. Sheet Label_var
  addWorksheet(wb, sheetName = "Label_var")
  label_var <- data.frame(
    Onderwerpcode = c(glue("{variabele}_{names(variabel_labels)}"), glue("{variabele}_ONG")),
    Naam = c(glue("Aantal {names(variabel_labels)}"), "Totaal aantal ongewogen"),
    Eenheid = "Personen"
  )
  writeData(wb, "Label_var", label_var)
  
  # 4. Sheet Indicators
  addWorksheet(wb, sheetName = "Indicators")
  
  # Indicator codes opbouwen
  indicator_codes <- c(
    glue("{variabele}_{names(variabel_labels)}"), # Aantallen
    glue("{variabele}_ONG"), # Ongewogen totaal
    glue("{variabele}_GEW"), # Gewogen totaal (wordt berekend in formula)
    if(is_dichotoom) { glue("{variabele}_perc") } else { glue("{variabele}_{names(variabel_labels)}_perc") } # Percentages
  )
  
  # Indicator names
  indicator_names <- c(
    glue("Aantal {names(variabel_labels)}"),
    "Totaal aantal ongewogen",
    "Totaal aantal gewogen",
    if(is_dichotoom) { substr(variabel_naam, 1, max_char_labels) } else { glue("{substr(variabel_naam,1,45)}, {names(variabel_labels)}") }
  )
  
  # Formules
  som_formule <- paste(glue("{variabele}_{names(variabel_labels)}"), collapse = "+")
  
  formulas <- c(
    rep("", n_labels + 1), # Leeg voor aantallen + ongewogen
    som_formule, # Gewogen totaal
    if(is_dichotoom) {
      glue("({variabele}_1/({som_formule}))*100")
    } else {
      glue("({variabele}_{names(variabel_labels)}/({som_formule}))*100")
    }
  )
  
  indicators_df <- data.frame(
    `Indicator code` = indicator_codes,
    Name = indicator_names,
    Unit = c(rep("personen", n_labels + 2), if(is_dichotoom) "percentage" else rep("percentage", n_labels)),
    `Aggregation indicator` = c(rep("", n_labels + 2), if(is_dichotoom) glue("{variabele}_GEW") else rep(glue("{variabele}_GEW"), n_labels)),
    Formula = formulas,
    `Data type` = c(rep("Numeric", n_labels + 2), if(is_dichotoom) "Percentage" else rep("Percentage", n_labels)),
    Visible = c(rep(0, n_labels + 2), if(is_dichotoom) 1 else rep(1, n_labels)),
    `Threshold value` = c(rep("", n_labels + 2), if(is_dichotoom) min_observaties else rep(min_observaties, n_labels)),
    `Threshold Indicator` = c(rep("", n_labels + 2), if(is_dichotoom) glue("{variabele}_ONG") else rep(glue("{variabele}_ONG"), n_labels)),
    Cube = is_kubus,
    Source = bron,
    check.names = FALSE
  )
  
  writeData(wb, "Indicators", indicators_df)
  
  # Opslaan
  bestandsnaam <- glue("{output_folder}/kubus_{variabele}.xlsx")
  tryCatch({
    saveWorkbook(wb, file = bestandsnaam, overwrite = TRUE)
    msg("Kubusdata opgeslagen voor %s: %s", variabele, bestandsnaam, level = MSG)
  }, error = function(e) {
    msg("Fout bij opslaan kubusdata voor %s: %s", variabele, e$message, level = ERR)
  })
}
