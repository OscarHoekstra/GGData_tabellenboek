# Functie voor het maken van platte kubusdata voor Swing (ABF)
# Gebaseerd op de logica uit swing_kubusdata_wizard/global.R

MaakKubusData <- function(
    data,
    configuraties,
    variabelen,
    crossings,
    output_bestandsnaam_prefix = "kubusdata",
    output_folder = "output_swing",
    missing_voor_privacy = -99996,
    max_char_labels = 100,
    dummy_crossing_var = "dummy_crossing"
) {
  
  # Check of output map bestaat
  if (!dir.exists(output_folder)) dir.create(output_folder, recursive = TRUE)
  
  
  vars <- unique(variabelen$variabelen)
  
  if (missing(crossings) || is.null(crossings)) {
    crossings <- c()
  } else if (is.data.frame(crossings)) {
    crossings <- crossings[[1]]
  }
  crossings <- unique(crossings)
  
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
  crossings_niet_in_data <- crossings[!crossings %in% names(data)]
  if (length(crossings_niet_in_data) > 0 |
      (length(crossings_niet_in_data) == 1 && crossings_niet_in_data == dummy_crossing_var)) {
    msg("Crossings %s niet gevonden in data voor kubus export.\r\n", paste(crossings_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  
  if (any(crossings %in% vars)) {
    vars_in_crossing <- crossings[crossings %in% vars]
    msg("Crossings %s komt/komen ook voor in vars. Dit levert problemen op dus wordt niet als var verwerkt.\r\n", vars_in_crossing, level = WARN)
    rm(vars_in_crossing)
    vars <- vars[!vars %in% crossings]
  }
  
  
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
  
  # Verwijder alle rows waar geen_crossing maar toch wel crossing
  configs <- configs |> 
    filter(!(geen_crossings & crossings != dummy_crossing_var))
  
  # Bij geen_crossings alles behalve dummy var weghalen
  configs <- configs |> 
    filter(!(geen_crossings == TRUE & crossings != dummy_crossing_var))
  
  # Bij wel crossings dummy var weghalen
  configs <- configs |> 
    filter(!(geen_crossings == FALSE & crossings == dummy_crossing_var))
  
  # Over elke regel van de config lopen om de platte kubusdata te maken
  configs |> 
    pwalk(function(...){
      args <- list(...)
      
      min_observaties <- algemeen$min_observaties_per_vraag
      min_observaties_per_cel <- algemeen$min_observaties_per_antwoord
      bron <- args$bron
      is_kubus <- if (isTRUE(args$geen_crossings)) 0 else 1
      
      
      
      
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
      
      
      # Crossing opnemen als geen_crossings == FALSE en deze bestaat (anders niet groeperen op crossing)
      use_crossing <- !isTRUE(args$geen_crossings) &&
        args$crossings != dummy_crossing_var &&
        (args$crossings %in% names(kubusdata))
      
      grouping_cols <- c("Jaar", "geolevel", "geoitem")
      if (use_crossing) {
        grouping_cols <- c(grouping_cols, args$crossings)
      }
      
      kubusdata <- kubusdata |>
        dplyr::filter(!is.na(.data$geoitem),
                      !is.na(.data[[args$vars]]),
                      !is.na(.data[[args$weegfactor]]))
      
      # Alleen crossing filteren wanneer use_crossing == TRUE
      if (use_crossing) {
        kubusdata <- kubusdata |>
          dplyr::filter(!is.na(.data[[args$crossings]]))
      }
      
      kubusdata <- kubusdata |> 
        group_by(across(all_of(c(grouping_cols, args$vars)))) |> 
        summarise(
          n_gewogen = sum(!!sym(args$weegfactor), na.rm = TRUE),
          n_ongewogen = n(),
          .groups = "drop"
        )
      
      if (nrow(kubusdata) == 0) {
        if (is_kubus) {
          out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), as.character(args$gebiedsniveau))
          bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, "_", args$crossings, ".xlsx"))
        } else {
          out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), paste0(as.character(args$gebiedsniveau), "_totaal"))
          bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, ".xlsx"))
        }
        msg("Geen data over voor kubus export voor variabele %s: %s (mogelijk alles missing).",args$vars, bestandsnaam, level = WARN)
        return()
      }
      
      # Pivot naar breed: zowel gewogen als ongewogen meenemen (voor celmaskering)
      kubusdata <- kubusdata |>
        tidyr::pivot_wider(
          id_cols = dplyr::all_of(grouping_cols),
          names_from = dplyr::all_of(args$vars),
          values_from = c(n_gewogen, n_ongewogen),
          values_fill = 0,
          names_sep = "_"
        )
      
      # Herbereken totaal ongewogen ONG als som van n_ongewogen_* kolommen
      ongewogen_cols <- grep("^n_ongewogen_", names(kubusdata), value = TRUE)
      if (length(ongewogen_cols) > 0) {
        kubusdata <- kubusdata |>
          dplyr::rowwise() |>
          dplyr::mutate(!!paste0(args$vars, "_ONG") := sum(dplyr::c_across(dplyr::all_of(ongewogen_cols)), na.rm = TRUE)) |>
          dplyr::ungroup()
      } else {
        # als er om wat voor reden dan ook geen ongewogen kolommen zijn, fallback op 0
        kubusdata[[paste0(args$vars, "_ONG")]] <- 0
      }
      
      # Gewogen kolommen hernoemen naar {variabele}_{waarde} (n_gewogen_0 -> VAR_0)
      gewogen_cols <- grep("^n_gewogen_", names(kubusdata), value = TRUE)
      if (length(gewogen_cols) > 0) {
        nieuwe_namen <- sub("^n_gewogen_", paste0(args$vars, "_"), gewogen_cols)
        names(kubusdata)[match(gewogen_cols, names(kubusdata))] <- nieuwe_namen
      }
      
      # Labels van de inhoudelijke variabele
      vl <- labelled::val_labels(data[[args$vars]])
      if (is.null(vl) || length(vl) == 0) {
        unieke_vals <- sort(unique(data[[args$vars]]))
        vl <- stats::setNames(as.numeric(unieke_vals), as.character(unieke_vals))
      }
      ans_values <- as.character(unname(vl))   # codes zoals "0","1","2"
      ans_labels <- names(vl)                  # labels zoals "Ja","Nee",...
      
      # Lijsten met kolommen
      antwoordkolommen <- paste0(args$vars, "_", ans_values)     # gewogen data-kolommen
      ongewogen_celkolommen <- paste0("n_ongewogen_", ans_values) # per-antwoord ongewogen aantallen (alleen voor maskering)
      ong_col <- paste0(args$vars, "_ONG")
      
      # Zorg dat alle kolommen aanwezig zijn (ontbrekende aanmaken)
      ontbrekend_w <- setdiff(antwoordkolommen, names(kubusdata))
      if (length(ontbrekend_w) > 0) for (mc in ontbrekend_w) kubusdata[[mc]] <- 0
      ontbrekend_u <- setdiff(ongewogen_celkolommen, names(kubusdata))
      if (length(ontbrekend_u) > 0) for (mc in ontbrekend_u) kubusdata[[mc]] <- 0
      if (!ong_col %in% names(kubusdata)) kubusdata[[ong_col]] <- 0
      
      # Basis-kolommen voor Data
      base_cols <- c("Jaar", "geolevel", "geoitem")
      if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubusdata)) {
        base_cols <- c(base_cols, args$crossings)
      }
      
      # Filter op volgorde en houd alleen Swing-kolommen in Data (ongewogen celniveau kolommen blijven alleen voor maskering)
      data_cols <- c(base_cols, antwoordkolommen, ong_col)
      data_cols <- data_cols[data_cols %in% names(kubusdata)]
      
      # Privacy: rijniveau (ONG < min_observaties) -> alle waarden missing
      rij_te_weinig <- kubusdata[[ong_col]] < min_observaties & kubusdata[[ong_col]] != missing_voor_privacy
      if (any(rij_te_weinig, na.rm = TRUE)) {
        kubusdata[rij_te_weinig, antwoordkolommen] <- missing_voor_privacy
        kubusdata[rij_te_weinig, ong_col] <- missing_voor_privacy
      }
      
      # Privacy: celniveau (n_ongewogen_<code> < min_observaties_per_cel) -> alleen die cel missing
      if (min_observaties_per_cel > 0) {
        for (i in seq_along(ans_values)) {
          cel_u_col <- ongewogen_celkolommen[i]
          cel_w_col <- antwoordkolommen[i]
          if (cel_u_col %in% names(kubusdata) && cel_w_col %in% names(kubusdata)) {
            mask <- kubusdata[[cel_u_col]] < min_observaties_per_cel & kubusdata[[cel_u_col]] != missing_voor_privacy
            if (any(mask, na.rm = TRUE)) {
              kubusdata[mask, cel_w_col] <- missing_voor_privacy
            }
          }
        }
      }
      
      # Data klaarzetten (zonder de n_ongewogen_* hulpkolommen)
      kubus_df <- kubusdata |>
        dplyr::select(dplyr::all_of(data_cols))
      
      # Excel workbook opbouwen
      wb <- openxlsx::createWorkbook()
      
      # 1) Data
      openxlsx::addWorksheet(wb, "Data")
      openxlsx::writeData(wb, "Data", kubus_df)
      openxlsx::addStyle(wb, "Data", openxlsx::createStyle(numFmt = "0.000"), cols = which(names(kubus_df) %in% antwoordkolommen), 
                         rows = 2:(nrow(kubus_df) + 1), gridExpand = TRUE)
      
      # 2) Data_def
      openxlsx::addWorksheet(wb, "Data_def")
      data_def_cols <- c("Jaar", "geolevel", "geoitem")
      if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubus_df)) {
        data_def_cols <- c(data_def_cols, args$crossings)
      }
      data_def_cols <- c(data_def_cols, antwoordkolommen, ong_col)
      data_def_cols <- data_def_cols[data_def_cols %in% names(kubus_df)]
      data_def <- data.frame(
        col = data_def_cols,
        type = c(
          "period", "geolevel", "geoitem",
          if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubus_df)) "dim" else NULL,
          rep("var", sum(data_def_cols %in% antwoordkolommen) + as.integer(ong_col %in% data_def_cols))
        )
      )
      openxlsx::writeData(wb, "Data_def", data_def)
      
      # 3) Label_var
      openxlsx::addWorksheet(wb, "Label_var")
      label_var <- data.frame(
        Onderwerpcode = c(antwoordkolommen, ong_col),
        Naam = c(paste0("Aantal ", ans_labels), "Totaal aantal ongewogen"),
        Eenheid = rep("Personen", length(antwoordkolommen) + 1)
      )
      openxlsx::writeData(wb, "Label_var", label_var)
      
      # 4) Dimensies + items (alleen met crossings)
      if (!isTRUE(args$geen_crossings) && args$crossings %in% colnames(data)) {
        openxlsx::addWorksheet(wb, "Dimensies")
        dim_naam <- labelled::var_label(data[[args$crossings]])
        if (is.null(dim_naam) || is.na(dim_naam)) dim_naam <- args$crossings
        openxlsx::writeData(
          wb, "Dimensies",
          data.frame(Dimensiecode = args$crossings, Naam = dim_naam)
        )
        
        openxlsx::addWorksheet(wb, args$crossings)
        if (args$crossings %in% names(data) && "labelled" %in% class(data[[args$crossings]])) {
          cr_lbls <- labelled::val_labels(data[[args$crossings]])
        } else {
          cr_lbls <- NULL
        }
        if (is.null(cr_lbls) || length(cr_lbls) == 0) {
          vals <- sort(unique(data[[args$crossings]]))
          cr_lbls <- stats::setNames(as.character(vals), as.character(vals))
        }
        df_items <- data.frame(
          Itemcode = unname(cr_lbls),
          Naam = names(cr_lbls),
          Volgnr = seq_along(cr_lbls)
        )
        openxlsx::writeData(wb, args$crossings, df_items)
      }
      
      # 5) Indicators
      openxlsx::addWorksheet(wb, "Indicators")
      variabel_naam <- labelled::var_label(data[[args$vars]])
      if (is.null(variabel_naam) || is.na(variabel_naam)) variabel_naam <- args$vars
      
      is_dichotoom <- suppressWarnings(all(as.numeric(ans_values) %in% c(0,1)) && length(ans_values) == 2)
      
      indicator_codes <- c(
        antwoordkolommen,
        ong_col,
        paste0(args$vars, "_GEW"),
        if (is_dichotoom) paste0(args$vars, "_perc") else paste0(args$vars, "_", ans_values, "_perc")
      )
      indicator_names <- c(
        paste0("Aantal ", ans_labels),
        "Totaal aantal ongewogen",
        "Totaal aantal gewogen",
        if (is_dichotoom) substr(variabel_naam, 1, max_char_labels)
        else {
          nm <- paste0(variabel_naam, ", ", ans_labels)
          ifelse(nchar(nm) > max_char_labels, substr(nm, 1, max_char_labels), nm)
        }
      )
      som_formule <- paste(antwoordkolommen, collapse = "+")
      if (is_dichotoom) {
        perc_formulas <- paste0("(", args$vars, "_1/(", som_formule, "))*100")
        formulas <- c(rep("", length(ans_values) + 1), som_formule, perc_formulas)
      } else {
        perc_formulas <- vapply(ans_values, function(v) paste0("(", args$vars, "_", v, "/(", som_formule, "))*100"), character(1))
        formulas <- c(rep("", length(ans_values) + 1), som_formule, perc_formulas)
      }
      
      indicators_df <- data.frame(
        `Indicator code` = indicator_codes,
        Name = indicator_names,
        Unit = c(rep("personen", length(ans_values) + 2),
                 if (is_dichotoom) "percentage" else rep("percentage", length(ans_values))),
        `Aggregation indicator` = c(rep("", length(ans_values) + 2),
                                    if (is_dichotoom) paste0(args$vars, "_GEW") else rep(paste0(args$vars, "_GEW"), length(ans_values))),
        Formula = formulas,
        `Data type` = c(rep("Numeric", length(ans_values) + 2),
                        if (is_dichotoom) "Percentage" else rep("Percentage", length(ans_values))),
        Visible = c(rep(0, length(ans_values) + 2),
                    if (is_dichotoom) 1 else rep(1, length(ans_values))),
        `Threshold value` = c(rep("", length(ans_values) + 2),
                              if (is_dichotoom) min_observaties else rep(min_observaties, length(ans_values))),
        `Threshold Indicator` = c(rep("", length(ans_values) + 2),
                                  if (is_dichotoom) paste0(args$vars, "_ONG") else rep(paste0(args$vars, "_ONG"), length(ans_values))),
        Cube = c(rep(is_kubus, length(ans_values) + 2),
                 if (is_dichotoom) is_kubus else rep(is_kubus, length(ans_values))),
        Source = rep(bron, length(indicator_codes)),
        check.names = FALSE
      )
      openxlsx::writeData(wb, "Indicators", indicators_df)
      
      # Opslaan in submap {output_folder}/{jaar}/{gebiedsniveau}
      
      if (is_kubus) {
        out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), as.character(args$gebiedsniveau))
        if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
        bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, "_", args$crossing, ".xlsx"))
      } else {
        out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), paste0(as.character(args$gebiedsniveau), "_totaal"))
        if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
        bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, ".xlsx"))
      }
      
      tryCatch({
        openxlsx::saveWorkbook(wb, file = bestandsnaam, overwrite = TRUE)
        msg("Kubusdata opgeslagen voor %s: %s", args$vars, bestandsnaam, level = MSG)
      }, error = function(e) {
        msg("Fout bij opslaan kubusdata voor %s: %s", args$vars, e$message, level = ERR)
      })

    },
    .progress = TRUE
    )
  
}
