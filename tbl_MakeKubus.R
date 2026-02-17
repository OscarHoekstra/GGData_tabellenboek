# Functie voor het maken van platte kubusdata voor Swing (ABF).
# Doel: gestandaardiseerde output voor Swing, inclusief labels en dimensies.
library(data.table)

extract_subset <- function(args, data, dummy_crossing_var) {
  # Start met dataset filter
  idx <- which(data$tbl_dataset == args$tbl_dataset)
  if (length(idx) == 0) return(NULL)

  # Jaar filter indien van toepassing
  if (!is.na(args$jaarvariabele) && args$jaarvariabele %in% names(data)) {
    jaar_vals <- data[[args$jaarvariabele]][idx]
    idx <- idx[!is.na(jaar_vals) & jaar_vals == as.integer(args$jaar_voor_analyse)]
  }
  if (length(idx) == 0) return(NULL)

  # Bepaal benodigde kolommen voor deze specifieke config
  cols <- c(args$vars, args$current_weegfactor)
  if (!is.na(args$gebiedsindeling_kolom) && args$gebiedsindeling_kolom %in% names(data)) {
    cols <- c(cols, args$gebiedsindeling_kolom)
  }
  if (args$crossings != dummy_crossing_var && args$crossings %in% names(data)) {
    cols <- c(cols, args$crossings)
  }
  if (!is.na(args$jaarvariabele) && args$jaarvariabele %in% names(data)) {
    cols <- c(cols, args$jaarvariabele)
  }
  cols <- unique(cols[cols %in% names(data)])

  # Extraheer subset
  subset_df <- data[idx, cols, drop = FALSE]

  # Strip haven_labelled class voor verwerking
  for (col in names(subset_df)) {
    if (inherits(subset_df[[col]], "haven_labelled")) {
      subset_df[[col]] <- as.vector(subset_df[[col]])
    }
  }

  subset_df
}

process_single_config <- function(
  args,
  data,
  var_labels_cache,
  var_name_cache,
  crossing_labels_cache,
  crossing_name_cache,
  min_obs_vraag,
  min_obs_cel,
  afkapwaarde,
  output_folder,
  output_bestandsnaam_prefix,
  max_char_labels,
  dummy_crossing_var
) {

  bron <- args$bron
  is_kubus <- if (isTRUE(args$geen_crossings)) 0 else 1
  current_weegfactor <- args$current_weegfactor

  # Extract data subset on-the-fly om geheugen te sparen
  kubusdata <- extract_subset(args, data, dummy_crossing_var)
  if (is.null(kubusdata) || nrow(kubusdata) == 0) {
    return(list(success = TRUE, skipped = TRUE, var = args$vars, reason = "no_data"))
  }

  # Voeg geolevel toe
  kubusdata$geolevel <- args$gebiedsniveau

  # Jaar bepalen
  if (is.na(args$jaarvariabele)) {
    kubusdata$Jaar <- args$jaar_voor_analyse
  } else {
    kubusdata$Jaar <- kubusdata[[args$jaarvariabele]]
  }

  # Geoitem bepalen
  if (is.na(args$gebiedsindeling_kolom)) {
    kubusdata$geoitem <- 1
  } else {
    kubusdata$geoitem <- kubusdata[[args$gebiedsindeling_kolom]]
  }

  # Crossing opnemen als geen_crossings == FALSE
  use_crossing <- !isTRUE(args$geen_crossings) &&
    args$crossings != dummy_crossing_var &&
    (args$crossings %in% names(kubusdata))

  grouping_cols <- c("Jaar", "geolevel", "geoitem")
  if (use_crossing) {
    grouping_cols <- c(grouping_cols, args$crossings)
  }

  # Filter NA waarden
  keep_rows <- !is.na(kubusdata$geoitem) &
    !is.na(kubusdata[[args$vars]]) &
    !is.na(kubusdata[[current_weegfactor]])
  if (use_crossing) {
    keep_rows <- keep_rows & !is.na(kubusdata[[args$crossings]])
  }
  kubusdata <- kubusdata[keep_rows, , drop = FALSE]

  if (nrow(kubusdata) == 0) {
    return(list(success = TRUE, skipped = TRUE, var = args$vars, reason = "all_filtered"))
  }

  # Aggregatie met data.table
  dt <- data.table::as.data.table(kubusdata)
  by_cols <- c(grouping_cols, args$vars)
  wf_col <- current_weegfactor
  kubusdata <- dt[, .(n_gewogen = sum(get(wf_col), na.rm = TRUE),
                      n_ongewogen = .N),
                  by = by_cols]
  kubusdata <- tibble::as_tibble(kubusdata)

  if (nrow(kubusdata) == 0) {
    if (is_kubus) {
      out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), as.character(args$gebiedsniveau))
      bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, "_", args$crossings, ".xlsx"))
    } else {
      out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), paste0(as.character(args$gebiedsniveau), "_totaal"))
      bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, ".xlsx"))
    }
    return(list(success = TRUE, skipped = TRUE, var = args$vars, file = bestandsnaam, reason = "empty_after_agg"))
  }

  # Pivot naar breed formaat
  kubusdata <- kubusdata |>
    tidyr::pivot_wider(
      id_cols = dplyr::all_of(grouping_cols),
      names_from = dplyr::all_of(args$vars),
      values_from = c(n_gewogen, n_ongewogen),
      values_fill = 0,
      names_sep = "_"
    )

  # ONG berekenen met vectorized rowSums
  ongewogen_cols <- grep("^n_ongewogen_", names(kubusdata), value = TRUE)
  ong_col_name <- paste0(args$vars, "_ONG")
  if (length(ongewogen_cols) > 0) {
    kubusdata[[ong_col_name]] <- rowSums(kubusdata[, ongewogen_cols, drop = FALSE], na.rm = TRUE)
  } else {
    kubusdata[[ong_col_name]] <- 0
  }

  # Gewogen kolommen hernoemen
  gewogen_cols <- grep("^n_gewogen_", names(kubusdata), value = TRUE)
  if (length(gewogen_cols) > 0) {
    nieuwe_namen <- sub("^n_gewogen_", paste0(args$vars, "_"), gewogen_cols)
    names(kubusdata)[match(gewogen_cols, names(kubusdata))] <- nieuwe_namen
  }

  # Labels ophalen uit cache
  vl <- var_labels_cache[[args$vars]]
  ans_values <- as.character(unname(vl))
  ans_labels <- names(vl)

  antwoordkolommen <- paste0(args$vars, "_", ans_values)
  ongewogen_celkolommen <- paste0("n_ongewogen_", ans_values)
  ong_col <- paste0(args$vars, "_ONG")

  # Ontbrekende kolommen aanmaken
  ontbrekend_w <- setdiff(antwoordkolommen, names(kubusdata))
  if (length(ontbrekend_w) > 0) for (mc in ontbrekend_w) kubusdata[[mc]] <- 0
  ontbrekend_u <- setdiff(ongewogen_celkolommen, names(kubusdata))
  if (length(ontbrekend_u) > 0) for (mc in ontbrekend_u) kubusdata[[mc]] <- 0
  if (!ong_col %in% names(kubusdata)) kubusdata[[ong_col]] <- 0

  # Data kolommen bepalen
  base_cols <- c("Jaar", "geolevel", "geoitem")
  if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubusdata)) {
    base_cols <- c(base_cols, args$crossings)
  }
  data_cols <- c(base_cols, antwoordkolommen, ong_col)
  data_cols <- data_cols[data_cols %in% names(kubusdata)]

  # Privacy masking: rijniveau
  rij_te_weinig <- kubusdata[[ong_col]] < min_obs_vraag & kubusdata[[ong_col]] != afkapwaarde
  if (any(rij_te_weinig, na.rm = TRUE)) {
    kubusdata[rij_te_weinig, antwoordkolommen] <- afkapwaarde
    kubusdata[rij_te_weinig, ong_col] <- afkapwaarde
  }

  # Privacy masking: celniveau
  if (min_obs_cel > 0) {
    for (i in seq_along(ans_values)) {
      cel_u_col <- ongewogen_celkolommen[i]
      cel_w_col <- antwoordkolommen[i]
      if (cel_u_col %in% names(kubusdata) && cel_w_col %in% names(kubusdata)) {
        mask <- kubusdata[[cel_u_col]] < min_obs_cel & kubusdata[[cel_u_col]] != afkapwaarde
        if (any(mask, na.rm = TRUE)) {
          kubusdata[mask, cel_w_col] <- afkapwaarde
        }
      }
    }
  }

  # Data klaarzetten voor export
  kubus_df <- kubusdata[, data_cols, drop = FALSE]

  # Excel workbook opbouwen
  wb <- openxlsx::createWorkbook()

  # 1) Data sheet
  openxlsx::addWorksheet(wb, "Data")
  openxlsx::writeData(wb, "Data", kubus_df)
  style_fmt <- openxlsx::createStyle(numFmt = "[=-99996]0;0.000")
  openxlsx::addStyle(wb, "Data", style_fmt, cols = which(names(kubus_df) %in% antwoordkolommen),
                     rows = 2:(nrow(kubus_df) + 1), gridExpand = TRUE)

  # 2) Data_def sheet
  openxlsx::addWorksheet(wb, "Data_def")
  data_def_cols <- c("Jaar", "geolevel", "geoitem")
  if (!isTRUE(args$geen_crossings) && args$crossings %in% names(kubus_df)) {
    data_def_cols <- c(data_def_cols, args$crossings)
  }
  data_def_cols <- c(data_def_cols, antwoordkolommen, ong_col)
  data_def_cols <- data_def_cols[data_def_cols %in% names(kubus_df)]
  data_def <- data.frame(
    col = data_def_cols,
    type = dplyr::case_when(
      data_def_cols == "Jaar" ~ "period",
      data_def_cols == "geolevel" ~ "geolevel",
      data_def_cols == "geoitem" ~ "geoitem",
      data_def_cols == args$crossings ~ "dim",
      TRUE ~ "var"
    )
  )
  openxlsx::writeData(wb, "Data_def", data_def)

  # 3) Label_var sheet
  openxlsx::addWorksheet(wb, "Label_var")
  label_var <- data.frame(
    Onderwerpcode = c(antwoordkolommen, ong_col),
    Naam = c(paste0("Aantal ", ans_labels), "Totaal aantal ongewogen"),
    Eenheid = rep("Personen", length(antwoordkolommen) + 1)
  )
  openxlsx::writeData(wb, "Label_var", label_var)

  # 4) Dimensies + items (alleen met crossings)
  if (!isTRUE(args$geen_crossings) && !is.null(crossing_labels_cache[[args$crossings]])) {
    openxlsx::addWorksheet(wb, "Dimensies")
    dim_naam <- crossing_name_cache[[args$crossings]]
    openxlsx::writeData(
      wb, "Dimensies",
      data.frame(Dimensiecode = args$crossings, Naam = dim_naam)
    )

    openxlsx::addWorksheet(wb, args$crossings)
    cr_lbls <- crossing_labels_cache[[args$crossings]]
    df_items <- data.frame(
      Itemcode = unname(cr_lbls),
      Naam = names(cr_lbls),
      Volgnr = seq_along(cr_lbls)
    )
    openxlsx::writeData(wb, args$crossings, df_items)
  }

  # 5) Indicators sheet
  openxlsx::addWorksheet(wb, "Indicators")
  variabel_naam <- var_name_cache[[args$vars]]
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
                          if (is_dichotoom) min_obs_vraag else rep(min_obs_vraag, length(ans_values))),
    `Threshold Indicator` = c(rep("", length(ans_values) + 2),
                              if (is_dichotoom) paste0(args$vars, "_ONG") else rep(paste0(args$vars, "_ONG"), length(ans_values))),
    Cube = rep(1, length(indicator_codes)),
    Source = rep(bron, length(indicator_codes)),
    check.names = FALSE
  )
  openxlsx::writeData(wb, "Indicators", indicators_df)

  # Bestandsnaam bepalen en opslaan
  if (is_kubus) {
    out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), as.character(args$gebiedsniveau))
    if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
    bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, "_", args$crossings, ".xlsx"))
  } else {
    out_dir <- file.path(output_folder, as.character(args$jaar_voor_analyse), paste0(as.character(args$gebiedsniveau), "_totaal"))
    if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
    bestandsnaam <- file.path(out_dir, paste0(output_bestandsnaam_prefix, "_", args$vars, ".xlsx"))
  }

  openxlsx::saveWorkbook(wb, file = bestandsnaam, overwrite = TRUE)

  return(list(success = TRUE, skipped = FALSE, var = args$vars, file = bestandsnaam))
}

process_single_config_safe <- function(args, ...) {
  tryCatch(
    process_single_config(args = args, ...),
    error = function(e) {
      list(success = FALSE, var = args$vars, error = conditionMessage(e))
    }
  )
}

MaakKubusData <- function(
    data,
    configuraties,
    variabelen,
    crossings,
    algemeen = NULL,
    parallel = FALSE,
    workers = max(1, parallel::detectCores() - 1),
    future_globals_max_size = 4 * 1024^3,
    afkapwaarde = -99996,
    output_bestandsnaam_prefix = "kubusdata",
    output_folder = "output_swing",
    max_char_labels = 100,
    dummy_crossing_var = "dummy_crossing"
) {
  
  # Outputmap aanmaken indien nodig zodat schrijfacties nooit falen op ontbrekende paden.
  if (!dir.exists(output_folder)) dir.create(output_folder, recursive = TRUE)
  
  # Config uit de globale omgeving gebruiken wanneer die niet expliciet is meegegeven.
  # Dit behoudt backward compatibility met bestaande scripts.
  if (is.null(algemeen) && exists("algemeen", envir = .GlobalEnv)) {
    algemeen <- get("algemeen", envir = .GlobalEnv)
  }
  
  # Drempels voor privacy-masking: rijniveau (vraag) en celniveau (antwoord).
  min_obs_vraag <- if(!is.null(algemeen$min_observaties_per_vraag)) as.numeric(algemeen$min_observaties_per_vraag) else 0
  min_obs_cel   <- if(!is.null(algemeen$min_observaties_per_antwoord)) as.numeric(algemeen$min_observaties_per_antwoord) else 0
  
  # Afkapwaarde uit config overschrijft standaardwaarde zodat masking consistent is met Swing.
  if (!is.na(algemeen$swing_afkapwaarde)) {
    afkapwaarde <- as.numeric(algemeen$swing_afkapwaarde)
    msg("Afkapwaarde voor swing opgehaald uit config: %s", afkapwaarde, level = MSG)
  }
  
  # Variabelen voorbereiden en ontdubbelen om dubbele rijen in configs te voorkomen.
  vars_df <- variabelen
  if (any(duplicated(vars_df$variabelen))) {
    vars_df <- vars_df |> dplyr::distinct()
  }
  # Weegfactor kolommen uit variabelen hernoemen zodat ze niet botsen met configuratie
  # en altijd terug te vinden zijn in de config-rij.
  wf_cols <- grep("^weegfactor", names(vars_df), value=TRUE)
  if (length(wf_cols) > 0) {
    vars_df <- vars_df |> dplyr::rename_with(~paste0("var_", .), dplyr::all_of(wf_cols))
  }
  
  # Hernoemen van de variabele kolom naar 'vars' voor consistente verwijzingen.
  if ("variabelen" %in% names(vars_df)) {
    vars_df <- vars_df |> dplyr::rename(vars = variabelen)
  }
  
  vars <- unique(vars_df$vars)
  
  # Crossings normaliseren naar een karaktervector (ook wanneer input een data.frame is).
  if (missing(crossings) || is.null(crossings)) {
    crossings <- c()
  } else if (is.data.frame(crossings)) {
    crossings <- crossings[[1]]
  }
  crossings <- unique(crossings)
  
  # Dimensies uit configuratie ophalen om outputpaden te bouwen.
  gebiedsniveaus <- unique(configuraties$gebiedsniveau)
  jaren <- unique(configuraties$jaar_voor_analyse)
  
  # NA verwijderen zodat latere selects en filters geen fouten geven.
  vars <- vars[!is.na(vars)] 
  crossings <- crossings[!is.na(crossings)] 
  jaren <- jaren[!is.na(jaren)]
  
  
  # Check of variabelen bestaan om vroegtijdig te stoppen met duidelijke melding.
  vars_niet_in_data <- vars[!vars %in% names(data)]
  if (length(vars_niet_in_data) > 0) {
    msg("Variabele %s niet gevonden in data voor kubus export.\r\n", paste(vars_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  
  # Check of crossing bestaan
  crossings_niet_in_data <- crossings[!crossings %in% names(data)]
  if (length(crossings_niet_in_data) > 0 &&
      !(length(crossings_niet_in_data) == 1 && crossings_niet_in_data == dummy_crossing_var)) {
    msg("Crossings %s niet gevonden in data voor kubus export.\r\n", paste(crossings_niet_in_data, collapse = ", "), level = WARN)
    return()
  }
  
  if (any(crossings %in% vars)) {
    vars_in_crossing <- crossings[crossings %in% vars]
    msg("Crossings %s komt/komen ook voor in vars. Dit levert problemen op dus wordt niet als var verwerkt.\r\n", vars_in_crossing, level = WARN)
    rm(vars_in_crossing)
    vars <- vars[!vars %in% crossings]
  }
  
  
  # Controleer per gebiedsindeling of de kolom bestaat.
  # Niet-bestaande kolommen worden NA zodat downstream code voorspelbaar blijft.
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
  
  # Zorg dat outputmappen bestaan voor alle jaar/gebiedsniveau combinaties.
  for (jaar in jaren) {
    for (gebiedsniveau in gebiedsniveaus) {
      if (!dir.exists(file.path(output_folder, jaar, gebiedsniveau))){
        dir.create(file.path(output_folder, jaar, gebiedsniveau), recursive = TRUE)
        msg("Output folder gemaakt: %s", file.path(output_folder, jaar, gebiedsniveau), level = MSG)
      }
    }
  }
  
  
  # Dummy crossing toevoegen zodat ook zonder echte crossing een dimensie beschikbaar is.
  crossings <- unique(c(crossings, dummy_crossing_var))

  # Configs bouwen: elke combinatie van configuratie, variabele en crossing.
  configs <- tidyr::expand_grid(
    configuraties,
    vars_df,
    crossings
  )
  
  # Verwijder combinaties waarin geen_crossings geldt maar toch een echte crossing staat.
  configs <- configs |> 
    filter(!(geen_crossings & crossings != dummy_crossing_var))

  # Bij geen_crossings alles behalve dummy var weghalen
  configs <- configs |> 
    filter(!(geen_crossings == TRUE & crossings != dummy_crossing_var))
  
  # Bij wel crossings dummy var weghalen
  configs <- configs |> 
    filter(!(geen_crossings == FALSE & crossings == dummy_crossing_var))
  
  # Pre-cache labels voor variabelen en crossings (performance optimalisatie)
  var_labels_cache <- lapply(stats::setNames(vars, vars), function(v) {
    vl <- labelled::val_labels(data[[v]])
    if (is.null(vl) || length(vl) == 0) {
      unieke_vals <- sort(unique(data[[v]][!is.na(data[[v]])]))
      vl <- stats::setNames(as.numeric(unieke_vals), as.character(unieke_vals))
    }
    vl
  })
  
  var_name_cache <- lapply(stats::setNames(vars, vars), function(v) {
    nm <- labelled::var_label(data[[v]])
    if (is.null(nm) || is.na(nm)) nm <- v
    nm
  })
  
  crossing_labels_cache <- lapply(stats::setNames(crossings, crossings), function(cr) {
    if (!cr %in% names(data)) return(NULL)
    cr_lbls <- labelled::val_labels(data[[cr]])
    if (is.null(cr_lbls) || length(cr_lbls) == 0) {
      vals <- sort(unique(data[[cr]][!is.na(data[[cr]])]))
      cr_lbls <- stats::setNames(as.character(vals), as.character(vals))
    }
    cr_lbls
  })
  
  crossing_name_cache <- lapply(stats::setNames(crossings, crossings), function(cr) {
    if (!cr %in% names(data)) return(cr)
    nm <- labelled::var_label(data[[cr]])
    if (is.null(nm) || is.na(nm)) nm <- cr
    nm
  })
  
  # ============================================================================
  # PERFORMANCE OPTIMALISATIE: Lazy extractie per config (geen list van subsets)
  # ============================================================================
  
  # Bepaal weegfactor per config (nodig voor subset extractie)
  configs$current_weegfactor <- purrr::pmap_chr(
    list(
      configs$weegfactor,
      configs$naam_dataset,
      configs$tbl_dataset
    ),
    function(wf, naam_ds, tbl_ds) {
      ds_name_col <- paste0("var_weegfactor.d_", naam_ds)
      ds_id_col <- paste0("var_weegfactor.d", tbl_ds)
      global_col <- "var_weegfactor"
      
      if (ds_name_col %in% names(configs) && !is.na(configs[[ds_name_col]][1])) {
        return(configs[[ds_name_col]][1])
      } else if (ds_id_col %in% names(configs) && !is.na(configs[[ds_id_col]][1])) {
        return(configs[[ds_id_col]][1])
      } else if (global_col %in% names(configs) && !is.na(configs[[global_col]][1])) {
        return(configs[[global_col]][1])
      }
      wf
    }
  )

  # ============================================================================
  # Uitvoering: parallel of sequentieel
  # ============================================================================
  
  if (parallel && nrow(configs) > 1) {
    # Controleer of furrr beschikbaar is
    if (!requireNamespace("furrr", quietly = TRUE) || !requireNamespace("future", quietly = TRUE)) {
      msg("Packages 'furrr' en 'future' zijn vereist voor parallelle verwerking. Installeer met: install.packages(c('furrr', 'future'))", level = WARN)
      msg("Terugvallen op sequentiële verwerking...", level = MSG)
      parallel <- FALSE
    }
  }
  
  run_sequential <- function() {
    # Gebruik data.table threads in sequentiële modus
    data.table::setDTthreads(workers)

    msg("Start sequentiële verwerking voor %d configuraties...", nrow(configs), level = MSG)

    config_list <- purrr::transpose(as.list(configs))

    purrr::walk(
      config_list,
      function(args) {
        result <- tryCatch(
          process_single_config(
            args = args,
            data = data,
            var_labels_cache = var_labels_cache,
            var_name_cache = var_name_cache,
            crossing_labels_cache = crossing_labels_cache,
            crossing_name_cache = crossing_name_cache,
            min_obs_vraag = min_obs_vraag,
            min_obs_cel = min_obs_cel,
            afkapwaarde = afkapwaarde,
            output_folder = output_folder,
            output_bestandsnaam_prefix = output_bestandsnaam_prefix,
            max_char_labels = max_char_labels,
            dummy_crossing_var = dummy_crossing_var
          ),
          error = function(e) {
            msg("Fout bij verwerking van %s: %s", args$vars, conditionMessage(e), level = ERR)
            list(success = FALSE)
          }
        )

        if (result$success && !isTRUE(result$skipped)) {
          msg("Kubusdata opgeslagen voor %s: %s", args$vars, result$file, level = MSG)
        } else if (isTRUE(result$skipped)) {
          msg("Geen data voor kubus export: %s (reden: %s)", args$vars, result$reason, level = WARN)
        }
      },
      .progress = TRUE
    )
  }

  if (parallel && nrow(configs) > 1) {
    options(future.globals.maxSize = future_globals_max_size)
    
    # Converteer configs naar lijst van rijen voor pmap
    config_list <- purrr::transpose(as.list(configs))

    # Preflight: check grootte van globals om parallel fouten te voorkomen
    estimated_globals <- object.size(config_list) + object.size(process_single_config) + object.size(extract_subset) + object.size(data)
    if (estimated_globals > future_globals_max_size) {
      msg("Parallelle verwerking overgeslagen: globals zijn %.2f GiB (limiet %.2f GiB).", 
          as.numeric(estimated_globals) / 1024^3, future_globals_max_size / 1024^3, level = WARN)
      msg("Terugvallen op sequentiële verwerking...", level = MSG)
      parallel <- FALSE
    }
  }

  if (parallel && nrow(configs) > 1) {
    msg("Start parallelle verwerking met %d workers voor %d configuraties...", workers, nrow(configs), level = MSG)
    
    # Setup parallel backend
    future::plan(future::multisession, workers = workers)
    on.exit(future::plan(future::sequential), add = TRUE)
    
    # Voer parallel uit - elke worker krijgt alleen zijn eigen kleine data subset
    results <- tryCatch(
      furrr::future_map(
      config_list,
      process_single_config_safe,
      data = data,
      var_labels_cache = var_labels_cache,
      var_name_cache = var_name_cache,
      crossing_labels_cache = crossing_labels_cache,
      crossing_name_cache = crossing_name_cache,
      min_obs_vraag = min_obs_vraag,
      min_obs_cel = min_obs_cel,
      afkapwaarde = afkapwaarde,
      output_folder = output_folder,
      output_bestandsnaam_prefix = output_bestandsnaam_prefix,
      max_char_labels = max_char_labels,
      dummy_crossing_var = dummy_crossing_var,
      .options = furrr::furrr_options(
        globals = list(
          process_single_config_safe = process_single_config_safe,
          process_single_config = process_single_config,
          extract_subset = extract_subset,
          data = data,
          var_labels_cache = var_labels_cache,
          var_name_cache = var_name_cache,
          crossing_labels_cache = crossing_labels_cache,
          crossing_name_cache = crossing_name_cache,
          min_obs_vraag = min_obs_vraag,
          min_obs_cel = min_obs_cel,
          afkapwaarde = afkapwaarde,
          output_folder = output_folder,
          output_bestandsnaam_prefix = output_bestandsnaam_prefix,
          max_char_labels = max_char_labels,
          dummy_crossing_var = dummy_crossing_var
        ),
        packages = c("data.table", "tidyr", "dplyr", "tibble", "openxlsx"),
        seed = TRUE
      ),
      .progress = TRUE
      ),
      error = function(e) e
    )

    if (inherits(results, "error") || inherits(results, "simpleError")) {
      msg("Parallelle verwerking mislukt: %s", conditionMessage(results), level = WARN)
      msg("Terugvallen op sequentiële verwerking...", level = MSG)
      run_sequential()
      return(invisible(NULL))
    }
    
    # Rapporteer resultaten
    successes <- sum(sapply(results, function(x) x$success && !isTRUE(x$skipped)))
    skipped <- sum(sapply(results, function(x) x$success && isTRUE(x$skipped)))
    failures <- Filter(function(x) !x$success, results)
    
    msg("Parallelle verwerking voltooid: %d succesvol, %d overgeslagen, %d mislukt", 
        successes, skipped, length(failures), level = MSG)
    
    if (length(failures) > 0) {
      msg("Mislukte exports:", level = WARN)
      for (f in failures) {
        msg("  - %s: %s", f$var, f$error, level = WARN)
      }
    }
    
  } else {
    run_sequential()
  }
  
  invisible(NULL)
}
