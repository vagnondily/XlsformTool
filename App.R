# =====================================================================
# app.R — XLSForm : Validation + Conversion Word (Stable)
# - Résumé global aligné
# - Récap sections/sous-sections avec nombre de questions
# - Génération Word (ignore end_group/end_repeat)
# - Détection dynamique des colonnes label et hint (gère label::/hint::English (en), French (fr))
# - Vérification de structure (unmatched begin/end group/repeat)
# - Progress bar pour analyse et génération Word
# =====================================================================

# --- Packages requis ---
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(stringr); library(tidyr)
  library(purrr);  library(officer); library(flextable); library(glue); library(tools)
  library(tibble); library(rlang); library(shiny); library(DT); library(writexl); library(htmltools);library(shinythemes)
})

# =====================================================================
# XLSForm -> Word (Rendu Word) - Script V1.0
# =====================================================================

# Opérateur de coalescence robuste pour gérer les vecteurs
`%||%` <- function(a, b) {
  if (is.null(a) || length(a) == 0 || (is.atomic(a) && length(a) == 1 && (is.na(a) || !nzchar(a)))) {
    b
  } else {
    a
  }
}


`%coalesce%` <- function(a, b) {
  if (is.null(a) || length(a) == 0 || (is.atomic(a) && length(a) == 1 && (is.na(a) || !nzchar(a)))) b else a
}

# ---------------------------------------------------------------------
# PARAMÈTRES DE STYLE POUR UN RENDU PROFESSIONNEL
# ---------------------------------------------------------------------
WFP_BLUE    <- "#0A66C2"   
GREY_TXT    <- "#777777"
WFP_DARK_BLUE <- "#001F3F" 
GREY_BG     <- "#F2F2F2"   
WHITE_TXT   <- "#FFFFFF"   
RED_TXT     <- "#C00000" 
FONT_FAMILY <- "Cambria (Body)"     
LINE_SP     <- 1.0        
INDENT_Q    <- 0.3        
INDENT_C    <- 0.5        

FS_TITLE    <- 14
FS_SUBTITLE <- 12
FS_BLOCK    <- 12
FS_META     <- 10    
FS_HINT     <- 9   
FS_RELV     <- 9   
FS_MISC     <- 10     
FS_Q_BLUE   <- 11  
FS_Q_GREY   <- 9   

EXCLUDE_TYPES <- c("calculate","start","end","today","deviceid","subscriberid","simserial","phonenumber","username","instanceid","end_group", "end group")
EXCLUDE_NAMES <- c("start","end","today","deviceid","username","instanceID","instanceid")

LOGO_PATH  <- "cp_logo.png"

# ---------------------------------------------------------------------
# Fonctions utilitaires 
# ---------------------------------------------------------------------

generate_output_path <- function(base_path) {
  if (!file.exists(base_path)) return(base_path)
  path_dir <- dirname(base_path)
  path_ext <- tools::file_ext(base_path)
  path_name <- tools::file_path_sans_ext(base_path)
  i <- 1
  new_path <- base_path
  while (file.exists(new_path)) {
    new_name <- paste0(path_name, "_", i, ".", path_ext)
    new_path <- file.path(path_dir, basename(new_name)) 
    i <- i + 1
  }
  return(new_path)
}

# Fonctions de détection de colonne mises à jour pour la langue et robustesse
detect_label_col <- function(df, target_language = NULL) {
  cols <- names(df)
  cols_clean <- tolower(str_replace_all(iconv(cols, to = "ASCII//TRANSLIT"), "\\s", ""))
  
  if (!is.null(target_language) && target_language != "Auto") {
    specific_label <- cols[grepl(glue("label(::|:){target_language}"), cols_clean, ignore.case = TRUE)]
    if (length(specific_label) > 0) return(specific_label)
  }
  
  fr <- cols[grepl("label(::|:)french", cols_clean, ignore.case = TRUE)]
  if (length(fr) > 0) return(fr)
  ll <- cols[grepl("label(::|:)", cols_clean)]
  if (length(ll) > 0) return(ll)
  if ("label" %in% cols_clean) return(cols[cols_clean == "label"])
  NA_character_
}

detect_hint_col <- function(df, target_language = NULL) {
  cols <- names(df)
  cols_clean <- tolower(str_replace_all(iconv(cols, to = "ASCII//TRANSLIT"), "\\s", ""))
  
  if (!is.null(target_language) && target_language != "Auto") {
    specific_hint <- cols[grepl(glue("hint(::|:){target_language}"), cols_clean, ignore.case = TRUE)]
    if (length(specific_hint) > 0) return(specific_hint)
  }
  
  fr <- cols[grepl("hint(::|:)french", cols_clean, ignore.case = TRUE)]
  if (length(fr) > 0) return(fr)
  if ("hint" %in% cols_clean) return(cols[cols_clean == "hint"])
  ll <- cols[grepl("hint(::|:)", cols_clean)]
  if (length(ll) > 0) return(ll)
  if ("hint" %in% cols_clean) return(cols[cols_clean == "hint"])
  NA_character_
}

# Fonction translate_relevant complète
translate_relevant <- function(expr, labels, choices, df_survey) {
  if (is.null(expr) || is.na(expr) || !nzchar(trimws(expr))) return(NA_character_)
  txt <- expr
  
  choices <- choices %>% mutate(
    name = tolower(str_trim(as.character(name))),
    list_name = tolower(str_trim(as.character(list_name)))
  )
  df_survey_clean <- df_survey %>% mutate(
    name = tolower(str_trim(as.character(name))),
    type = tolower(as.character(type))
  )
  
  get_choice_label <- function(code, list_name){
    code_clean <- tolower(str_trim(as.character(code)))
    list_name_clean <- tolower(str_trim(as.character(list_name)))
    r <- choices %>% filter(name == code_clean & list_name == list_name_clean)
    if (nrow(r) > 0) {
      label_col_name_choices <- detect_label_col(choices)
      if (!is.na(label_col_name_choices)) {
        lab_val <- as.character(r[[label_col_name_choices]]) 
        lab <- lab_val %||% as.character(code)
      } else {
        lab <- as.character(r$name) %||% as.character(code)
      }
      return(str_replace_all(lab, "<.*?>", ""))
    } else { return(as.character(code)) }
  }
  
  get_var_label <- function(v){ 
    val <- tryCatch(labels[[v]], error = function(e) NULL)
    if(is.null(val) || is.na(val) || !nzchar(val)) {
      clean_val <- as.character(v)
    } else {
      clean_val <- str_replace_all(as.character(val), "<.*?>", "")
    }
    return(clean_val) 
  }
  
  get_listname_from_varname <- function(var_name, survey_df) {
    # Dépendances légères
    if (!requireNamespace("dplyr", quietly = TRUE) ||
        !requireNamespace("stringr", quietly = TRUE)) {
      stop("Veuillez installer dplyr et stringr.")
    }
    
    stopifnot(is.data.frame(survey_df))
    needed <- c("name", "type")
    missing <- setdiff(needed, names(survey_df))
    if (length(missing)) {
      stop("Colonnes manquantes dans survey_df: ", paste(missing, collapse = ", "))
    }
    
    var_name_clean <- tolower(stringr::str_trim(as.character(var_name)[1]))
    row <- survey_df %>% filter(name == var_name_clean) %>% head(1)
    if (nrow(row) > 0) {
      type_val <- tolower(as.character(row$type))
      list_name <- stringr::str_replace(type_val, "^select_(one|multiple)\\s+", "")
      final_list_name <- ifelse(
        test = list_name == type_val,
        yes = as.character(row$list_name %||% NA_character_),
        no = list_name
      )
      return(str_trim(final_list_name))
    } else { 
      return(NA_character_) 
    }
  }
  
  txt <- stringr::str_replace_all(txt, "selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+'\\s*\\)", function(x){ 
    m <- stringr::str_match(x, "selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)")
    vars <- m[,2]; codes <- m[,3]
    vapply(seq_along(vars), function(i){
      current_listname <- get_listname_from_varname(vars[i], df_survey_clean)
      choice_label <- get_choice_label(codes[i], current_listname)
      sprintf("`%s` a l'option «%s» cochée", get_var_label(vars[i]), choice_label)
    }, character(1)) 
  })
  
  txt <- stringr::str_replace_all(txt, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{[^}]+\\}\\s*,\\s*'[^']+?'\\s*\\)\\s*\\)", function(x){ 
    m <- stringr::str_match(x, "not\\s*\\(\\s*selected\\s*\\(\\s*\\$\\{([^}]+)\\}\\s*,\\s*'([^']+)'\\s*\\)\\s*\\)")
    vars <- m[,2]; codes <- m[,3]
    vapply(seq_along(vars), function(i){
      current_listname <- get_listname_from_varname(vars[i], df_survey_clean)
      choice_label <- get_choice_label(codes[i], current_listname)
      sprintf("`%s` N'A PAS l'option «%s» cochée", get_var_label(vars[i]), choice_label)
    }, character(1)) 
  })
  
  comparison_regex <- "\\$\\{[^}]+\\}\\s*[=><!]+\\s*('|\")?[^'\"]+('|\")?"
  txt <- stringr::str_replace_all(txt, comparison_regex, function(comp_expr) {
    var_match <- str_match(comp_expr, "\\$\\{([^}]+)\\}")
    var_name <- var_match[, 2]
    code_match <- str_match(comp_expr, "('|\")?([^'\"]+)('|\")?$")
    code_val <- code_match[, 3]
    current_listname <- get_listname_from_varname(var_name, df_survey_clean)
    choice_label <- get_choice_label(code_val, current_listname)
    reconstructed_expr <- sprintf("`%s` a la valeur «%s»", get_var_label(var_name), choice_label)
    reconstructed_expr <- str_replace(reconstructed_expr, "\\s*=\\s*", " est égal à ")
    reconstructed_expr <- str_replace(reconstructed_expr, "\\s*!=\\s*", " est différent de ")
    return(reconstructed_expr)
  })
  
  txt <- stringr::str_replace_all(txt, "\\$\\{[^}]+\\}", function(x){ 
    m <- stringr::str_match(x, "\\$\\{([^}]+)\\}")
    vars <- m[,2]
    vapply(vars, get_var_label, character(1)) 
  })
  
  txt <- str_replace_all(txt, "\\bandand\\b", " et "); 
  txt <- str_replace_all(txt, "\\bor\\b",  " ou "); 
  txt <- str_replace_all(txt, "\\bnot\\b", " non "); 
  txt <- str_replace_all(txt, "\\s*!=\\s*", " est différent de ");
  txt <- str_replace_all(txt, "\\s*=\\s*", " est égal à ");
  txt <- str_replace_all(txt, "\\s*>\\s*", " est supérieur à ");
  txt <- str_replace_all(txt, "\\s*>=\\s*", " est supérieur ou égal à ");
  txt <- str_replace_all(txt, "\\s*<\\s*", " est inférieur à ");
  txt <- str_replace_all(txt, "\\s*<=\\s*", " est inférieur ou égal à ");
  
  txt <- str_replace_all(txt, "count-selected\\s*\\(\\s*[^\\)]+\\)\\s*>=\\s*1", function(x){ 
    m <- str_match(x, "count-selected\\s*\\(\\s*(.+?)\\s*\\)\\s*>=\\s*1")
    v <- m[,2]
    vapply(v, function(z) sprintf("Au moins une option est cochée pour %s", z), character(1)) 
  })
  
  txt <- str_replace_all(txt, "\\s*\\(\\s*", "\\(") %>% str_replace_all("\\s*\\)\\s*", "\\)") %>% str_replace_all("\\s+", " ") %>% str_trim()
  return(txt)
}

# ---------------------------------------------------------------------
# Styles texte & paragraphe / Blocs visuels 
# ---------------------------------------------------------------------
fp_txt <- function(color="black", size=11, bold=FALSE, italic=FALSE, underline=FALSE) { fp_text(color = color, font.size = size, bold = bold, italic = italic, underline = underline, font.family = FONT_FAMILY) }
fp_q_blue   <- fp_txt(color = WFP_DARK_BLUE, size = FS_Q_BLUE, bold = FALSE) 
fp_q_grey   <- fp_txt(color = GREY_TXT, size = FS_Q_GREY)             
fp_sec_title <- fp_txt(color = WFP_BLUE, size = FS_TITLE, bold = TRUE, underline = TRUE)
fp_sub_title <- fp_txt(color = WFP_BLUE, size = FS_SUBTITLE, bold = TRUE, underline = TRUE)
fp_block     <- fp_txt(color = WFP_BLUE, size = FS_BLOCK,  bold = TRUE) 
fp_meta      <- fp_txt(color = GREY_TXT, size = FS_META)
fp_hint      <- fp_txt(color = GREY_TXT, size = FS_HINT, italic = TRUE)
fp_relevant  <- fp_txt(color = RED_TXT,  size = FS_RELV)
fp_missing_list <- fp_txt(color = GREY_TXT, size = FS_MISC, italic = TRUE)
p_default <- fp_par(text.align = "left", line_spacing = LINE_SP)  
p_q_indent_fixed <- fp_par(text.align = "left", line_spacing = LINE_SP, padding.left = INDENT_Q * 72)
p_c_indent_fixed <- fp_par(text.align = "left", line_spacing = LINE_SP, padding.left = INDENT_C * 72)

add_hrule <- function(doc, width = 1){ doc <- body_add_par(doc, ""); return(doc) }
add_band <- function(doc, text, txt_fp){ doc <- body_add_fpar(doc, fpar(ftext(text, txt_fp), fp_p = fp_par(padding.top = 4, padding.bottom = 4, padding.left = INDENT_Q*72))); return(doc) }

add_choice_lines <- function(doc, choices_map, list_name_to_filter, symbol = "○") {
  if (is.na(list_name_to_filter) || list_name_to_filter == "") return(doc)
  list_name_to_filter_clean <- str_replace_all(list_name_to_filter, "\\s", "")
  
  df <- choices_map[[list_name_to_filter_clean]]
  
  if (is.null(df) || nrow(df) == 0) {
    doc <- body_add_fpar(doc, fpar(ftext(glue("Commentaire: La liste de choix '{list_name_to_filter}' est introuvable ou vide dans l'onglet 'choices'."), fp_missing_list), fp_p = p_c_indent_fixed)) 
    return(doc)
  }
  
  df <- df %>% mutate(
    label_final = coalesce(label_col, as.character(name)),
    txt_final = sprintf("%s %s (%s)", symbol, label_final, name)
  )
  
  purrr::walk(df$txt_final, function(text_label) { 
    doc <<- body_add_fpar(doc, fpar(ftext(text_label, fp_txt(size = FS_MISC)), fp_p = p_c_indent_fixed)) 
  }); 
  return(doc)
}

add_placeholder_box <- function(doc, txt = "Réponse : [insérer votre réponse ici]"){
  doc <- body_add_fpar(doc, fpar(ftext(txt, fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_c_indent_fixed))
  doc <- body_add_par(doc, ""); return(doc)
}

# ---------------------------------------------------------------------
# Rendu d’une question
# ---------------------------------------------------------------------
render_question <- function(doc, row, number, label_col_name, hint_col_name, choices_map, lab_map, full_choices_sheet,full_survey_sheet){
  if (is.null(row) || nrow(row) == 0) return(doc)
  q_type <- tolower(row$type %||% "")
  q_name <- row$name %||% ""
  
  clean_html <- function(text) {
    text <- str_replace_all(text, "<[^>]+>", "")
    text <- str_replace_all(text, "&nbsp;", "")
    text <- str_replace_all(text, "\\s+", " ")
    return(str_trim(text))
  }
  
  q_lab  <- clean_html(row[[label_col_name]] %||% q_name)
  
  if (is.na(q_name) || q_name %in% EXCLUDE_NAMES) return(doc)
  
  if (q_type == "note") {
    ftext_blue <- ftext(sprintf("Note : %s", q_lab), fp_q_blue)
  } else {
    if (is.null(number) || is.na(number) || !nzchar(as.character(number))) return(doc)
    ftext_blue <- ftext(sprintf("%s. %s", number, q_lab), fp_q_blue)
  }
  ftext_grey <- ftext(sprintf(" (%s – %s)", q_name, q_type), fp_q_grey)
  doc <- body_add_fpar(doc, fpar(ftext_blue, ftext_grey, fp_p = p_q_indent_fixed)) 
  
  rel <- row$relevant %||% NA_character_
  h <- NA_character_
  if (!is.na(hint_col_name) && hint_col_name %in% names(row)) { 
    h <- row[[hint_col_name]] %||% NA_character_ 
    h <- clean_html(h)
  }
  
  if (!is.na(rel) && nzchar(rel)) { tr <- translate_relevant(rel, lab_map, full_choices_sheet, full_survey_sheet); doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) }
  if (!is.na(h) && nzchar(h)) { doc <- body_add_fpar(doc, fpar(ftext(h, fp_hint), fp_p = p_q_indent_fixed)) }
  
  if (str_starts(q_type, "select_one")) { 
    doc <- body_add_fpar(doc, fpar(ftext("Choisir la réponse parmi la liste ci-bas", fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_q_indent_fixed))
    ln <- str_trim(sub("^select_one\\s+", "", q_type));
    doc <- add_choice_lines(doc, choices_map, ln, symbol = "○");
  } 
  else if (str_starts(q_type, "select_multiple")) { 
    doc <- body_add_fpar(doc, fpar(ftext("Choisir les réponses pertinentes parmi la liste ci-bas", fp_txt(size = FS_MISC, italic = TRUE)), fp_p = p_q_indent_fixed))
    ln <- str_trim(sub("^select_multiple\\s+", "", q_type)); 
    doc <- add_choice_lines(doc, choices_map, ln, symbol = "☐"); 
  } 
  else if (str_detect(q_type, "^note")) { } 
  else { 
    placeholder <- "Réponse : [insérer votre réponse ici]"
    if (str_detect(q_type, "integer")) { placeholder <- "Réponse : [insérer un entier]" } 
    else if (str_detect(q_type, "decimal")) { placeholder <- "Réponse : [insérer un décimal]" }
    else if (str_detect(q_type, "date")) { placeholder <- "Réponse : [insérer une date]" }
    else if (str_detect(q_type, "geopoint")) { placeholder <- "Réponse : [capturer les coordonnées GPS]" }
    else if (str_detect(q_type, "image")) { placeholder <- "Réponse : [prendre une photo] ou inserer une image" }
    else if (str_detect(q_type, "photo")) { placeholder <- "Réponse : [prendre une photo]" }
    else if (str_detect(q_type, "rank")) { placeholder <- "Réponse : [classer la liste de choix par ordre de préférence]" }
    else if (str_detect(q_type, "file")) { placeholder <- "Réponse : [joindre un fichier]" }
    else if (str_detect(q_type, "time")) { placeholder <- "Réponse : [saisir une heure]" }
    else if (str_detect(q_type, "video")) { placeholder <- "Réponse : [enregistrer ou sélectionner une vidéo]" }
    else if (str_detect(q_type, "audio")) { placeholder <- "Réponse : [enregistrer ou sélectionner un audio]" }
    else if (str_detect(q_type, "barcode")) { placeholder <- "Réponse : [scanner un code-barres]" }
    else if (str_detect(q_type, "range")) { placeholder <- "Réponse : [ choisir une valeur dans la plage ci-bas]" }
    else if (str_detect(q_type, "geoshape")) { placeholder <- "Réponse : [enregistrer un polygone (premier et dernier point identiques)]" }
    else if (str_detect(q_type, "geotrace")) { placeholder <- "Réponse : [enregistrer une ligne de points]" }
    else if (str_detect(q_type, "acknowledge")) { placeholder <- "Réponse : [confirmer votre déclaration]" }
    
    
    doc <- add_placeholder_box(doc, placeholder)
  }
  
  if (!str_detect(q_type, "^note|select_")) {
    doc <- body_add_fpar(doc, fpar(ftext("__________________________________________________________", fp_txt(size=FS_Q_BLUE)), fp_p = p_c_indent_fixed))
  }
  
  doc <- body_add_par(doc, "")
  return(doc)
}

# ---------------------------------------------------------------------
# Function xlsform_to_wordRev pour Shiny
# ---------------------------------------------------------------------
xlsform_to_wordRev <- function(xlsx, output_path, selected_language = NULL) {
  
  # Lecture des données
  survey   <- read_excel(xlsx, sheet = 'survey', col_types = 'text')
  choices  <- read_excel(xlsx, sheet = 'choices', col_types = 'text')
  settings <- tryCatch(read_excel(xlsx, sheet = "settings"), error = function(e) NULL)
  
  # Détection de la colonne label/hint en fonction de la langue sélectionnée par l'utilisateur
  label_col_name <- detect_label_col(survey, target_language = selected_language)
  hint_col_name <- detect_hint_col(survey, target_language = selected_language)
  
  if (is.na(label_col_name)) stop("Colonne label introuvable pour la langue sélectionnée.")
  label_col_sym <- sym(label_col_name) 
  
  # Normalisation et préparation des données
  names(choices) <- tolower(iconv(names(choices), to = "ASCII//TRANSLIT"))
  doc_title <- if (!is.null(settings) && "form_title" %in% names(settings)) settings$form_title %||% "XLSForm – Rendu Word" else "XLSForm – Rendu Word"
  survey <- survey %>% filter(!duplicated(name))
  lab_map <- survey %>% select(name, !!label_col_sym) %>% mutate(name = as.character(name)) %>% tibble::deframe()
  
  # Initialisation du document Word (utilisation d'un document vierge pour Shiny)
  doc <- read_docx() 
  doc <- body_set_default_section(doc, prop_section(page_size = page_size(orient = "portrait", width = 8.5, height = 11), page_margins = page_mar(top = 1, bottom = 1, left = 1.0, right = 1.0, header = 0.5, footer = 0.5)))
  
  # Ajout du titre (le logo est optionnel et dépend du chemin local)
  fp_title_main <- fp_txt(color = RED_TXT, size = 16, bold = TRUE)
  p_title_main <- fp_par(text.align = "center", line_spacing = LINE_SP)
  

  if (exists("LOGO_PATH") && !is.null(LOGO_PATH) && file.exists(LOGO_PATH)) { 
    doc <- body_add_par(doc, "", style = "Normal")
    doc <- body_add_fpar(doc, fpar(external_img(src = LOGO_PATH, height = 0.6, width = 0.6, unit = "in"), ftext("  "), ftext(doc_title, prop = fp_title_main), fp_p = p_title_main)) 
  } else { 
    doc <- body_add_fpar(doc, fpar(ftext(doc_title, fp_title_main), fp_p = p_title_main)) 
  }
  doc <- add_hrule(doc)
  
  # Préparation de la map des choix
  choices <- choices %>% 
    mutate_all(as.character) %>% 
    mutate_all(~ifelse(is.na(.), NA_character_, .)) %>%
    mutate(list_name = as.character(str_trim(str_replace_all(tolower(list_name), "[[:space:]]+", ""))))
  
  label_col_choices <- detect_label_col(choices, target_language = selected_language)
  if (is.na(label_col_choices)) stop("Colonne label choix introuvable pour la langue sélectionnée.")
  choices_map <- split(choices, choices$list_name)
  choices_map <- lapply(choices_map, function(df) { df %>% mutate(label_col = .data[[label_col_choices]]) })
  
  # Boucle principale (Logique de parcours du formulaire)
  sec_id <- 0L; sub_id <- 0L; q_id <- 0L; names_deja_traites <- character(0) 
  
  for (i in seq_len(nrow(survey))) {
    r <- survey[i, , drop = FALSE]
    t_raw <- as.character(r$type %||% "")
    t <- tolower(t_raw)
    qname <- as.character(r$name)
    if (!is.na(qname) && nzchar(qname)) { if (qname %in% names_deja_traites) { next } else { names_deja_traites <- c(names_deja_traites, qname) } }
    
    if (t %in% EXCLUDE_TYPES) next
    if (!is.null(r$name) && any(tolower(r$name) %in% tolower(EXCLUDE_NAMES))) next
    
    current_number <- NULL
    is_group <- str_starts(t, "begin_group") && !str_starts(t, "begin_repeat")
    is_repeat <- str_starts(t, "begin_repeat") || str_starts(t, "begin repeat")
    is_repeat_end <- str_starts(t, "end_repeat") || str_starts(t, "end repeat")
    is_end <- str_starts(t, "end_group") || str_starts(t, "end group") || is_repeat_end
    
    if (is_group || is_repeat) {
      lbl <- r[[label_col_name]] %||% r$name %||% ""
      if (is_group) {
        prev_type <- if (i > 1) tolower(as.character(survey$type[i - 1] %||% "")) else ""
        prev_is_group <- str_starts(prev_type, "begin_group") || str_starts(prev_type, "begin repeat")
        if ( !prev_is_group) { sec_id <- sec_id + 1L; sub_id <- 0L; q_id <- 0L; doc <- add_band(doc, glue("Section {sec_id} : {lbl}"), txt_fp = fp_sec_title) } else { sub_id <- sub_id + 1L; q_id <- 0L; doc <- add_band(doc, glue("Sous-section {sec_id}.{sub_id} : {lbl}"), txt_fp = fp_sub_title) }
      } else if (is_repeat) { doc <- body_add_fpar(doc, fpar(ftext(glue("🔁 Bloc : {lbl}"), fp_block), fp_p = p_default)) }
      rel <- r$relevant %||% NA_character_
      if (!is.na(rel) && nzchar(rel)) { 
        tr <- translate_relevant(rel, lab_map, choices, survey); 
        doc <- body_add_fpar(doc, fpar(ftext("Afficher si : ", fp_relevant), ftext(tr, fp_relevant), fp_p = p_q_indent_fixed)) 
      }
      doc <- add_hrule(doc) ; next
    }
    if (is_end) {
      doc <- body_add_fpar(doc, fpar(ftext("--- Fin du bloc ---", fp_hint), fp_p = p_default))
      doc <- add_hrule(doc)
      next 
    }
    
    if(t != "note") q_id <- q_id + 1L
    
    current_number <- NULL
    if(t != "note") {
      current_number <- glue("{q_id}")
      if (sub_id > 0) { current_number <- glue("{sec_id}.{sub_id}.{q_id}") } else if (sec_id > 0) { current_number <- glue("{sec_id}.{q_id}") }
    }
    
    # Appel à render_question
    doc <- render_question(doc, r, current_number, label_col_name, hint_col_name, choices_map, full_choices_sheet = choices, lab_map = lab_map, full_survey_sheet = survey)
  }
  
  # Écriture du document final vers le chemin temporaire de Shiny
  print(doc, target = output_path)
  return(output_path)
}

# ---------------------------------------------------------------------
# Validation / Audit — catégories d’erreurs + cohérence relevant + structure
# ---------------------------------------------------------------------

validate_xlsform <- function(filepath, ignore_empty = TRUE) {
  errors <- list()
  
  # Fonction pour ajouter une erreur
  add_error <- function(sheet, line, type, description, suggestion = NULL) {
    errors <<- append(errors, list(data.frame(
      Feuille = sheet, Ligne = as.character(line), Type = type, Description = description,
      Suggestion = suggestion, stringsAsFactors = FALSE
    )))
  }
  
  # Utilitaires
  na_to_empty <- function(x) ifelse(is.na(x), "", x)
  is_empty_row <- function(row, cols = c("type", "name", "label")) {
    all(sapply(cols, function(c) is.null(row[[c]]) || is.na(row[[c]]) || !nzchar(row[[c]])))
  }
  
  # Lecture sécurisée des onglets
  survey <- tryCatch(read_excel(filepath, sheet = "survey", col_types = "text"), error = function(e) {
    add_error("Général", "N/A", "Erreur Critique", glue("Lecture Excel impossible : {e$message}"),
              "Vérifier le fichier et les onglets requis.")
    return(NULL)
  })
  if (is.null(survey)) return(bind_rows(errors))
  
  choices <- tryCatch(read_excel(filepath, sheet = "choices", col_types = "text"), error = function(e) NULL)
  settings <- tryCatch(read_excel(filepath, sheet = "settings", col_types = "text"), error = function(e) NULL)
  
  names(survey) <- tolower(names(survey))
  
  # Vérification colonnes obligatoires
  if (!all(c("type","name") %in% names(survey))) {
    miss <- setdiff(c("type","name"), names(survey))
    add_error("survey", "En-tête", "Colonnes Manquantes", glue("Colonnes obligatoires manquantes : {paste(miss, collapse=', ')}."),
              "Ajouter les colonnes 'type' et 'name'.")
  }
  
  # Vérification label
  if (!any(grepl("label", names(survey), ignore.case=TRUE))) {
    add_error("survey", "En-tête", "ODK Compliance", "Colonne label manquante.", "Ajouter une colonne label ou label::lang.")
  }
  
  # Vérification settings
  if (is.null(settings) || !all(c("form_id","form_title") %in% names(settings))) {
    add_error("settings", "En-tête", "ODK Compliance", "Onglet settings incomplet.", "Ajouter form_id et form_title pour compatibilité cloud.")
  }
  
  # Vérification types valides et casse
  valid_types <- c("integer","decimal","text","select_one","select_multiple","select one","select multiple","note","calculate","date","datetime","geopoint","image","audio","video","begin_group","begin_repeat","end_group","end_repeat","begin group","begin repeat","end group","end repeat","start","end","deviceid","today",
                   'barcode'
                   )
  
  for (i in seq_len(nrow(survey))) {
    if (ignore_empty && is_empty_row(survey[i,])) next
    t_raw <- na_to_empty(survey$type[i])
    t <- tolower(t_raw)
    if (nzchar(t) && !any(str_detect(t, paste(valid_types, collapse="|")))) {
      add_error("survey", i+1, "ODK Compliance", glue("Type '{t_raw}' non conforme à la spec ODK/XForms."),
                "Utiliser un type valide (ex: integer, text, select_one, begin_group, end_group).")
    }
    if (str_detect(t_raw, "(?i)begin_ |end_ ") && t_raw != t) {
      add_error("survey", i+1, "ODK Compliance", glue("Type '{t_raw}' doit être en minuscules."), glue("Corrigez en '{t}'."))
    }
  }
  
  # Vérification appearance
  if ("appearance" %in% names(survey)) {
    for (i in seq_len(nrow(survey))) {
      if (ignore_empty && is_empty_row(survey[i,])) next
      app <- na_to_empty(survey$appearance[i])
      if (nzchar(app) && !grepl("^(numbers|multiline|url|ex:|thousands-sep|bearing|vertical|no-ticks
|picker|rating|new|new-front|draw|annotate|signature|no-calendar|month-year
|year|ethiopian|coptic|islamic|bikram-sambat|myanmar|persian|placement-map|maps
|hide-input|minimal|search|quick|columns-pack|columns|columns-n|no-buttons|image-map
|likert|map|field-list|label|list-nolabel|list|table-list|hidden-answer|printer|masked|counter
)$", app)) {
        add_error("survey", i+1, "ODK Compliance", glue("Appearance '{app}' non reconnu."),
                  "Utiliser une valeur valide (minimal, compact, etc.).")
      }
    }
  }
  
  # Vérification fonctions interdites
  if ("calculation" %in% names(survey)) {
    for (i in seq_len(nrow(survey))) {
      if (ignore_empty && is_empty_row(survey[i,])) next
      calc <- na_to_empty(survey$calculation[i])
      if (nzchar(calc) && grepl("/ ", calc)) {
        add_error("survey", i+1, "ODK Compliance", "Fonction non supportée détectée.",
                  "Éviter '/', utiliser plutôt div() pour les divisions.")
      }
    }
  }
  
  # Missing calculation si type = calculate
  if ("calculation" %in% names(survey)) {
    for (i in seq_len(nrow(survey))) {
      if (ignore_empty && is_empty_row(survey[i,])) next
      if (tolower(na_to_empty(survey$type[i])) == "calculate" && !nzchar(na_to_empty(survey$calculation[i]))) {
        add_error("survey", i+1, "Calculation Manquante", "Type 'calculate' sans expression.",
                  "Ajouter une formule dans la colonne 'calculation'.")
      }
    }
  }
  
  # Variables non identifiées
  check_unknown_vars <- function(expr, all_names) {
    vars <- str_match_all(expr, "\\$\\{([^\\}]+)\\}")[[1]]
    if (nrow(vars) > 0) {
      unknowns <- setdiff(tolower(vars[,2]), all_names)
      return(unknowns)
    }
    return(character(0))
  }
  all_names <- tolower(na_to_empty(survey$name))
  for (col in c("calculation","relevant","constraint","choice_filter")) {
    if (col %in% names(survey)) {
      for (i in seq_len(nrow(survey))) {
        if (ignore_empty && is_empty_row(survey[i,])) next
        expr <- na_to_empty(survey[[col]][i])
        if (nzchar(expr)) {
          unknowns <- check_unknown_vars(expr, all_names)
          if (length(unknowns) > 0) {
            add_error("survey", i+1, glue("Références Invalides ({col})"),
                      glue("{col} référence des variables inconnues : {paste(unknowns, collapse=', ')}."),
                      "Corriger les références.")
          }
        }
      }
    }
  }
  
  # jr:choice-name invalide + erreur si utilisé avec select_multiple
  
  select_multiple_vars <- tolower(survey$name[grepl("^select_multiple", tolower(survey$type))])
  if ("calculation" %in% names(survey)) {
    for (i in seq_len(nrow(survey))) {
      if (ignore_empty && is_empty_row(survey[i,])) next
      calc <- na_to_empty(survey$calculation[i])
      if (grepl("jr:choice-name", calc)) {
        vars <- str_match_all(calc, "\\$\\{([^\\}]+)\\}")[[1]][,2]
        
        # Vérifier syntaxe : accepter selected-at()
        if (!grepl("jr:choice-name\\(selected-at\\(\\$\\{[^\\}]+\\},", calc) &&
            !grepl("jr:choice-name\\(\\$\\{[^\\}]+\\},\\s*'[^']+'\\)", calc)) {
          add_error("survey", i+1, "Syntaxe jr:choice-name Invalide", glue("Expression incorrecte : '{calc}'."),
                    "Utiliser jr:choice-name(${var}, 'list_name') ou selected-at() pour select_multiple.")
        }
        
        # Erreur si jr:choice-name avec select_multiple sans selected-at()
        if (any(tolower(vars) %in% select_multiple_vars) && !grepl("selected-at", calc)) {
          add_error("survey", i+1, "Erreur jr:choice-name",
                    glue("jr:choice-name() utilisé avec une question select_multiple : {paste(vars, collapse=', ')}."),
                    "Utiliser selected-at() pour extraire chaque choix avant jr:choice-name(), ou join() avec repeat.")
        }
      }
    }
  }
  
  
  
  # Duplication colonnes
  for (sheet_name in c("survey","choices","settings")) {
    df <- tryCatch(read_excel(filepath, sheet = sheet_name, col_types = "text"), error=function(e) NULL)
    if (!is.null(df)) {
      dup_cols <- names(df)[duplicated(names(df))]
      if (length(dup_cols) > 0) {
        add_error(sheet_name, "En-tête", "Colonnes Dupliquées", glue("Colonnes dupliquées : {paste(dup_cols, collapse=', ')}."),
                  "Supprimer les doublons.")
      }
    }
  }
  
  # Duplication noms survey
  names_for_dup <- survey$name
  names_for_dup[is.na(names_for_dup) | !nzchar(names_for_dup)] <- NA
  dups <- names_for_dup[duplicated(names_for_dup) & !is.na(names_for_dup)]
  if (length(dups) > 0) {
    for (dup_name in unique(dups)) {
      lines <- which(names_for_dup == dup_name) + 1
      add_error("survey", paste(lines, collapse=", "), "Duplication",
                glue("Le nom de variable '{dup_name}' est utilisé plusieurs fois."),
                "Les noms doivent être uniques.")
    }
  }
  # Choix manquants pour select_one/multiple
  sel <- survey %>% filter(stringr::str_detect(tolower(type %coalesce% ""), "^select_one |^select_multiple "))
  if (nrow(sel) > 0) {
    if (is.null(choices) || nrow(choices) == 0) {
      add_error("choices", "N/A", "Choices Non Disponibles", "L'onglet 'choices' est manquant ou vide.", "Ajouter l'onglet 'choices'.")
    } else {
      names(choices) <- tolower(names(choices))
      if (!"list_name" %in% names(choices)) {
        add_error("choices", "En-tête", "Colonnes Manquantes", "Colonne 'list_name' manquante.", "Ajouter 'list_name' dans 'choices'.")
      } else {
        required_lists <- stringr::str_trim(gsub("select_(one|multiple)\\s+", "", sel$type))
        available <- unique(choices$list_name)
        for (ln in required_lists) {
          if (!ln %in% available) {
            lines <- which(stringr::str_detect(survey$type, stringr::fixed(ln)))
            add_error("survey", paste(lines+1, collapse=", "), "Choices Non Disponibles", glue("La liste de choix '{ln}' est introuvable dans 'choices'."), "Vérifier l’orthographe ou définir la liste.")
          }
        }
      }
    }
  }
  
  # Duplication choix par list_name
  
  
  if (!is.null(choices) && all(c("list_name","name") %in% names(choices))) {
    
    choices_row_numbers <- seq(2, nrow(choices) + 1) 
    
    choices_filtered <- choices %>%
      filter(!is.na(list_name) & nzchar(list_name) & !is.na(name) & nzchar(name)) %>%
      mutate(list_name = tolower(list_name), name = tolower(name))
    
    dup_rows <- choices_filtered %>% group_by(list_name, name) %>% filter(n() > 1)
    
    if (nrow(dup_rows) > 0) {
      for (i in seq_len(nrow(dup_rows))) {
        ln <- dup_rows$list_name[i]
        nm <- dup_rows$name[i]
        
        lines <- choices_row_numbers[which(tolower(choices$list_name) == ln & tolower(choices$name) == nm)]
        
        add_error("choices", paste(lines, collapse=", "), "Duplication",
                  glue("Le choix '{nm}' est dupliqué dans la liste '{ln}'."),
                  "Les noms des choix doivent être uniques par liste.")
      }
    }
    
  }
  
  
  # Vérification des noms dans choices
  if (!is.null(choices) && all(c("list_name","name") %in% names(choices))) {
    choices_row_numbers <- seq(2, nrow(choices) + 1)
    
    for (i in seq_len(nrow(choices))) {
      row <- choices[i, ]
      
      if (all(is.na(row$list_name) | !nzchar(row$list_name),
              is.na(row$name) | !nzchar(row$name),
              !("label" %in% names(row)) || is.na(row$label) || !nzchar(row$label))) {
        next  # Ignorer la ligne vide
      }
      
      nm <- na_to_empty(row$name)
      
      
      # Si name est vide
      
      if (!nzchar(nm)) {
        add_error("choices", choices_row_numbers[i], "Nom Invalide",
                  "Le champ 'name' est vide. Chaque choix doit avoir un nom.",
                  "Ajouter une valeur valide dans la colonne 'name'. Voir : https://xlsform.org/en/#setting-up-your-worksheets")
      }
      
      # Si name contient des caractères invalides
      else if (!grepl("^[a-zA-Z0-9_]+$", nm)) {
        add_error("choices", choices_row_numbers[i], "Nom Invalide",
                  glue("Le nom '{nm}' contient des caractères non autorisés."),
                  "Utiliser uniquement lettres, chiffres et underscore.")
      }
    }
  }

  
  
  # Structure begin/end group/repeat
  stack <- list()
  for (i in seq_len(nrow(survey))) {
    if (ignore_empty && is_empty_row(survey[i,])) next
    t <- tolower(na_to_empty(survey$type[i]))
    nm <- na_to_empty(survey$name[i])
    
    if (str_detect(t, "^begin[_ ]group") || str_detect(t, "^begin[_ ]repeat")) {
      stack <- append(stack, list(list(type=t, name=nm, line=i+1)))
      next
    }
    
    if (str_detect(t, "^end[_ ]group") || str_detect(t, "^end[_ ]repeat")) {
      if (length(stack) == 0) {
        add_error("survey", i+1, "Structure", glue("Unmatched '{t}'. Aucun bloc ouvert."),
                  "Vérifiez l'ordre des begin/end.")
        next
      }
      
      last <- stack[[length(stack)]]
      stack <- stack[-length(stack)]
      
      if (str_detect(t, "^end[_ ]group") && !str_detect(last$type, "^begin[_ ]group")) {
        add_error("survey", i+1, "Structure", "Unmatched 'end_group' : dernier bloc ouvert est 'begin_repeat'.",
                  "Fermez avec 'end_repeat'.")
        next
      }
      
      if (str_detect(t, "^end[_ ]repeat") && !str_detect(last$type, "^begin[_ ]repeat")) {
        add_error("survey", i+1, "Structure", "Unmatched 'end_repeat' : dernier bloc ouvert est 'begin_group'.",
                  "Fermez avec 'end_group'.")
        next
      }
      
      if (nzchar(nm) && nzchar(last$name) && nm != last$name) {
        add_error("survey", i+1, "Structure", glue("No matching begin_group : fin '{nm}' ≠ début '{last$name}'."),
                  "Harmonisez les noms ou supprimez le name sur end_* si non nécessaire.")
      }
    }
  }
  
  if (length(stack) > 0) {
    for (grp in stack) {
      msg <- glue("Unmatched '{grp$type}' pour '{grp$name}' : aucun 'end_*' trouvé.")
      suggestion <- glue("Le bloc ouvert à la ligne {grp$line} doit être fermé.")
      add_error("survey", grp$line, "Structure", msg, suggestion)
    }
  }
  


  # Retour du rapport
  if (length(errors) == 0) {
    return(data.frame(Feuille="N/A", Ligne="N/A", Type="Info", Description="Aucune erreur détectée.",
                      Suggestion="OK pour conversion.", stringsAsFactors = FALSE))
  } else {
    return(bind_rows(errors))
  }
}


# ---------------------------------------------------------------------
# Récap Sections / Sous-sections + nombre de questions (affiché après analyse)
# ---------------------------------------------------------------------


compute_structure_summary <- function(survey, label_col, exclude_types = NULL) {
  # Helper
  is_empty_value <- function(x) is.na(x) | !nzchar(x)
  
  # Normalisation
  df <- survey %>% mutate(type_clean = tolower(coalesce(type, "")))
  exclude_types <- tolower(exclude_types %||% character(0))
  
  # Piles et registres
  blocks <- list()           # tous les blocs (sections/sous-sections)
  open_idx_stack <- list()   # pile des indices vers 'blocks' actuellement ouverts
  order_counter <- 0L
  
  get_label <- function(row) {
    if (!is.na(label_col) && label_col %in% names(row)) {
      as.character(coalesce(row[[label_col]], "")) %coalesce% as.character(coalesce(row$name, ""))
    } else {
      as.character(coalesce(row$name, ""))
    }
  }
  
  # Index du bloc "Sans Section" si/qd nécessaire
  sans_section_idx <- NA_integer_
  
  for (i in seq_len(nrow(df))) {
    t <- df$type_clean[i]
    nm <- if ("name" %in% names(df)) df$name[i] else NA_character_
    if (is_empty_value(t) && is_empty_value(nm)) next
    
    r <- df[i, , drop = FALSE]
    is_begin_group_safe  <- !is_empty_value(t) && stringr::str_detect(t, "^begin[_ ]group")
    is_begin_repeat_safe <- !is_empty_value(t) && stringr::str_detect(t, "^begin[_ ]repeat")
    is_end_group_safe    <- !is_empty_value(t) && stringr::str_detect(t, "^end[_ ]group")
    is_end_repeat_safe   <- !is_empty_value(t) && stringr::str_detect(t, "^end[_ ]repeat")
    is_end_safe          <- is_end_group_safe || is_end_repeat_safe
    is_note_safe         <- !is_empty_value(t) && stringr::str_detect(t, "^note")
    
    # Ouverture de bloc
    if (is_begin_group_safe || is_begin_repeat_safe) {
      lbl <- get_label(r)
      order_counter <- order_counter + 1L
      # Créer le bloc
      new_block <- list(
        level = length(open_idx_stack) + 1L,
        title = lbl,
        count = 0L,
        order = order_counter,
        type  = if (is_begin_group_safe) "group" else "repeat"
      )
      blocks <- append(blocks, list(new_block))
      # Pousser son indice sur la pile des blocs ouverts
      open_idx_stack <- append(open_idx_stack, list(length(blocks)))
      next
    }
    
    # Fermeture de bloc
    if (is_end_safe) {
      if (length(open_idx_stack) > 0) {
        # Pop : fermer le dernier bloc ouvert
        open_idx_stack <- open_idx_stack[-length(open_idx_stack)]
      } else {
        # end_* sans begin_* : on ignore mais on pourrait logger
        # (tu as déjà une validation structure séparée)
      }
      next
    }
    
    # Comptage des "questions"
    if (!is_note_safe && !(t %in% exclude_types)) {
      if (length(open_idx_stack) == 0) {
        # Question hors section
        if (is.na(sans_section_idx)) {
          order_counter <- order_counter + 1L
          blocks <- append(blocks, list(list(
            level = 1L, title = "Sans Section", count = 0L, order = order_counter, type = "none"
          )))
          sans_section_idx <- length(blocks)
        }
        blocks[[sans_section_idx]]$count <- blocks[[sans_section_idx]]$count + 1L
      } else {
        # Incrémenter le bloc actuellement ouvert (dernier de la pile)
        idx_cur <- open_idx_stack[[length(open_idx_stack)]]
        blocks[[idx_cur]]$count <- blocks[[idx_cur]]$count + 1L
      }
    }
  }
  
  # S'il n'y a aucun bloc
  if (length(blocks) == 0) {
    return(tibble(Level = integer(0), Titre = character(0), Nb_Questions = integer(0)))
  }
  
  # Construire le tibble avec l'ordre conservé
  out <- tibble(
    Ordre        = purrr::map_int(blocks, ~ .x$order),
    Level        = purrr::map_int(blocks, ~ .x$level),
    Titre        = purrr::map_chr(blocks, ~ .x$title),
    Nb_Questions = purrr::map_int(blocks, ~ .x$count),
    Type         = purrr::map_chr(blocks, ~ .x$type)
  ) %>%
    arrange(Ordre)
  
  # Optionnel : si tu veux regrouper par (Level, Titre) et sommer les questions
  # tout en gardant le premier 'Ordre' pour l’affichage :
  out_grouped <- out %>%
    group_by(Level, Titre) %>%
    summarise(
      Nb_Questions = sum(Nb_Questions),
      Ordre = min(Ordre),
      .groups = "drop"
    ) %>%
    arrange(Ordre, Level, Titre) %>%
    select(Level, Titre, Nb_Questions)
  
  return(out_grouped)
}


# ---------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------
ui <- fluidPage(
  theme = shinytheme("flatly"),
  tags$head(tags$style(HTML('
    .card {border:1px solid #ddd; border-radius:8px; padding:15px; margin-bottom:15px;text-align: left !important;}
    .btn {border-radius:4px;}
    table.dataTable th, table.dataTable td {white-space: nowrap;text-align: left !important;vertical-align: middle !important;}
    .dataTables_wrapper .dataTables_scrollBody {overflow-x: auto;}
  '))),
  titlePanel("XLSFormTools v1.1 : Outil de Validation et de Conversion Word"),
  sidebarLayout(
    sidebarPanel(
      fileInput("file1", "1. Choisir le fichier XLSForm (.xlsx)", accept = c(".xlsx")),
      tags$hr(),
      h4("2. Configuration de la langue"),
      actionButton("detect_lang_btn", "Détecter les langues", icon("globe"), class = "btn-secondary"),
      uiOutput("language_select_ui"),
      tags$hr(),
      h4("3. Actions"),
      actionButton("validate_btn", "Analyser les erreurs", icon("check-circle"), class = "btn-primary"),
      downloadButton("download_word", "Générer Word (.docx)", class = "btn-success"),
      downloadButton("download_excel_report", "Rapport Excel (.xlsx)", class = "btn-info"),
      tags$div(
        tags$h4("Étapes :"),
        tags$ol(
          tags$li("Choisir le fichier XLSForm (.xlsx)"),
          tags$li("Sélectionner la langue pour la conversion et l'analyse"),
          tags$li("Analyser et générer le rapport d'erreurs"),
            tags$ul(
              tags$li("Analyser les erreurs dans votre XLSForm"),
              tags$li("Générer le rapport d'erreurs (Excel)")
            )
          ),
          tags$li("Télécharger la version Word (.docx) (ne pas oublier de sélectionner la langue)")
        ),
      
      tags$hr(),
      p("Documentation ODK :"),
      a("docs.getodk.org", href="https://docs.getodk.org", target="_blank")
    ),
    mainPanel(
      tabsetPanel(id = "errorTabs",
                  tabPanel("Résumé Général", value = "summary",
                           div(class="card",
                               h4("Résultats de l'Analyse de l'Xlsform (Résumé)"),
                               p("Le tableau ci‑dessous montre le nombre d'erreurs détectées pour chaque catégorie."),
                               DTOutput("summary_error_table")
                           ),
                           div(class="card",
                               h4("Structure du Formulaire"),
                               p("Les Level cibas montre le niveau des section, si level 1 == Section ; si level 2 == Sous section, ainsi de suite "),
                               DTOutput("section_summary_table")
                           )
                  ),
                  tabPanel("Critique", value = "critical", DTOutput("critical_errors")),
                  tabPanel("Structure/Données", value = "data_errors", DTOutput("data_errors_table")),
                  tabPanel("Amélioration/Info", value = "improvement", DTOutput("improvement_table"))
      ),
      tags$script('
        $(document).on("shown.bs.tab", function(){
          $(".dataTable").each(function(){
            try { $(this).DataTable().columns.adjust().draw(); } catch(e){}
          });
        });
      ')
    )
  ),
  
  
  tags$head(tags$style(HTML("
  /* Forcer l'alignement à gauche et centrage vertical */
  table.dataTable td, table.dataTable th {
    text-align: left !important;
    vertical-align: middle !important;
  }
  /* Police uniforme */
  .dataTables_wrapper table.dataTable {
    font-family: 'Segoe UI', Arial, sans-serif;
    font-size: 14px;
  }
  /* Scroll horizontal si nécessaire */
  .dataTables_wrapper .dataTables_scrollBody {
    overflow-x: auto;
  }
")))
  
  
)

render_structure_ui <- function(tree) {
  tagList(lapply(tree, function(section) {
    wellPanel(
      h4(section$title),
      # 1) Questions directes de la section
      if (length(section$questions) > 0) {
        tags$ul(lapply(section$questions, function(q) tags$li(q$label)))
      },
      # 2) Sous-sections en bas
      if (length(section$children) > 0) render_structure_ui(section$children)
    )
  }))
}

# ---------------------------------------------------------------------
# Server
# ---------------------------------------------------------------------
server <- function(input, output, session) {
  rv <- reactiveValues(errors = NULL, languages = NULL, selected_lang = "Auto", survey_data = NULL, summary_final = NULL)
  
  detect_languages <- function(filepath) {
    survey <- read_excel(filepath, sheet = "survey", col_types = "text")
    cols <- names(survey)
    langs <- tolower(cols)
    has_lang <- any(stringr::str_detect(tolower(cols), "^label::") | stringr::str_detect(tolower(cols), "^hint::"))
    if (has_lang) c("Auto","French","English","Malagasy") else c("Auto (Une seule langue détectée)")
  }
  
  observeEvent(input$detect_lang_btn, {
    req(input$file1)
    filepath <- input$file1$datapath
    tryCatch({
      rv$languages <- detect_languages(filepath)
      showNotification("Langues détectées avec succès.", type = "message", duration = 3)
    }, error = function(e) {
      rv$languages <- c("Auto (Erreur lecture langues)")
      showNotification(glue("Erreur lors de la détection des langues : {e$message}"), type = "error", duration = NULL)
    })
  })
  
  output$language_select_ui <- renderUI({
    if (!is.null(rv$languages)) selectInput("selected_language", "Choisir la langue d'affichage :", choices = rv$languages, selected = "Auto") else p("Cliquez sur 'Détecter les langues' d'abord.")
  })
  
  # Analyse erreurs
  observeEvent(input$validate_btn, {
    req(input$file1)
    withProgress(message = "Analyse en cours...", value = 0, {
      filepath <- input$file1$datapath
      rv$survey_data <- read_excel(filepath, sheet = "survey", col_types = "text") %>% mutate(across(everything(), as.character))
      rv$errors <- validate_xlsform(filepath)
      
      # Summary par catégories
      summary_data_raw <- rv$errors %>% mutate(
        Category = case_when(
          Type == "Duplication" ~ "Duplication de noms",
          Type == "Choices Non Disponibles" ~ "Listes de choix manquantes",
          Type == "Colonnes Manquantes" ~ "Colonnes obligatoires manquantes",
          Type == "Références Invalides (relevant)" ~ "Références invalides (relevant)",
          Type == "Références Invalides (calculation)" ~ "Références Invalides (calculation)",
          Type == "Structure" ~ "Erreurs de structure (group/repeat)",
          TRUE ~ "Autres/Amélioration/Info"
        )
      ) %>% group_by(Category) %>% summarise(NombreErreurs = n(), .groups = 'drop') %>% rename(TypeErreur = Category)
      
      all_categories <- c("Duplication de noms","Listes de choix manquantes","Colonnes obligatoires manquantes","Références Invalides (calculation)","Références invalides (relevant)","Erreurs de structure (group/repeat)","Autres/Amélioration/Info")
      rv$summary_final <- data.frame(TypeErreur = all_categories, NombreErreurs = 0L, stringsAsFactors = FALSE) %>%
        left_join(summary_data_raw, by = "TypeErreur") %>%
        mutate(NombreErreurs = coalesce(NombreErreurs.y, NombreErreurs.x, 0L)) %>%
        select(TypeErreur, NombreErreurs)
      
      updateTabsetPanel(session, "errorTabs", selected = "summary")
      incProgress(1)
    })
  })
  
  
  render_table <- function(data) {
    datatable(
      data,
      rownames = FALSE,
      class = "stripe hover", # bandes alternées + survol
      options = list(
        pageLength = 10,
        autoWidth = TRUE,
        scrollX = TRUE,
        columnDefs = list(
          list(className = 'dt-left', targets = "_all") # tout à gauche
        )
      )
    )
  }
  
  
  
  output$summary_error_table <- renderDT({ req(rv$summary_final); render_table(rv$summary_final) })
  
  
  output$structure_ui <- renderUI({
    req(rv$survey_data)
    lbl_col <- detect_label_col(rv$survey_data)
    tree <- compute_structure_tree(rv$survey_data, lbl_col, exclude_types = EXCLUDE_TYPES) # fonction auxiliaire pour arborescence
    render_structure_ui(tree)
  })
  
  
  output$section_summary_table <- renderDT({
    req(rv$survey_data)
    lbl_col <- detect_label_col(rv$survey_data)
    summary_df <- compute_structure_summary(rv$survey_data, lbl_col, exclude_types = EXCLUDE_TYPES)
    render_table(summary_df)
  })
  
  
  output$critical_errors <- renderDT({ req(rv$errors); render_table(rv$errors %>% filter(Type %in% c("Erreur Critique","Nom Invalide","Duplication","Choices Non Disponibles","Structure","Références Invalides (calculation)","Références invalides (relevant)"))) })
  output$data_errors_table <- renderDT({ req(rv$errors); render_table(rv$errors %>% filter(Type %in% c("Colonnes Manquantes","Données Manquantes"))) })
  output$improvement_table <- renderDT({ req(rv$errors); render_table(rv$errors %>% filter(!(Type %in% c("Erreur Critique","Nom Invalide","Duplication","Choices Non Disponibles","Colonnes Manquantes","Données Manquantes","Structure")))) })
  
  output$download_excel_report <- downloadHandler(
    filename = function() { paste0("Rapport_erreurs_", Sys.Date(), ".xlsx") },
    content  = function(file) { req(rv$errors); write_xlsx(rv$errors, path = file) }
  )
  
  output$download_word <- downloadHandler(
    filename = function() {
      # Nom de fichier par défaut même si input$file1 n'est pas encore prêt.
      base <- if (!is.null(input$file1)) tools::file_path_sans_ext(input$file1$name) else "formulaire"
      lang_suffix <- ifelse(identical(input$selected_language, "Auto") || is.null(input$selected_language), "", paste0("_", tolower(input$selected_language)))
      paste0(base, "_rendu", lang_suffix, ".docx")
    },
    content = function(file) {
      req(input$file1, input$selected_language)
      withProgress(message = "Génération Word...", value = 0, {

        xlsform_to_wordRev(xlsx = input$file1$datapath, output_path = file)
        incProgress(1)
        showNotification("Document Word généré avec succès !", type = "message")
      })
    }
  )
  

  outputOptions(output, "critical_errors", suspendWhenHidden = FALSE)
  outputOptions(output, "data_errors_table", suspendWhenHidden = FALSE)
  outputOptions(output, "improvement_table", suspendWhenHidden = FALSE)
  outputOptions(output, "summary_error_table", suspendWhenHidden = FALSE)
  outputOptions(output, "section_summary_table", suspendWhenHidden = FALSE)
}

# Lancer l'app
shinyApp(ui = ui, server = server)
