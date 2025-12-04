
{
  "name": "XLSFormTools v1.1 — Validation & Conversion Word (Shiny)",
  "description": {
    "short": "Application Shiny pour analyser des formulaires XLSForm (.xlsx), détecter les erreurs de conformité ODK/XForms et générer un rendu Word (.docx) formaté.",
    "details": [
      "Validation technique et logique (types, colonnes obligatoires, références inconnues, structure begin/end group/repeat, duplications).",
      "Résumé global des erreurs par catégories et récapitulatif de la structure (Sections/Sous-sections avec comptage des questions).",
      "Conversion du XLSForm vers un document Word hiérarchisé, avec titrage, placeholders de réponse, et rendu des options pour select_one/multiple (○ / ☐).",
      "Détection dynamique des colonnes label/hint y compris label::French/English/Malagasy (mode Auto et sélection manuelle)."
    ]
  },
  "core_functions": {
    "validate_xlsform": "Analyse et retourne un tableau des erreurs (feuille, ligne, type, description, suggestion).",
    "xlsform_to_wordRev": "Convertit le XLSForm en .docx avec hiérarchie (Section, Sous-section, Bloc), numérotation, hints, relevant (phrase naturelle).",
    "compute_structure_summary": "Calcule un tableau synthétique des blocs (level, titre, nb_questions).",
    "translate_relevant": "Reformule les expressions 'relevant' en texte lisible (fr), en mappant variables et choix avec leurs labels."
  },
  "ui": {
    "theme": "shinythemes::flatly",
    "inputs": [
      {"id": "file1", "label": "Choisir le fichier XLSForm (.xlsx)", "type": "fileInput"},
      {"id": "detect_lang_btn", "label": "Détecter les langues", "type": "actionButton"},
      {"id": "selected_language", "label": "Choisir la langue d'affichage", "type": "selectInput", "source": "language_select_ui"},
      {"id": "validate_btn", "label": "Analyser les erreurs", "type": "actionButton"},
      {"id": "download_word", "label": "Générer Word (.docx)", "type": "downloadButton"},
      {"id": "download_excel_report", "label": "Rapport Excel (.xlsx)", "type": "downloadButton"}
    ],
    "tabs": [
      {"id": "summary", "title": "Résumé Général", "outputs": ["summary_error_table", "section_summary_table"]},
      {"id": "critical", "title": "Critique", "outputs": ["critical_errors"]},
      {"id": "data_errors", "title": "Structure/Données", "outputs": ["data_errors_table"]},
      {"id": "improvement", "title": "Amélioration/Info", "outputs": ["improvement_table"]}
    ]
  },
  "downloads": [
    {
      "id": "download_word",
      "filename_pattern": "<base_form_name>_rendu_<lang>.docx",
      "source": "xlsform_to_wordRev(input$file1$datapath, output_path=file)"
    },
    {
      "id": "download_excel_report",
      "filename_pattern": "Rapport_erreurs_<YYYY-MM-DD>.xlsx",
      "source": "write_xlsx(rv$errors, path=file)"
    }
  ],
  "features": {
    "word_rendering": {
      "styles": {
        "colors": {"WFP_BLUE": "#0A66C2", "WFP_DARK_BLUE": "#001F3F", "GREY_BG": "#F2F2F2", "GREY_TXT": "#777777", "RED_TXT": "#C00000", "WHITE_TXT": "#FFFFFF"},
        "fonts": {"family": "Cambria (Body)", "sizes": {"title": 14, "subtitle": 12, "block": 12, "meta": 10, "hint": 9, "relevant": 9, "q_blue": 11, "q_grey": 9}},
        "layout": {"line_spacing": 1.0, "indent_question_inches": 0.3, "indent_content_inches": 0.5}
      },
      "exclusions": {
        "types": ["calculate", "start", "end", "today", "deviceid", "subscriberid", "simserial", "phonenumber", "username", "instanceid", "end_group", "end group"],
        "names": ["start", "end", "today", "deviceid", "username", "instanceID", "instanceid"]
      },
      "choices_rendering": {"select_one_symbol": "○", "select_multiple_symbol": "☐"},
      "relevant_translation": "Expressions XPath/XForm traduites en français naturel (et, ou, non, égal/différent/supérieur/inférieur; selected(), count-selected(), jr:choice-name).",
      "repeat_handling": "Affiche les blocs begin_repeat sans étendre les répétitions; indique la fin de bloc; ignore l'évaluation repeat_count/indexed-repeat()."
    },
    "validation": {
      "checks": [
        "Colonnes obligatoires manquantes (type, name, label).",
        "Types non conformes et casse incorrecte (doit être en minuscules).",
        "Appearances non reconnues.",
        "Formules 'calculate' vides ou fonctions non supportées (ex: '/').",
        "Références inconnues dans calculation/relevant/constraint/choice_filter.",
        "Utilisation de jr:choice-name et règles avec select_multiple.",
        "Duplications de noms (survey) et de choix (par list_name).",
        "Structure begin/end group/repeat non appariée.",
        "Listes de choix manquantes par rapport aux select_one/multiple."
      ],
      "outputs": "Tableau consolidé des erreurs avec suggestions correctives."
    },
    "structure_summary": {
      "description": "Synthèse ordonnée des blocs (group/repeat et Sans Section) avec le nombre de questions par bloc.",
      "columns": ["Level", "Titre", "Nb_Questions"]
    }
  },
  "requirements": {
    "r_version": ">= 4.1",
    "packages": ["readxl", "dplyr", "stringr", "tidyr", "purrr", "officer", "flextable", "glue", "tools", "tibble", "rlang", "shiny", "DT", "writexl", "htmltools", "shinythemes"]
  },
  "setup": {
    "install": ["install.packages(c('readxl','dplyr','stringr','tidyr','purrr','officer','flextable','glue','tools','tibble','rlang','shiny','DT','writexl','htmltools','shinythemes'))"],
    "optional_assets": [{"LOGO_PATH": "cp_logo.png", "note": "Logo affiché dans l'en-tête du document Word si présent dans le répertoire de l'app."}]
  },
  "run": {
    "command": "R -e \"shiny::runApp('app.R', host='0.0.0.0', port=1234)\"",
    "entry": "app.R",
    "notes": [
      "Lancement depuis le dossier contenant app.R et les assets (ex: cp_logo.png).",
      "L'app utilise un document Word vierge (officer::read_docx())."
    ]
  },
  "usage": {
    "steps": [
      "Cliquez sur 'Choisir le fichier' et sélectionnez votre XLSForm (.xlsx).",
      "Cliquez sur 'Détecter les langues' (optionnel) puis choisissez la langue d'affichage (Auto/French/English/Malagasy).",
      "Cliquez sur 'Analyser les erreurs' pour générer le résumé et les tables détaillées.",
      "Téléchargez 'Rapport Excel' pour les erreurs ou 'Générer Word' pour le rendu du questionnaire."
    ],
    "tips": [
      "Assurez-vous que les onglets survey/choices/settings existent et que 'list_name' est défini dans choices.",
      "Évitez les espaces et caractères spéciaux dans 'name' (autorisé: lettres, chiffres, underscore).",
      "Utilisez label::lang et hint::lang pour les formulaires multilingues; sinon 'label'/'hint' simples suffisent."
    ]
  },
  "limitations": [
    "Ne pas évaluer ni étendre les groupes répétés via repeat_count ou indexed-repeat().",
    "Certaines appearances ODK non listées peuvent ne pas être reconnues; suivez la spec officielle.",
    "La traduction de relevant est centrée sur le français; adapter si besoin."
  ],
  "links": {
    "odk_docs": "https://docs.getodk.org",
    "xlsform_guide": "https://xlsform.org/en/"
  },
  "authoring": {
    "maintainer": "Manantsoa VAGNONDILY",
    "role": "Monitoring Evaluation Associate",
    "contact": "(à compléter)",
    "version": "1.1",
    "last_update": "2025-12-04"
  },
  "license": "MIT (à confirmer)"
}
