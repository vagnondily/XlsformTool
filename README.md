
# XLSFormTools v1.1 — Validation & Conversion Word (Shiny)

## Description
XLSFormTools est une application **Shiny** permettant :
- Faire une premiere analyse des **erreurs** et **validation** des formulaires XLSForm (.xlsx) selon les standards ODK/XForms.
- La **conversion** en document Word (.docx) avec structure hiérarchique, titres, placeholders et options formatées.

Elle détecte les erreurs critiques, propose des suggestions correctives et génère un rendu professionnel du questionnaire.

---

## Fonctionnalités principales
- ✅ Analyse des erreurs : colonnes manquantes, types invalides, duplications, structure begin/end group/repeat.
- ✅ Résumé global : tableau des erreurs par catégorie.
- ✅ Structure du formulaire : sections, sous-sections, nombre de questions.
- ✅ Conversion Word : titrage, placeholders, options (○ / ☐), traduction des conditions `relevant` en français quasi-naturel.
- ✅ Support multilingue : détection automatique des colonnes `label::lang` et `hint::lang`.

---

## Règles de rendu Word
### Styles & Design
- Couleurs : BLUE (#0A66C2), DARK_BLUE (#001F3F), GREY_BG (#F2F2F2), GREY_TXT (#777777), RED_TXT (#C00000).
- Police : Cambria (Body).
- Tailles : Titre section (14), Sous-section (12), Bloc repeat (12), Question (11), Métadonnées (9), Hint (9 italique), Relevant (9 rouge).
- Indentation : Question ~0,3″ ; Contenu ~0,5″.
- Espacement : line_spacing = 1.0.

### Structure hiérarchique
- Sections : `Section X : <label>`.
- Sous-sections : `Sous-section X.Y : <label>`.
- Blocs repeat : affichés avec symbole 🔁.
- Fin de bloc : `--- Fin du bloc ---`.

### Questions
- Format : `N° Question. Label (name – type)`.
- Hint : italique sous la question.
- Relevant : traduit en français quasi-naturel (ex. « Afficher si : … »).

### Placeholders
- integer → [insérer un entier].
- decimal → [insérer un décimal].
- date → [insérer une date].
- geopoint → [capturer les coordonnées GPS].
- image/photo → [prendre une photo].
- audio/video → [enregistrer ou sélectionner un média].
- ......
- autres → [insérer votre réponse ici].

### Options de choix
- select_one : symbole `○`.
- select_multiple : symbole `☐`.

### Conditions (relevant)
- Traduction XPath → français : and=et, or=ou, not=non, = est égal à, != est différent de, > est supérieur à, < est inférieur à.
- selected(${var}, 'code') → « `<label>` a l'option `«<choix>»` cochée ».
- count-selected(...) >= 1 → « Au moins une option est cochée pour … ».

### Exclusions
- Types ignorés : calculate, start, end, today, deviceid, etc.
- Pas d’expansion des groupes répétés (ignore repeat_count, indexed-repeat()).

---

## Installation
```R
install.packages(c('readxl','dplyr','stringr','tidyr','purrr','officer','flextable','glue','tools','tibble','rlang','shiny','DT','writexl','htmltools','shinythemes'))
``

