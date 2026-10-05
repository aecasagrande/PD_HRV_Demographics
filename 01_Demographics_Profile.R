# ==============================================================================
# SCRIPT 1: DEMOGRAPHICS & CLINICAL CHARACTERISTICS PROFILE
# ==============================================================================
# - Loads & merges Demographics, Cognitive, Motor UPDRS-III, Giladi FOG and Home
#   Questionnaires, with explicit numeric coercion and data-quality warnings
# - Generates Table 1 (Full Sample and Excluding P-007) via gtsummary, and
#   exports it to Word (.docx), HTML and high-resolution PNG in Results/
# - Prints a missing-data audit (per variable and per participant)
# - Produces faceted boxplots with clinical reference bands, individual
#   participant points and mean diamonds (Results/Demographics_Distributions.png)
#
# Usage:  Rscript 01_Demographics_Profile.R
# Data directory (first that exists):
#   1. $PD_HRV_DATA_DIR
#   2. /Users/aecasagr/Desktop/PD_HRV/Data
#   3. Data/ next to this script
# ==============================================================================

t_start <- Sys.time()

# 1. Packages ------------------------------------------------------------------
needed_pkgs <- c("readxl", "tidyverse", "gtsummary", "cardx", "gt", "flextable",
                 "officer", "RColorBrewer")
to_install <- needed_pkgs[!vapply(needed_pkgs, requireNamespace, logical(1), quietly = TRUE)]
if (length(to_install)) install.packages(to_install, repos = "https://cloud.r-project.org")

suppressPackageStartupMessages({
  library(readxl)
  library(tidyverse)
  library(gtsummary)
  library(gt)
  library(flextable)
  library(officer)
  library(RColorBrewer)
})

# 2. Configuration -------------------------------------------------------------
PID_PATTERN <- "^P-\\d{3}$"
EXCLUDE_IDS <- c("P-007")           # participants dropped in the sensitivity table
SALVAGE_FORMATTED_NUMBERS <- TRUE   # recover "45s", "1,200", "27/30" (with a warning)

# Strings that unambiguously encode "missing" (case-insensitive, after trimming)
NA_TOKENS <- c("", "na", "n/a", "n.a.", "nan", "null", "none", "-", "--", ".",
               "missing", "not done", "nd", "#n/a", "#value!", "#div/0!")

banner <- function(title) cat("\n===", title, "===\n")

n_warnings <- 0L
log_warn <- function(...) {
  n_warnings <<- n_warnings + 1L
  cat("  [WARN] ", ..., "\n", sep = "")
}
log_info <- function(...) cat("  [INFO] ", ..., "\n", sep = "")

get_script_dir <- function() {
  file_arg <- grep("^--file=", commandArgs(trailingOnly = FALSE), value = TRUE)
  if (length(file_arg)) return(dirname(normalizePath(sub("^--file=", "", file_arg[1]))))
  getwd()
}
SCRIPT_DIR <- get_script_dir()

# 3. File Paths ----------------------------------------------------------------
banner("RESOLVING FILE PATHS")

data_dir_candidates <- c(Sys.getenv("PD_HRV_DATA_DIR", unset = ""),
                         "/Users/aecasagr/Desktop/PD_HRV/Data",
                         file.path(SCRIPT_DIR, "Data"))
data_dir_candidates <- data_dir_candidates[nzchar(data_dir_candidates)]
DATA_DIR <- data_dir_candidates[dir.exists(data_dir_candidates)][1]
if (is.na(DATA_DIR)) {
  stop("No data directory found. Checked:\n  ",
       paste(data_dir_candidates, collapse = "\n  "),
       "\nSet PD_HRV_DATA_DIR or place the workbooks in Data/ next to this script.")
}
RESULTS_DIR <- file.path(SCRIPT_DIR, "Results")
dir.create(RESULTS_DIR, showWarnings = FALSE, recursive = TRUE)

# Each source: workbook, sheet, and analysis name -> exact source column header
SOURCES <- list(
  Demographics = list(
    file = "01_Demographics.xlsx", sheet = "Demographics",
    cols = c(Sex              = "Gender (self-report)",
             Age              = "Age (years)",
             Disease_Duration = "Disease duration (years since diagnosis)",
             Symptom_Duration = "Symptom duration (years)",
             LEDD             = "LEDD_total_mg",
             BMI              = "BMI (kg/m2)")),
  Cognitive = list(
    file = "02_Cognitive.xlsx", sheet = "Cognitive",
    cols = c(MoCA  = "MoCA (/30)",
             TMT_A = "TMT-A (s)",
             TMT_B = "TMT-B (s)",
             dTMT  = "TMT B - A (s)")),
  UPDRS = list(
    file = "03_Motor_UPDRS-III.xlsx", sheet = "UPDRS-III",
    cols = c(UPDRS_OFF = "UPDRS-III OFF",
             UPDRS_ON  = "UPDRS-III ON")),
  Giladi = list(
    file = "04_Freezing_Giladi.xlsx", sheet = "Summary",
    cols = c(Giladi_OFF = "OFF Total sum",
             Giladi_ON  = "ON Total sum")),
  Home = list(
    file = "06_Home_Questionnaires.xlsx", sheet = "Summary",
    cols = c(NFOG_Q     = "NFOG-Q total",
             STAI_Trait = "STAI-Trait total"))
)
CATEGORICAL_VARS <- c("Sex")

cat(sprintf("  Data directory   : %s\n", DATA_DIR))
cat(sprintf("  Results directory: %s\n", RESULTS_DIR))
for (src in SOURCES) {
  cat(sprintf("  %-28s %s\n", src$file,
              if (file.exists(file.path(DATA_DIR, src$file))) "found" else "NOT FOUND"))
}

# Plausible ranges: values outside are flagged (never altered or dropped)
PLAUSIBLE_RANGES <- tribble(
  ~Variable,           ~lo,  ~hi,
  "Age",                18,  110,
  "Disease_Duration",    0,   60,
  "Symptom_Duration",    0,   60,
  "LEDD",                0, 4000,
  "BMI",                13,   60,
  "MoCA",                0,   30,
  "TMT_A",               1,  300,
  "TMT_B",               1,  300,
  "dTMT",             -100,  300,
  "UPDRS_OFF",           0,  132,
  "UPDRS_ON",            0,  132,
  "Giladi_OFF",          0,   64,   # 8 tasks x (Dur 0-4 + Freq 0-4)
  "Giladi_ON",           0,   64,
  "NFOG_Q",              0,   28,
  "STAI_Trait",         20,   80
)

# 4. Load, Validate & Coerce ---------------------------------------------------
banner("LOADING & VALIDATING SOURCE WORKBOOKS")

# Coerce a character vector to numeric, reporting every value that is not a
# clean number. Missing-value codes become NA quietly (reported once);
# recoverable formatting is fixed with a warning; anything else becomes NA
# with a warning.
coerce_numeric <- function(x, pids, var, source_name) {
  x_chr  <- str_squish(as.character(x))
  is_tok <- is.na(x_chr) | tolower(x_chr) %in% NA_TOKENS
  out    <- suppressWarnings(as.numeric(ifelse(is_tok, NA, x_chr)))

  coded <- !is.na(x_chr) & x_chr != "" & is_tok
  if (any(coded)) {
    log_info(sprintf("%s$%s: %d missing-value code(s) read as NA: %s", source_name, var,
                     sum(coded), paste0(pids[coded], " ('", x_chr[coded], "')", collapse = ", ")))
  }

  for (i in which(!is_tok & is.na(out))) {
    raw <- x_chr[i]
    fixed <- NA_real_
    if (SALVAGE_FORMATTED_NUMBERS) {
      s <- gsub("\u2212", "-", raw)                                  # unicode minus
      if (grepl("^-?\\d{1,3}(,\\d{3})+(\\.\\d+)?$", s)) s <- gsub(",", "", s)  # 1,200
      s <- sub("^(-?\\d+(\\.\\d+)?)\\s*/\\s*\\d+$", "\\1", s)        # 27/30
      s <- sub("^(-?\\d+(\\.\\d+)?)\\s*(s|sec|secs|seconds|mg|kg|cm|yrs?|years|%)$",
               "\\1", s, ignore.case = TRUE)                         # 45s, 600 mg
      fixed <- suppressWarnings(as.numeric(s))
    }
    if (!is.na(fixed)) {
      log_warn(sprintf("%s$%s [%s]: formatted value '%s' coerced to %s - verify source",
                       source_name, var, pids[i], raw, format(fixed)))
      out[i] <- fixed
    } else {
      log_warn(sprintf("%s$%s [%s]: non-numeric value '%s' set to NA",
                       source_name, var, pids[i], raw))
    }
  }

  if (!is.numeric(out)) stop(sprintf("Coercion of %s$%s did not yield a numeric vector", source_name, var))
  if (any(is.infinite(out))) {
    log_warn(sprintf("%s$%s: infinite value(s) set to NA", source_name, var))
    out[is.infinite(out)] <- NA_real_
  }
  out
}

empty_source <- function(src) {
  tibble(Participant_ID = character(),
         !!!setNames(rep(list(numeric()), length(src$cols)), names(src$cols)))
}

load_source <- function(src, source_name) {
  path <- file.path(DATA_DIR, src$file)
  cat(sprintf("\n  -- %s (%s / sheet '%s')\n", source_name, src$file, src$sheet))

  if (!file.exists(path)) {
    log_warn(sprintf("%s not found; its %d variable(s) will be entirely missing",
                     src$file, length(src$cols)))
    return(empty_source(src))
  }
  sheets <- excel_sheets(path)
  if (!src$sheet %in% sheets) {
    log_warn(sprintf("sheet '%s' not in %s (available: %s); source skipped",
                     src$sheet, src$file, paste(sheets, collapse = ", ")))
    return(empty_source(src))
  }

  # Read every cell as text so that type coercion is explicit and auditable
  raw <- read_excel(path, sheet = src$sheet, col_types = "text", .name_repair = "minimal")
  names(raw) <- str_squish(names(raw))

  if (!"participant_id" %in% names(raw)) {
    log_warn(sprintf("no 'participant_id' column in %s; source skipped", src$file))
    return(empty_source(src))
  }

  # Filter rows strictly by participant pattern
  ids  <- str_squish(raw$participant_id)
  keep <- !is.na(ids) & grepl(PID_PATTERN, ids)
  dropped <- unique(ids[!keep & !is.na(ids)])
  near_miss <- dropped[grepl("^\\s*p\\s*-?\\s*\\d+\\s*$", dropped, ignore.case = TRUE)]
  if (length(near_miss)) {
    log_warn(sprintf("malformed participant ID(s) excluded - fix in source: %s",
                     paste0("'", near_miss, "'", collapse = ", ")))
  }
  summary_rows <- setdiff(dropped, near_miss)
  if (length(summary_rows)) {
    log_info(sprintf("dropped %d summary/group row(s): %s", length(summary_rows),
                     paste(summary_rows, collapse = "; ")))
  }
  raw <- raw[keep, , drop = FALSE]
  raw$participant_id <- ids[keep]

  dups <- unique(raw$participant_id[duplicated(raw$participant_id)])
  if (length(dups)) {
    log_warn(sprintf("duplicate participant row(s) %s; keeping the first occurrence",
                     paste(dups, collapse = ", ")))
    raw <- raw[!duplicated(raw$participant_id), , drop = FALSE]
  }

  missing_cols <- setdiff(src$cols, names(raw))
  if (length(missing_cols)) {
    log_warn(sprintf("expected column(s) not found in %s: %s", src$file,
                     paste0("'", missing_cols, "'", collapse = ", ")))
  }

  # Surface verification flags written into the workbook by the build script
  if ("double_check_items" %in% names(raw)) {
    flagged <- which(!is.na(raw$double_check_items) & nzchar(str_squish(raw$double_check_items)))
    if (length(flagged)) {
      log_info(sprintf("%d participant(s) carry source double_check_items: %s", length(flagged),
                       paste(raw$participant_id[flagged], collapse = ", ")))
    }
  }

  out <- tibble(Participant_ID = raw$participant_id)
  for (v in names(src$cols)) {
    col <- src$cols[[v]]
    vals <- if (col %in% names(raw)) raw[[col]] else rep(NA_character_, nrow(raw))
    if (v %in% CATEGORICAL_VARS) {
      vals <- str_squish(vals)
      vals[tolower(vals) %in% NA_TOKENS] <- NA
      out[[v]] <- vals
    } else {
      out[[v]] <- coerce_numeric(vals, raw$participant_id, v, source_name)
    }
  }
  cat(sprintf("  Loaded %d participant row(s), %d variable(s)\n", nrow(out), length(src$cols)))
  out
}

loaded <- imap(SOURCES, load_source)

# Merge onto the demographic roster (the cohort definition, README [D1])
df_demo <- reduce(loaded[-1], left_join, by = "Participant_ID", .init = loaded$Demographics) %>%
  mutate(Sex = factor(Sex)) %>%
  arrange(Participant_ID)

for (nm in names(loaded)[-1]) {
  orphans <- setdiff(loaded[[nm]]$Participant_ID, df_demo$Participant_ID)
  if (length(orphans)) {
    log_warn(sprintf("%s has participant(s) absent from the demographic roster (not merged): %s",
                     nm, paste(orphans, collapse = ", ")))
  }
}

banner("TYPE & PLAUSIBILITY CHECKS")
VAR_NAMES <- unlist(lapply(SOURCES, function(s) names(s$cols)), use.names = FALSE)
CONT_VARS <- setdiff(VAR_NAMES, CATEGORICAL_VARS)

bad_types <- CONT_VARS[!vapply(df_demo[CONT_VARS], is.numeric, logical(1))]
if (length(bad_types)) stop("Non-numeric clinical variable(s) after coercion: ", paste(bad_types, collapse = ", "))
cat(sprintf("  All %d clinical variables are numeric; %s is a factor (levels: %s)\n",
            length(CONT_VARS), CATEGORICAL_VARS, paste(levels(df_demo$Sex), collapse = ", ")))

for (v in VAR_NAMES) {
  if (all(is.na(df_demo[[v]]))) {
    log_warn(sprintf("%s is entirely missing - if it is a formula column, open and save the ",
                     v), "workbook in Excel/LibreOffice so cached values exist (README [D15])")
  }
}
for (i in seq_len(nrow(PLAUSIBLE_RANGES))) {
  r <- PLAUSIBLE_RANGES[i, ]
  x <- df_demo[[r$Variable]]
  out_rng <- which(!is.na(x) & (x < r$lo | x > r$hi))
  for (j in out_rng) {
    log_warn(sprintf("%s [%s] = %s is outside the plausible range %s-%s - verify source",
                     r$Variable, df_demo$Participant_ID[j], format(round(x[j], 2)), r$lo, r$hi))
  }
}

excluded_present <- intersect(EXCLUDE_IDS, df_demo$Participant_ID)
if (length(setdiff(EXCLUDE_IDS, excluded_present))) {
  log_warn(sprintf("exclusion ID(s) %s not in the roster; the 'Excluding' table equals the full sample",
                   paste(setdiff(EXCLUDE_IDS, excluded_present), collapse = ", ")))
}
df_excl <- df_demo %>% filter(!Participant_ID %in% EXCLUDE_IDS)
cat(sprintf("\n  Roster: %d participants (%d after excluding %s)\n",
            nrow(df_demo), nrow(df_excl), paste(EXCLUDE_IDS, collapse = ", ")))

# 5. Table 1 (gtsummary) -------------------------------------------------------
banner("TABLE 1: DISTRIBUTION CHECKS")

VAR_LABELS <- c(
  Sex              = "Sex (self-reported gender)",
  Age              = "Age (years)",
  Disease_Duration = "Disease duration (years)",
  Symptom_Duration = "Symptom duration (years)",
  LEDD             = "Total LEDD (mg/day)",
  BMI              = "BMI (kg/m\u00b2)",
  MoCA             = "MoCA (/30)",
  TMT_A            = "Trail Making Test A (s)",
  TMT_B            = "Trail Making Test B (s)",
  dTMT             = "TMT B \u2212 A (s)",
  UPDRS_OFF        = "MDS-UPDRS Part III, OFF",
  UPDRS_ON         = "MDS-UPDRS Part III, ON",
  Giladi_OFF       = "Giladi FOG total, OFF",
  Giladi_ON        = "Giladi FOG total, ON",
  NFOG_Q           = "NFOG-Q total",
  STAI_Trait       = "STAI-Trait total"
)

# Standard error of the mean, used by name inside the gtsummary statistic template
se <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) < 2) NA_real_ else sd(x) / sqrt(length(x))
}

STAT_NORMAL <- "{mean} \u00b1 {sd} (SE: {se}); Range: {min} \u2013 {max}"
STAT_SKEWED <- "{median} [{p25} \u2013 {p75}]; Range: {min} \u2013 {max}"

# Force a variable's summary irrespective of the test, e.g. c(LEDD = "skewed")
DIST_OVERRIDE <- c()

# Skewed = Shapiro-Wilk p < .05 or |skewness| > 1, assessed once on the full
# sample so both tables report each variable with the same statistic
classify_distribution <- function(x) {
  x <- x[!is.na(x)]
  if (length(x) < 3 || length(unique(x)) < 3) {
    return(tibble(n = length(x), sw_p = NA_real_, skew = NA_real_, dist = "normal"))
  }
  m2 <- mean((x - mean(x))^2)
  skew <- mean((x - mean(x))^3) / m2^1.5
  sw_p <- shapiro.test(x)$p.value
  tibble(n = length(x), sw_p = sw_p, skew = skew,
         dist = if (sw_p < 0.05 || abs(skew) > 1) "skewed" else "normal")
}

dist_check <- map(CONT_VARS, ~ classify_distribution(df_demo[[.x]])) %>%
  list_rbind() %>%
  mutate(Variable = CONT_VARS, .before = 1) %>%
  mutate(dist = coalesce(unname(DIST_OVERRIDE[Variable]), dist))

NORMAL_VARS <- dist_check$Variable[dist_check$dist == "normal"]
SKEWED_VARS <- dist_check$Variable[dist_check$dist == "skewed"]

# Print a data frame as an aligned plain-text table; `wrap` names a column whose
# long entries are wrapped onto continuation lines
print_aligned <- function(df, wrap = NULL, wrap_width = 60, right = NULL) {
  df <- as.data.frame(lapply(df, function(x) ifelse(is.na(x), "", as.character(x))),
                      check.names = FALSE, stringsAsFactors = FALSE)
  rows <- list()
  for (i in seq_len(nrow(df))) {
    cells <- as.list(df[i, , drop = FALSE])
    pieces <- if (!is.null(wrap)) strwrap(cells[[wrap]], width = wrap_width) else character()
    if (!is.null(wrap) && length(pieces) > 1) {
      cells[[wrap]] <- pieces[1]
      rows[[length(rows) + 1]] <- cells
      for (p in pieces[-1]) {
        cont <- setNames(as.list(rep("", ncol(df))), names(df))
        cont[[wrap]] <- p
        rows[[length(rows) + 1]] <- cont
      }
    } else {
      rows[[length(rows) + 1]] <- cells
    }
  }
  body <- bind_rows(lapply(rows, as_tibble))
  widths <- vapply(names(df), function(n) max(nchar(c(n, body[[n]]), type = "width")), numeric(1))
  fmt <- function(vals) {
    paste(vapply(seq_along(vals), function(k) {
      formatC(vals[k], width = widths[k], flag = if (names(df)[k] %in% right) " " else "-")
    }, character(1)), collapse = "  ")
  }
  cat(" ", fmt(names(df)), "\n")
  cat(" ", paste(strrep("-", widths), collapse = "  "), "\n")
  for (i in seq_len(nrow(body))) cat(" ", fmt(unlist(body[i, ])), "\n")
}

print_aligned(
  dist_check %>% transmute(Variable, n,
                           `Shapiro-Wilk p` = ifelse(is.na(sw_p), "-", formatC(sw_p, format = "f", digits = 3)),
                           Skewness = ifelse(is.na(skew), "-", formatC(skew, format = "f", digits = 2)),
                           Summary = ifelse(dist == "normal", "mean \u00b1 SD (SE); range",
                                            "median [IQR]; range")),
  right = c("n", "Shapiro-Wilk p", "Skewness")
)

# The JAMA theme appends a statistic label to each row label, derived from the
# raw template (e.g. "SE: se"); it is replaced with a compact descriptor
STAT_DESCRIPTOR <- c(setNames(rep("mean \u00b1 SD (SE); range", length(NORMAL_VARS)), NORMAL_VARS),
                     setNames(rep("median [IQR]; range", length(SKEWED_VARS)), SKEWED_VARS),
                     setNames(rep("n (%)", length(CATEGORICAL_VARS)), CATEGORICAL_VARS))
TABLE_NOTE <- paste0(
  "Normally distributed variables: mean \u00b1 SD (SE); range. Skewed variables ",
  "(Shapiro\u2013Wilk p < .05 or |skewness| > 1 in the full sample): median [Q1 \u2013 Q3]; range. ",
  "Abbreviations: BMI, body mass index; FOG, freezing of gait; LEDD, levodopa-equivalent daily dose; ",
  "MDS-UPDRS, Movement Disorder Society Unified Parkinson's Disease Rating Scale; MoCA, Montreal ",
  "Cognitive Assessment; NFOG-Q, New Freezing of Gait Questionnaire; STAI, State-Trait Anxiety ",
  "Inventory; TMT, Trail Making Test."
)

suppressMessages(theme_gtsummary_journal(journal = "jama", set_theme = TRUE))

# Variables with no observed values in a sample are left out of that table
# (gtsummary cannot summarise them) and reported
build_table1 <- function(data, sample_name) {
  empty <- VAR_NAMES[vapply(data[VAR_NAMES], function(x) all(is.na(x)), logical(1))]
  if (length(empty)) {
    log_warn(sprintf("Table 1 (%s): no data for %s - omitted from the table",
                     sample_name, paste(empty, collapse = ", ")))
  }
  vars <- setdiff(VAR_NAMES, empty)

  stat_list <- c(
    if (length(intersect(NORMAL_VARS, vars))) list(any_of(NORMAL_VARS) ~ STAT_NORMAL),
    if (length(intersect(SKEWED_VARS, vars))) list(any_of(SKEWED_VARS) ~ STAT_SKEWED),
    list(all_categorical() ~ "{n} ({p}%)")
  )

  data %>%
    select(all_of(vars)) %>%
    tbl_summary(
      type      = list(any_of(CONT_VARS) ~ "continuous"),
      statistic = stat_list,
      digits    = list(all_continuous() ~ 1, all_categorical() ~ c(0, 1)),
      label     = as.list(VAR_LABELS[vars]),
      missing   = "ifany",
      missing_text = "Missing, n"
    ) %>%
    modify_table_body(~ .x %>%
      mutate(stat_label = ifelse(row_type == "label", STAT_DESCRIPTOR[variable], stat_label))) %>%
    modify_header(label = "**Characteristic**", stat_0 = "**N = {N}**") %>%
    bold_labels()
}

excl_label <- paste(EXCLUDE_IDS, collapse = ", ")
base_all  <- build_table1(df_demo, "full sample")
base_excl <- build_table1(df_excl, paste("excluding", excl_label))

table_1_all <- base_all %>%
  modify_caption("**Table 1. Demographic and clinical characteristics (full sample)**") %>%
  modify_source_note(TABLE_NOTE)

table_1_filtered <- base_excl %>%
  modify_caption(sprintf("**Table 1. Demographic and clinical characteristics (excluding %s)**", excl_label)) %>%
  modify_source_note(TABLE_NOTE)

# The excluded sample is a subset of the full sample, so its rows are a subset
# of the full table's and merge in order (row-mismatch notice silenced)
table_1_combined <- tbl_merge(
  list(base_all, base_excl),
  tab_spanner = c("**Full sample**", sprintf("**Excluding %s**", excl_label)),
  quiet = TRUE
) %>%
  modify_caption("**Table 1. Demographic and clinical characteristics**") %>%
  modify_source_note(TABLE_NOTE)

# Terminal view: strip markdown emphasis and indent level rows under their label
print_table1 <- function(tbl) {
  tb <- as_tibble(tbl, col_labels = TRUE)
  names(tb) <- gsub("\\*\\*", "", names(tb))
  tb <- tb %>% mutate(across(everything(), ~ gsub("__", "", .x)))
  is_label <- tbl$table_body$row_type == "label"
  tb[[1]] <- ifelse(is_label, tb[[1]], paste0("    ", tb[[1]]))
  print_aligned(tb)
}

banner("TABLE 1: FULL SAMPLE")
print_table1(table_1_all)

banner(sprintf("TABLE 1: EXCLUDING %s", excl_label))
print_table1(table_1_filtered)

# Export: Word + PNG via flextable, HTML via gt
banner("EXPORTING TABLE 1")
export_table <- function(tbl, stem) {
  ft <- as_flex_table(tbl) %>%
    flextable::font(fontname = "Arial", part = "all") %>%
    flextable::fontsize(size = 9, part = "all") %>%
    flextable::fontsize(size = 8, part = "footer") %>%
    flextable::padding(padding.top = 2, padding.bottom = 2, part = "body") %>%
    flextable::autofit()
  landscape <- prop_section(page_size = page_size(orient = "landscape"),
                            page_margins = page_mar(top = 0.6, bottom = 0.6, left = 0.6, right = 0.6))
  # PNG is rendered first: rendering it after save_as_docx() has crashed R in
  # some flextable/officer builds
  outputs <- c(png  = file.path(RESULTS_DIR, paste0(stem, ".png")),
               html = file.path(RESULTS_DIR, paste0(stem, ".html")),
               docx = file.path(RESULTS_DIR, paste0(stem, ".docx")))
  writers <- list(
    png  = function(p) save_as_image(ft, path = p, res = 300),
    html = function(p) gtsave(as_gt(tbl), filename = p),
    docx = function(p) save_as_docx(ft, path = p, pr_section = landscape)
  )
  for (fmt in names(outputs)) {
    tryCatch({
      writers[[fmt]](outputs[[fmt]])
      cat(sprintf("  Saved %s\n", outputs[[fmt]]))
    }, error = function(e) log_warn(sprintf("could not write %s: %s", outputs[[fmt]], conditionMessage(e))))
  }
}
export_table(table_1_all,      "Table1_Full_Sample")
export_table(table_1_filtered, paste0("Table1_Excluding_", paste(EXCLUDE_IDS, collapse = "_")))
export_table(table_1_combined, "Table1_Combined")

# 6. Missing Data & Quality Audit ----------------------------------------------
banner("MISSING DATA AUDIT BY VARIABLE (FULL SAMPLE)")

source_of <- unlist(lapply(names(SOURCES), function(nm) setNames(rep(nm, length(SOURCES[[nm]]$cols)),
                                                               names(SOURCES[[nm]]$cols))))
n_total <- nrow(df_demo)
missing_by_var <- map(VAR_NAMES, function(v) {
  miss <- is.na(df_demo[[v]])
  tibble(Variable = v,
         Source = source_of[[v]],
         `N obs` = sum(!miss),
         Missing = sum(miss),
         `% Missing` = sprintf("%.1f%%", 100 * mean(miss)),
         `Missing participant IDs` = if (any(miss)) paste(df_demo$Participant_ID[miss], collapse = ", ") else "-")
}) %>% list_rbind()

print_aligned(missing_by_var, wrap = "Missing participant IDs", wrap_width = 50,
              right = c("N obs", "Missing", "% Missing"))
cat(sprintf("\n  Overall: %d of %d cells missing (%.1f%%) across %d participants x %d variables\n",
            sum(missing_by_var$Missing), n_total * length(VAR_NAMES),
            100 * sum(missing_by_var$Missing) / (n_total * length(VAR_NAMES)),
            n_total, length(VAR_NAMES)))

banner("PARTICIPANT-LEVEL COMPLETENESS")
completeness <- df_demo %>%
  pivot_longer(all_of(VAR_NAMES), names_to = "Variable", values_to = "Value",
               values_transform = as.character) %>%
  group_by(Participant_ID) %>%
  summarise(Fields = n(),
            Complete = sum(!is.na(Value)),
            Pct = 100 * Complete / Fields,
            `Missing fields` = if (any(is.na(Value))) paste(Variable[is.na(Value)], collapse = ", ") else "-",
            .groups = "drop") %>%
  mutate(Excluded = ifelse(Participant_ID %in% EXCLUDE_IDS, "yes", "")) %>%
  arrange(Pct, Participant_ID)

print_aligned(completeness %>%
                transmute(Participant = Participant_ID, Fields, Complete,
                          `% Complete` = sprintf("%.1f%%", Pct), Excluded, `Missing fields`),
              wrap = "Missing fields", wrap_width = 55,
              right = c("Fields", "Complete", "% Complete"))
cat(sprintf("\n  Fully complete: %d/%d participants; median completeness %.1f%% (range %.1f%% - %.1f%%)\n",
            sum(completeness$Pct == 100), nrow(completeness), median(completeness$Pct),
            min(completeness$Pct), max(completeness$Pct)))

write_csv(missing_by_var, file.path(RESULTS_DIR, "Missing_Data_By_Variable.csv"))
write_csv(completeness %>% rename(Pct_Complete = Pct), file.path(RESULTS_DIR, "Participant_Completeness.csv"))
cat(sprintf("  Saved audit CSVs to %s\n", RESULTS_DIR))

# 7. Multi-Panel Clinical Distributions ----------------------------------------
banner("PLOTTING CLINICAL DISTRIBUTIONS")

plot_vars <- c(
  Age              = "Age\n(years)",
  Disease_Duration = "Disease Duration\n(years)",
  LEDD             = "Total LEDD\n(mg/day)",
  MoCA             = "MoCA Score\n(/30)",
  TMT_A            = "TMT-A\n(seconds)",
  TMT_B            = "TMT-B\n(seconds)",
  dTMT             = "TMT B\u2212A\n(seconds)",
  UPDRS_OFF        = "MDS-UPDRS III\n(OFF Med)",
  UPDRS_ON         = "MDS-UPDRS III\n(ON Med)",
  NFOG_Q           = "NFOG-Q\n(Total)",
  STAI_Trait       = "STAI-Trait\n(Total)"
)

# Reference bands (display aids, not diagnostic thresholds):
# - MoCA: normal >= 26 (Nasreddine et al., 2005); mild 18-25, moderate 10-17,
#   severe < 10 (commonly used severity ranges)
# - MDS-UPDRS III: mild <= 32, moderate 33-58, severe >= 59 (Martinez-Martin
#   et al., 2015); no "normal" band is defined for a PD cohort
# - TMT: severe = Reitan "deficient" cut-offs (A > 78 s, B > 273 s); normal
#   A < 40 s, B < 92 s; mild/moderate split at the midpoint. Prefer age- and
#   education-stratified norms (e.g. Tombaugh, 2004) for inference.
clinical_bands <- tribble(
  ~Var_Key,     ~ymin, ~ymax, ~Status,
  "MoCA",        25.5,   Inf, "Normal",
  "MoCA",        17.5,  25.5, "Mild",
  "MoCA",         9.5,  17.5, "Moderate",
  "MoCA",        -Inf,   9.5, "Severe",
  "TMT_A",       -Inf,    40, "Normal",
  "TMT_A",         40,    59, "Mild",
  "TMT_A",         59,    78, "Moderate",
  "TMT_A",         78,   Inf, "Severe",
  "TMT_B",       -Inf,  91.5, "Normal",
  "TMT_B",       91.5, 182.5, "Mild",
  "TMT_B",      182.5,   273, "Moderate",
  "TMT_B",        273,   Inf, "Severe",
  "UPDRS_OFF",   -Inf,  32.5, "Mild",
  "UPDRS_OFF",   32.5,  58.5, "Moderate",
  "UPDRS_OFF",   58.5,   Inf, "Severe",
  "UPDRS_ON",    -Inf,  32.5, "Mild",
  "UPDRS_ON",    32.5,  58.5, "Moderate",
  "UPDRS_ON",    58.5,   Inf, "Severe"
)

# Colours are assigned over the full roster so each participant keeps the same
# colour in every figure, whichever subset is plotted
pids <- sort(df_demo$Participant_ID)
paired <- brewer.pal(12, "Paired")
pid_cols <- setNames(if (length(pids) <= 12) paired[seq_along(pids)]
                     else colorRampPalette(paired)(length(pids)), pids)

df_plot_long <- df_excl %>%
  select(Participant_ID, all_of(names(plot_vars))) %>%
  pivot_longer(-Participant_ID, names_to = "Var_Key", values_to = "Value") %>%
  filter(!is.na(Value))

facet_n <- count(df_plot_long, Var_Key)
facet_labels <- setNames(
  sprintf("%s  [n = %d]", plot_vars, coalesce(facet_n$n[match(names(plot_vars), facet_n$Var_Key)], 0L)),
  names(plot_vars)
)
df_plot_long <- df_plot_long %>%
  mutate(Variable = factor(facet_labels[Var_Key], levels = facet_labels))

# Clip bands to each facet's data range so they never stretch the free y-scales
facet_range <- df_plot_long %>%
  group_by(Var_Key) %>%
  summarise(lo = min(Value), hi = max(Value), .groups = "drop") %>%
  mutate(pad = pmax((hi - lo) * 0.08, 1), lo = lo - pad, hi = hi + pad)

bands_plot <- clinical_bands %>%
  inner_join(facet_range, by = "Var_Key") %>%
  mutate(ymin = pmax(ymin, lo), ymax = pmin(ymax, hi)) %>%
  filter(ymax > ymin) %>%
  mutate(Variable = factor(facet_labels[Var_Key], levels = facet_labels),
         Status = factor(Status, levels = c("Normal", "Mild", "Moderate", "Severe")))

p_demo <- ggplot(df_plot_long, aes(x = "", y = Value)) +
  geom_rect(data = bands_plot,
            aes(xmin = -Inf, xmax = Inf, ymin = ymin, ymax = ymax, fill = Status),
            inherit.aes = FALSE, alpha = 0.30) +
  geom_boxplot(outlier.shape = NA, fill = "white", colour = "grey30",
               width = 0.5, linewidth = 0.5, alpha = 0.7) +
  geom_point(aes(colour = Participant_ID),
             position = position_jitter(width = 0.15, height = 0, seed = 2024),
             size = 2.4, alpha = 0.9) +
  stat_summary(fun = mean, geom = "point", shape = 23, size = 3, fill = "white") +
  facet_wrap(~ Variable, scales = "free_y", ncol = 4) +
  scale_colour_manual(values = pid_cols, name = "Participant", drop = TRUE) +
  scale_fill_manual(values = c(Normal = "#66bd63", Mild = "#fee08b",
                               Moderate = "#fdae61", Severe = "#d73027"),
                    name = "Clinical range", drop = FALSE) +
  guides(fill = guide_legend(order = 1, override.aes = list(alpha = 0.5)),
         colour = guide_legend(order = 2, ncol = 2, override.aes = list(size = 3))) +
  labs(title = sprintf("Clinical & Demographic Cohort Distributions (Excluding %s)", excl_label),
       subtitle = paste("Box = median & IQR; white diamond = mean; points = individual participants.",
                        "Shaded bands = clinical reference ranges (MoCA, TMT, MDS-UPDRS III)."),
       x = NULL, y = "Score / Units") +
  theme_bw(base_size = 11) +
  theme(
    plot.title = element_text(face = "bold", size = 14, hjust = 0.5),
    plot.subtitle = element_text(size = 9.5, hjust = 0.5, colour = "grey30"),
    axis.text.x = element_blank(),
    axis.ticks.x = element_blank(),
    strip.text = element_text(face = "bold", size = 8.5),
    strip.background = element_rect(fill = "grey95"),
    panel.grid.major.x = element_blank(),
    legend.key.size = unit(0.4, "cm"),
    legend.text = element_text(size = 8)
  )

plot_path <- file.path(RESULTS_DIR, "Demographics_Distributions.png")
ggsave(plot_path, plot = p_demo, width = 12, height = 8, dpi = 300, bg = "white")
cat(sprintf("  Saved %s\n", plot_path))

# 8. Session Summary -----------------------------------------------------------
banner("RUN COMPLETE")
cat(sprintf("  Participants: %d (full) / %d (excluding %s)\n", nrow(df_demo), nrow(df_excl), excl_label))
cat(sprintf("  Data-quality warnings raised: %d (see [WARN] lines above)\n", n_warnings))
cat(sprintf("  Outputs written to: %s\n", RESULTS_DIR))
cat(sprintf("  R %s | gtsummary %s | ggplot2 %s | elapsed %.1f s\n",
            getRversion(), packageVersion("gtsummary"), packageVersion("ggplot2"),
            as.numeric(difftime(Sys.time(), t_start, units = "secs"))))
