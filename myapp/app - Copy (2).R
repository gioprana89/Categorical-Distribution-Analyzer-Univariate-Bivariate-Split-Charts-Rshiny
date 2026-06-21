# ============================================================
# STATCAL ONLINE - Categorical Frequency Distribution Analyzer
# R Shiny Version
# ============================================================
# Required packages:
# install.packages(c(
#   "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
#   "ggplot2", "shinycssloaders", "scales", "openxlsx"
# ))

required_packages <- c(
  "shiny", "shinydashboard", "DT", "readxl", "dplyr", "tidyr",
  "ggplot2", "shinycssloaders", "scales", "openxlsx"
)

missing_packages <- required_packages[!vapply(required_packages, requireNamespace, logical(1), quietly = TRUE)]
if (length(missing_packages) > 0) {
  stop(
    "Please install the following R packages first: ",
    paste(missing_packages, collapse = ", "),
    call. = FALSE
  )
}

library(shiny)
library(shinydashboard)
library(DT)
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(shinycssloaders)
library(scales)
library(openxlsx)

# ============================================================
# CONSTANTS
# ============================================================

APP_NAME <- "STATCAL ONLINE"
APP_TITLE <- "Categorical Frequency Distribution and Bivariate Bar Chart Analyzer"
APP_UPDATED <- "Last updated on June 21, 2026"
WEBSITE_URL <- "https://statcal.com/"
URL_DATA_TRAINING <- "https://drive.google.com/drive/folders/1x623-vbvaCj7-YA-HljZXTPU7LSdgJmW?usp=sharing"
STATCAL_ONLINE_URL <- "https://statcal.com/statcal%20online.html"
SAMPLE_DATA_PATH <- "data variabel kategori.xlsx"

THEMES <- list(
  "White Publication" = list(
    figure_facecolor = "white", axes_facecolor = "white", text_color = "#111111",
    grid_color = "#D9D9D9", spine_color = "#222222"
  ),
  "Light Gray Editorial" = list(
    figure_facecolor = "#F7F7F7", axes_facecolor = "#FFFFFF", text_color = "#111111",
    grid_color = "#D0D0D0", spine_color = "#333333"
  ),
  "Warm Ivory Journal" = list(
    figure_facecolor = "#FBF7EF", axes_facecolor = "#FFFDF8", text_color = "#1F1F1F",
    grid_color = "#DDD4C4", spine_color = "#3A3A3A"
  ),
  "Cool Blue Scientific" = list(
    figure_facecolor = "#F3F7FB", axes_facecolor = "#FFFFFF", text_color = "#0B1F33",
    grid_color = "#C8D6E5", spine_color = "#1F4E79"
  ),
  "Dark Navy Presentation" = list(
    figure_facecolor = "#0B1320", axes_facecolor = "#111C2E", text_color = "#FFFFFF",
    grid_color = "#3B4A5F", spine_color = "#B8C7D9"
  ),
  "Minimal Scopus Style" = list(
    figure_facecolor = "#FFFFFF", axes_facecolor = "#FFFFFF", text_color = "#111111",
    grid_color = "#EAEAEA", spine_color = "#111111"
  ),
  "Soft Blue Journal" = list(
    figure_facecolor = "#F4F8FC", axes_facecolor = "#FFFFFF", text_color = "#102A43",
    grid_color = "#D8E6F2", spine_color = "#243B53"
  )
)

DEFAULT_COLORS <- c(
  "#1F4E79", "#E97132", "#70AD47", "#FFC000", "#7030A0",
  "#00A6A6", "#C00000", "#595959", "#2166AC", "#B2182B",
  "#4D9221", "#762A83", "#D6604D", "#4393C3", "#F4A582"
)

# ============================================================
# HELPER FUNCTIONS
# ============================================================

clean_dataframe <- function(df) {
  names(df) <- trimws(gsub("\\s+", " ", as.character(names(df))))
  df <- df[rowSums(is.na(df)) < ncol(df), , drop = FALSE]
  unnamed_cols <- grepl("^unnamed", tolower(names(df)))
  if (any(unnamed_cols)) {
    keep_unnamed <- vapply(df[unnamed_cols], function(x) !all(is.na(x)), logical(1))
    drop_names <- names(df)[unnamed_cols][!keep_unnamed]
    if (length(drop_names) > 0) df <- df[, !names(df) %in% drop_names, drop = FALSE]
  }
  rownames(df) <- NULL
  as.data.frame(df)
}

make_display_safe <- function(df) {
  df <- as.data.frame(df)
  for (nm in names(df)) {
    if (is.factor(df[[nm]])) df[[nm]] <- as.character(df[[nm]])
  }
  df
}

clean_category_value <- function(x) {
  x <- trimws(as.character(x))
  x[x %in% c("", "NA", "NaN", "nan", "NULL", "None", "NaT", "<NA>")] <- NA_character_
  x
}

sorted_unique_values <- function(x) {
  vals <- unique(clean_category_value(x))
  vals <- vals[!is.na(vals)]
  vals[order(as.character(vals))]
}

safe_id <- function(x) {
  x <- gsub("[^A-Za-z0-9_]", "_", as.character(x))
  x <- gsub("_+", "_", x)
  x <- gsub("^_|_$", "", x)
  if (!nzchar(x)) x <- "value"
  x
}

safe_number <- function(x, default_value, min_value = NULL, max_value = NULL) {
  out <- suppressWarnings(as.numeric(x))
  if (length(out) == 0 || is.na(out) || !is.finite(out)) out <- default_value
  if (!is.null(min_value)) out <- max(out, min_value)
  if (!is.null(max_value)) out <- min(out, max_value)
  out
}

parse_order_text <- function(x) {
  if (is.null(x) || length(x) == 0 || is.na(x) || !nzchar(trimws(as.character(x)))) return(character(0))
  out <- trimws(unlist(strsplit(as.character(x), ",", fixed = TRUE)))
  out <- out[nzchar(out)]
  unique(out)
}

resolve_category_order <- function(values, order_text = NULL) {
  values <- clean_category_value(values)
  available <- unique(values[!is.na(values)])
  available <- available[order(as.character(available))]
  wanted <- parse_order_text(order_text)
  if (length(wanted) == 0) return(available)
  ordered <- wanted[wanted %in% available]
  remaining <- setdiff(available, ordered)
  unique(c(ordered, remaining))
}

is_hex_color <- function(x) {
  is.character(x) && length(x) == 1 && grepl("^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$", x)
}

get_manual_colors <- function(levels, input, prefix) {
  out <- character(length(levels))
  for (i in seq_along(levels)) {
    default_col <- DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1]
    input_id <- paste0(prefix, safe_id(levels[i]))
    val <- input[[input_id]]
    if (is.null(val) || length(val) == 0 || !is_hex_color(val)) val <- default_col
    out[i] <- val
  }
  names(out) <- levels
  out
}

get_theme <- function(theme_name) {
  if (is.null(theme_name) || length(theme_name) == 0 || is.na(theme_name) || !(theme_name %in% names(THEMES))) {
    return(THEMES[["White Publication"]])
  }
  THEMES[[theme_name]]
}

safe_theme_bg <- function(theme_name) {
  tryCatch(get_theme(theme_name)$figure_facecolor, error = function(e) "white")
}

legend_position_value <- function(position_label) {
  if (is.null(position_label) || position_label == "None / Hide legend") return("none")
  tolower(position_label)
}

statcal_theme_gg <- function(theme_name,
                             title_size = 16,
                             subtitle_size = 11,
                             axis_title_size = 11,
                             axis_text_size = 9,
                             panel_title_size = 11,
                             legend_title_size = 10,
                             legend_text_size = 9,
                             legend_position = "Right",
                             x_text_angle = 35) {
  th <- get_theme(theme_name)
  theme_minimal(base_size = axis_text_size) +
    theme(
      plot.background = element_rect(fill = th$figure_facecolor, color = NA),
      panel.background = element_rect(fill = th$axes_facecolor, color = NA),
      panel.grid.major = element_line(color = th$grid_color, linewidth = 0.35),
      panel.grid.minor = element_line(color = th$grid_color, linewidth = 0.15),
      axis.text = element_text(color = th$text_color, size = axis_text_size),
      axis.text.x = element_text(angle = x_text_angle, hjust = ifelse(x_text_angle == 0, 0.5, 1), color = th$text_color, size = axis_text_size),
      axis.title = element_text(color = th$text_color, face = "bold", size = axis_title_size),
      plot.title = element_text(color = th$text_color, face = "bold", size = title_size),
      plot.subtitle = element_text(color = th$text_color, size = subtitle_size),
      strip.text = element_text(color = th$text_color, face = "bold", size = panel_title_size),
      legend.text = element_text(color = th$text_color, size = legend_text_size),
      legend.title = element_text(color = th$text_color, face = "bold", size = legend_title_size),
      legend.background = element_rect(fill = th$figure_facecolor, color = NA),
      legend.position = legend_position_value(legend_position)
    )
}

# ============================================================
# FREQUENCY AND PERCENTAGE TABLES
# ============================================================

compute_univariate_distribution <- function(df, variables, digits = 2, category_order = NULL) {
  validate(need(length(variables) > 0, "Please select at least one categorical variable."))
  rows <- list()
  for (var in variables) {
    vals <- clean_category_value(df[[var]])
    cats <- resolve_category_order(vals, category_order)
    vals <- vals[!is.na(vals)]
    total <- length(vals)
    if (length(cats) == 0) cats <- character(0)
    counts <- vapply(cats, function(cat) sum(vals == cat, na.rm = TRUE), numeric(1))
    pcts <- if (total > 0) counts / total * 100 else rep(NA_real_, length(cats))
    tmp <- data.frame(
      Variable = var,
      Category = cats,
      Frequency = as.numeric(counts),
      Percentage = round(as.numeric(pcts), digits),
      Total = as.numeric(total),
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
    rows[[length(rows) + 1]] <- tmp
  }
  dplyr::bind_rows(rows)
}

compute_bivariate_distribution <- function(df, dependent_var, independent_vars, digits = 2,
                                           dependent_order = NULL, independent_order = NULL) {
  validate(need(!is.null(dependent_var) && nzchar(dependent_var), "Please select one dependent variable."))
  validate(need(length(independent_vars) > 0, "Please select at least one independent variable."))
  dep_values_all <- clean_category_value(df[[dependent_var]])
  dep_categories <- resolve_category_order(dep_values_all, dependent_order)
  rows <- list()
  for (indep in independent_vars) {
    indep_values_all <- clean_category_value(df[[indep]])
    indep_categories <- resolve_category_order(indep_values_all, independent_order)
    for (icat in indep_categories) {
      keep <- !is.na(indep_values_all) & indep_values_all == icat & !is.na(dep_values_all)
      dep_sub <- dep_values_all[keep]
      total <- length(dep_sub)
      counts <- vapply(dep_categories, function(dcat) sum(dep_sub == dcat, na.rm = TRUE), numeric(1))
      pcts <- if (total > 0) counts / total * 100 else rep(NA_real_, length(dep_categories))
      tmp <- data.frame(
        `Independent Variable` = indep,
        `Independent Category` = icat,
        `Dependent Variable` = dependent_var,
        `Dependent Category` = dep_categories,
        Frequency = as.numeric(counts),
        Percentage = round(as.numeric(pcts), digits),
        Total = as.numeric(total),
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
      rows[[length(rows) + 1]] <- tmp
    }
  }
  dplyr::bind_rows(rows)
}

make_bivariate_wide_table <- function(bivar_df) {
  if (nrow(bivar_df) == 0) return(bivar_df)
  freq_df <- bivar_df[, c("Independent Variable", "Independent Category", "Dependent Category", "Frequency"), drop = FALSE]
  pct_df <- bivar_df[, c("Independent Variable", "Independent Category", "Dependent Category", "Percentage"), drop = FALSE]
  names(freq_df)[4] <- "Value"
  names(pct_df)[4] <- "Value"
  freq_df$Measure <- "f"
  pct_df$Measure <- "%"
  out <- dplyr::bind_rows(freq_df, pct_df)
  out$Column <- paste(out$`Dependent Category`, out$Measure)
  tidyr::pivot_wider(
    out[, c("Independent Variable", "Independent Category", "Column", "Value"), drop = FALSE],
    names_from = "Column",
    values_from = "Value",
    values_fill = 0
  )
}

# ============================================================
# CHART FUNCTIONS
# ============================================================

prepare_univariate_chart_data <- function(univ_df, category_order = NULL) {
  if (nrow(univ_df) == 0) return(univ_df)
  order_levels <- parse_order_text(category_order)
  all_cats <- unique(as.character(univ_df$Category))
  if (length(order_levels) > 0) {
    category_levels <- unique(c(order_levels[order_levels %in% all_cats], setdiff(all_cats, order_levels)))
  } else {
    category_levels <- all_cats
  }
  univ_df$Category <- factor(as.character(univ_df$Category), levels = category_levels)
  univ_df$Variable <- factor(as.character(univ_df$Variable), levels = unique(as.character(univ_df$Variable)))
  univ_df
}

create_univariate_bar_plot <- function(chart_df, manual_colors,
                                       title, subtitle,
                                       y_metric = "Percentage",
                                       label_mode = "Percentage",
                                       orientation = "Vertical",
                                       theme_name = "White Publication",
                                       bar_width = 0.72,
                                       show_labels = TRUE,
                                       label_size = 3.2,
                                       label_color = "#111111",
                                       panel_cols = 2,
                                       facet_scales = "free_x",
                                       legend_position = "Right",
                                       title_size = 16,
                                       subtitle_size = 11,
                                       axis_title_size = 11,
                                       axis_text_size = 9,
                                       panel_title_size = 11,
                                       legend_title_size = 10,
                                       legend_text_size = 9,
                                       x_text_angle = 35,
                                       legend_title = "Category",
                                       x_axis_title = "Category",
                                       digits = 2) {
  validate(need(nrow(chart_df) > 0, "Univariate chart data is empty."))
  th <- get_theme(theme_name)
  panel_cols <- safe_number(panel_cols, 2, 1, 10)
  if (is.null(legend_title) || !nzchar(trimws(as.character(legend_title)))) legend_title <- "Category"
  if (is.null(x_axis_title) || !nzchar(trimws(as.character(x_axis_title)))) x_axis_title <- "Category"
  chart_df$Y_Value <- if (y_metric == "Frequency") chart_df$Frequency else chart_df$Percentage
  chart_df$Label <- ""
  if (label_mode == "Frequency") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, as.character(chart_df$Frequency), "")
  } else if (label_mode == "Percentage") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, paste0(format(round(chart_df$Percentage, digits), nsmall = digits), "%"), "")
  } else if (label_mode == "Frequency and Percentage") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, paste0(chart_df$Frequency, " (", format(round(chart_df$Percentage, digits), nsmall = digits), "%)"), "")
  }
  y_label <- if (y_metric == "Frequency") "Frequency" else "Percentage (%)"
  facet_scales <- ifelse(is.null(facet_scales) || !(facet_scales %in% c("fixed", "free_y", "free_x", "free")), "free_x", facet_scales)
  p <- ggplot(chart_df, aes(x = Category, y = Y_Value, fill = Category)) +
    geom_col(width = bar_width, color = "white", linewidth = 0.25) +
    scale_fill_manual(values = manual_colors, drop = FALSE) +
    facet_wrap(~ Variable, ncol = panel_cols, scales = facet_scales) +
    labs(title = title, subtitle = subtitle, x = x_axis_title, y = y_label, fill = legend_title) +
    statcal_theme_gg(
      theme_name = theme_name,
      title_size = title_size,
      subtitle_size = subtitle_size,
      axis_title_size = axis_title_size,
      axis_text_size = axis_text_size,
      panel_title_size = panel_title_size,
      legend_title_size = legend_title_size,
      legend_text_size = legend_text_size,
      legend_position = legend_position,
      x_text_angle = x_text_angle
    )
  if (show_labels && label_mode != "None") {
    if (orientation == "Horizontal") {
      p <- p + geom_text(aes(label = Label), hjust = -0.08, size = label_size, color = label_color, check_overlap = TRUE)
    } else {
      p <- p + geom_text(aes(label = Label), vjust = -0.35, size = label_size, color = label_color, check_overlap = TRUE)
    }
  }
  if (orientation == "Horizontal") {
    p <- p + coord_flip()
  }
  p + theme(plot.background = element_rect(fill = th$figure_facecolor, color = NA))
}

prepare_bivariate_chart_data <- function(bivar_df, dependent_order = NULL, independent_order = NULL) {
  if (nrow(bivar_df) == 0) return(bivar_df)
  dep_levels <- parse_order_text(dependent_order)
  dep_all <- unique(as.character(bivar_df$`Dependent Category`))
  if (length(dep_levels) > 0) {
    dep_levels <- unique(c(dep_levels[dep_levels %in% dep_all], setdiff(dep_all, dep_levels)))
  } else {
    dep_levels <- dep_all
  }
  indep_levels <- parse_order_text(independent_order)
  indep_all <- unique(as.character(bivar_df$`Independent Category`))
  if (length(indep_levels) > 0) {
    indep_levels <- unique(c(indep_levels[indep_levels %in% indep_all], setdiff(indep_all, indep_levels)))
  } else {
    indep_levels <- indep_all
  }
  bivar_df$`Dependent Category` <- factor(as.character(bivar_df$`Dependent Category`), levels = dep_levels)
  bivar_df$`Independent Category` <- factor(as.character(bivar_df$`Independent Category`), levels = indep_levels)
  bivar_df$`Independent Variable` <- factor(as.character(bivar_df$`Independent Variable`), levels = unique(as.character(bivar_df$`Independent Variable`)))
  bivar_df
}

create_bivariate_bar_plot <- function(chart_df, manual_colors,
                                      title, subtitle,
                                      y_metric = "Percentage",
                                      label_mode = "Percentage",
                                      bar_position = "Stacked",
                                      orientation = "Vertical",
                                      theme_name = "White Publication",
                                      bar_width = 0.72,
                                      show_labels = TRUE,
                                      label_size = 3.0,
                                      label_color = "#111111",
                                      label_vjust = 0.5,
                                      panel_cols = 2,
                                      facet_scales = "free_x",
                                      legend_position = "Right",
                                      title_size = 16,
                                      subtitle_size = 11,
                                      axis_title_size = 11,
                                      axis_text_size = 9,
                                      panel_title_size = 11,
                                      legend_title_size = 10,
                                      legend_text_size = 9,
                                      x_text_angle = 35,
                                      legend_title = "Dependent Category",
                                      x_axis_title = "Independent Category",
                                      digits = 2) {
  validate(need(nrow(chart_df) > 0, "Bivariate chart data is empty."))
  th <- get_theme(theme_name)
  panel_cols <- safe_number(panel_cols, 2, 1, 10)
  if (is.null(legend_title) || !nzchar(trimws(as.character(legend_title)))) legend_title <- "Dependent Category"
  if (is.null(x_axis_title) || !nzchar(trimws(as.character(x_axis_title)))) x_axis_title <- "Independent Category"
  chart_df$Y_Value <- if (y_metric == "Frequency") chart_df$Frequency else chart_df$Percentage
  chart_df$Label <- ""
  if (label_mode == "Frequency") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, as.character(chart_df$Frequency), "")
  } else if (label_mode == "Percentage") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, paste0(format(round(chart_df$Percentage, digits), nsmall = digits), "%"), "")
  } else if (label_mode == "Frequency and Percentage") {
    chart_df$Label <- ifelse(chart_df$Frequency > 0, paste0(chart_df$Frequency, " (", format(round(chart_df$Percentage, digits), nsmall = digits), "%)"), "")
  }
  y_label <- if (y_metric == "Frequency") "Frequency" else "Percentage (%)"
  facet_scales <- ifelse(is.null(facet_scales) || !(facet_scales %in% c("fixed", "free_y", "free_x", "free")), "free_x", facet_scales)
  position_obj <- if (bar_position == "Grouped") position_dodge(width = 0.78) else "stack"
  p <- ggplot(chart_df, aes(x = `Independent Category`, y = Y_Value, fill = `Dependent Category`)) +
    geom_col(width = bar_width, position = position_obj, color = "white", linewidth = 0.25) +
    scale_fill_manual(values = manual_colors, drop = FALSE) +
    facet_wrap(~ `Independent Variable`, ncol = panel_cols, scales = facet_scales) +
    labs(title = title, subtitle = subtitle, x = x_axis_title, y = y_label, fill = legend_title) +
    statcal_theme_gg(
      theme_name = theme_name,
      title_size = title_size,
      subtitle_size = subtitle_size,
      axis_title_size = axis_title_size,
      axis_text_size = axis_text_size,
      panel_title_size = panel_title_size,
      legend_title_size = legend_title_size,
      legend_text_size = legend_text_size,
      legend_position = legend_position,
      x_text_angle = x_text_angle
    )
  if (show_labels && label_mode != "None") {
    if (bar_position == "Grouped") {
      p <- p + geom_text(aes(label = Label), position = position_dodge(width = 0.78), vjust = -0.35, size = label_size, color = label_color, check_overlap = TRUE)
    } else {
      p <- p + geom_text(aes(label = Label), position = position_stack(vjust = label_vjust), size = label_size, color = label_color, check_overlap = TRUE)
    }
  }
  if (orientation == "Horizontal") {
    p <- p + coord_flip()
  }
  p + theme(plot.background = element_rect(fill = th$figure_facecolor, color = NA))
}


# ============================================================
# CATEGORICAL ASSOCIATION AND SIGNIFICANCE TESTS
# ============================================================

make_contingency_table <- function(df, row_var, col_var, row_order = NULL, col_order = NULL) {
  validate(need(!is.null(row_var) && row_var %in% names(df), "Please select a valid independent categorical variable."))
  validate(need(!is.null(col_var) && col_var %in% names(df), "Please select a valid dependent categorical variable."))
  row_values <- clean_category_value(df[[row_var]])
  col_values <- clean_category_value(df[[col_var]])
  keep <- !is.na(row_values) & !is.na(col_values)
  row_values <- row_values[keep]
  col_values <- col_values[keep]
  row_levels <- resolve_category_order(row_values, row_order)
  col_levels <- resolve_category_order(col_values, col_order)
  tab <- table(
    `Independent Category` = factor(row_values, levels = row_levels),
    `Dependent Category` = factor(col_values, levels = col_levels),
    useNA = "no"
  )
  tab
}

as_table_df <- function(tab, first_col_name = "Independent Category") {
  df <- as.data.frame.matrix(tab, stringsAsFactors = FALSE)
  df <- cbind(setNames(data.frame(rownames(df), stringsAsFactors = FALSE), first_col_name), df)
  rownames(df) <- NULL
  df
}

cramers_v_bias_corrected <- function(tab, chi_sq) {
  n <- sum(tab)
  r <- nrow(tab)
  k <- ncol(tab)
  if (n <= 1 || r < 2 || k < 2) return(NA_real_)
  phi2 <- chi_sq / n
  phi2_corr <- max(0, phi2 - ((k - 1) * (r - 1)) / (n - 1))
  r_corr <- r - ((r - 1)^2) / (n - 1)
  k_corr <- k - ((k - 1)^2) / (n - 1)
  denom <- min(k_corr - 1, r_corr - 1)
  if (!is.finite(denom) || denom <= 0) return(NA_real_)
  sqrt(phi2_corr / denom)
}

compute_nominal_association_measures <- function(tab, chisq_p_value = NA_real_, digits = 4) {
  n <- sum(tab)
  r <- nrow(tab)
  k <- ncol(tab)
  chi_obj <- suppressWarnings(chisq.test(tab, correct = FALSE))
  chi_sq <- unname(chi_obj$statistic)
  phi <- if (n > 0) sqrt(chi_sq / n) else NA_real_
  cramer <- if (n > 0 && min(r - 1, k - 1) > 0) sqrt(chi_sq / (n * min(r - 1, k - 1))) else NA_real_
  cramer_corr <- cramers_v_bias_corrected(tab, chi_sq)
  contingency <- if (n > 0) sqrt(chi_sq / (chi_sq + n)) else NA_real_
  out <- data.frame(
    Measure = c("Phi coefficient", "Cramer's V", "Bias-corrected Cramer's V", "Contingency coefficient"),
    Value = c(phi, cramer, cramer_corr, contingency),
    `p-value` = c(chisq_p_value, chisq_p_value, chisq_p_value, chisq_p_value),
    Interpretation = c(
      ifelse(r == 2 && k == 2, "Recommended for 2 x 2 tables", "Generalized phi; use Cramer's V for larger tables"),
      "Nominal association strength for r x c tables",
      "Bias-corrected nominal association strength",
      "Nominal association based on chi-square"
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  round_numeric_df(out, digits)
}

pair_count_from_table <- function(tab) {
  nr <- nrow(tab)
  nc <- ncol(tab)
  counts <- as.matrix(tab)
  concordant <- 0
  discordant <- 0
  tied_x <- 0
  tied_y <- 0
  total_pairs <- choose(sum(counts), 2)
  for (i in seq_len(nr)) {
    for (j in seq_len(nc)) {
      nij <- counts[i, j]
      if (nij <= 0) next
      # Pairs tied on independent category only
      if (nc > 1) {
        tied_x <- tied_x + nij * sum(counts[i, -j, drop = FALSE])
      }
      # Pairs tied on dependent category only
      if (nr > 1) {
        tied_y <- tied_y + nij * sum(counts[-i, j, drop = FALSE])
      }
      # Concordant and discordant pairs
      if (i < nr && j < nc) {
        concordant <- concordant + nij * sum(counts[(i + 1):nr, (j + 1):nc, drop = FALSE])
      }
      if (i < nr && j > 1) {
        discordant <- discordant + nij * sum(counts[(i + 1):nr, 1:(j - 1), drop = FALSE])
      }
    }
  }
  tied_x <- tied_x / 2
  tied_y <- tied_y / 2
  list(C = concordant, D = discordant, Tied_Independent = tied_x, Tied_Dependent = tied_y, Total_Pairs = total_pairs)
}

approx_normal_p <- function(stat, effective_pairs) {
  if (is.na(stat) || !is.finite(stat) || is.na(effective_pairs) || effective_pairs <= 1) return(NA_real_)
  se <- sqrt(max(1e-12, (1 - stat^2) / effective_pairs))
  z <- stat / se
  2 * stats::pnorm(abs(z), lower.tail = FALSE)
}

compute_ordinal_association_measures <- function(tab, digits = 4) {
  pc <- pair_count_from_table(tab)
  C <- pc$C
  D <- pc$D
  Tx <- pc$Tied_Independent
  Ty <- pc$Tied_Dependent
  gamma <- if ((C + D) > 0) (C - D) / (C + D) else NA_real_
  somers_dep_given_indep <- if ((C + D + Ty) > 0) (C - D) / (C + D + Ty) else NA_real_
  somers_indep_given_dep <- if ((C + D + Tx) > 0) (C - D) / (C + D + Tx) else NA_real_
  out <- data.frame(
    Measure = c("Goodman-Kruskal Gamma", "Somers' d: Dependent | Independent", "Somers' d: Independent | Dependent"),
    Value = c(gamma, somers_dep_given_indep, somers_indep_given_dep),
    `Approximate p-value` = c(
      approx_normal_p(gamma, C + D),
      approx_normal_p(somers_dep_given_indep, C + D + Ty),
      approx_normal_p(somers_indep_given_dep, C + D + Tx)
    ),
    Concordant = c(C, C, C),
    Discordant = c(D, D, D),
    `Tied Independent` = c(Tx, Tx, Tx),
    `Tied Dependent` = c(Ty, Ty, Ty),
    Note = c(
      "Ordinal association; category order follows the manual order settings.",
      "Asymmetric ordinal association using the dependent category as outcome.",
      "Asymmetric ordinal association using the independent category as outcome."
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  round_numeric_df(out, digits)
}

compute_chi_square_test <- function(tab, approach = "Asymptotic", monte_carlo_B = 2000, digits = 4) {
  approach <- as.character(approach)
  monte_carlo_B <- safe_number(monte_carlo_B, 2000, 100, 100000)
  out <- tryCatch({
    if (approach == "Monte Carlo") {
      test <- suppressWarnings(chisq.test(tab, correct = FALSE, simulate.p.value = TRUE, B = monte_carlo_B))
      data.frame(
        Test = "Pearson Chi-square test",
        Approach = paste0("Monte Carlo simulation, B = ", monte_carlo_B),
        Statistic = unname(test$statistic),
        df = unname(test$parameter),
        `p-value` = test$p.value,
        Method = test$method,
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    } else if (approach == "Exact") {
      test <- fisher.test(tab, workspace = 2e8)
      data.frame(
        Test = "Fisher's exact test for contingency table",
        Approach = "Exact conditional test",
        Statistic = NA_real_,
        df = NA_real_,
        `p-value` = test$p.value,
        Method = test$method,
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    } else {
      test <- suppressWarnings(chisq.test(tab, correct = FALSE))
      data.frame(
        Test = "Pearson Chi-square test",
        Approach = "Asymptotic",
        Statistic = unname(test$statistic),
        df = unname(test$parameter),
        `p-value` = test$p.value,
        Method = test$method,
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    }
  }, error = function(e) {
    data.frame(
      Test = ifelse(approach == "Exact", "Fisher's exact test for contingency table", "Pearson Chi-square test"),
      Approach = approach,
      Statistic = NA_real_,
      df = NA_real_,
      `p-value` = NA_real_,
      Method = paste("Test failed:", conditionMessage(e)),
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  })
  round_numeric_df(out, digits)
}

compute_expected_table <- function(tab, digits = 4) {
  expected <- tryCatch(suppressWarnings(chisq.test(tab, correct = FALSE)$expected), error = function(e) matrix(NA_real_, nrow = nrow(tab), ncol = ncol(tab), dimnames = dimnames(tab)))
  round_numeric_df(as_table_df(round(expected, digits)), digits)
}

compute_standardized_residuals_table <- function(tab, digits = 4) {
  residuals <- tryCatch(suppressWarnings(chisq.test(tab, correct = FALSE)$stdres), error = function(e) matrix(NA_real_, nrow = nrow(tab), ncol = ncol(tab), dimnames = dimnames(tab)))
  round_numeric_df(as_table_df(round(residuals, digits)), digits)
}

# ============================================================
# SAFE STATIC EXPORT HELPERS
# ============================================================

write_error_png <- function(file, message, width = 8, height = 5, dpi = 300, bg = "white") {
  width <- safe_number(width, 8, 3, 20)
  height <- safe_number(height, 5, 3, 20)
  dpi <- safe_number(dpi, 300, 72, 1500)
  grDevices::png(filename = file, width = width, height = height, units = "in", res = dpi, bg = bg)
  on.exit(grDevices::dev.off(), add = TRUE)
  graphics::par(bg = bg, mar = c(1, 1, 1, 1))
  graphics::plot.new()
  graphics::text(0.5, 0.62, "STATCAL ONLINE", cex = 1.6, font = 2)
  graphics::text(0.5, 0.50, "PNG export could not be generated.", cex = 1.1)
  graphics::text(0.5, 0.40, paste("Reason:", message), cex = 0.8)
}

make_static_export_dir <- function() {
  export_dir <- file.path(getwd(), "www", "statcal_exports")
  if (!dir.exists(export_dir)) {
    dir.create(export_dir, recursive = TRUE, showWarnings = FALSE)
  }
  export_dir
}

make_static_export_dir()

make_export_filename <- function(prefix, ext = "png", dpi = NULL) {
  stamp <- format(Sys.time(), "%Y%m%d_%H%M%S")
  if (!is.null(dpi)) {
    paste0(prefix, "_", dpi, "dpi_", stamp, ".", ext)
  } else {
    paste0(prefix, "_", stamp, ".", ext)
  }
}

make_static_href <- function(filename) {
  paste0("statcal_exports/", utils::URLencode(filename, reserved = TRUE))
}

export_ggplot_static_png <- function(plot_function, prefix, width, height, dpi, bg = "white") {
  export_dir <- make_static_export_dir()
  width <- safe_number(width, 8, 3, 30)
  height <- safe_number(height, 6, 3, 30)
  dpi <- safe_number(dpi, 1200, 72, 1500)
  bg <- ifelse(is.null(bg) || length(bg) == 0 || is.na(bg), "white", bg)
  filename <- make_export_filename(prefix, "png", dpi)
  out_file <- file.path(export_dir, filename)
  result <- tryCatch({
    plot_object <- plot_function()
    if (!inherits(plot_object, "ggplot")) {
      stop("The selected output is not a ggplot object. Please check the selected variables.")
    }
    grDevices::png(
      filename = out_file,
      width = width,
      height = height,
      units = "in",
      res = dpi,
      bg = bg,
      type = ifelse(.Platform$OS.type == "windows", "windows", "cairo")
    )
    print(plot_object)
    grDevices::dev.off()
    if (!file.exists(out_file) || file.info(out_file)$size <= 0) {
      stop("The PNG file was not created.")
    }
    list(
      ok = TRUE,
      message = "PNG file has been generated successfully.",
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  }, error = function(e) {
    while (grDevices::dev.cur() > 1) {
      try(grDevices::dev.off(), silent = TRUE)
    }
    write_error_png(out_file, conditionMessage(e), width = width, height = height, dpi = min(dpi, 600), bg = bg)
    list(
      ok = FALSE,
      message = paste("PNG export failed, but an error PNG was generated:", conditionMessage(e)),
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  })
  result
}

static_export_link_ui <- function(result, button_label = "Download generated file", open_label = "Open file in new tab", preview_image = FALSE) {
  if (is.null(result)) {
    return(tags$p("Click Generate first, then the download link will appear here."))
  }
  ui <- tagList(
    tags$p(if (isTRUE(result$ok)) result$message else result$message),
    tags$a(
      href = result$href,
      download = result$filename,
      target = "_blank",
      class = "btn btn-success",
      icon("download"),
      button_label
    ),
    tags$span(" "),
    tags$a(
      href = result$href,
      target = "_blank",
      class = "btn btn-info",
      icon("external-link-alt"),
      open_label
    ),
    tags$p(style = "font-size: 12px; margin-top: 8px; color: #555;", paste("Generated file:", result$filename)),
    tags$p(style = "font-size: 11px; color: #777; word-break: break-all;", paste("Local file path:", result$file))
  )
  if (preview_image) {
    ui <- tagList(ui, tags$img(src = result$href, style = "max-width: 100%; margin-top: 8px; border: 1px solid #ddd;"))
  }
  ui
}

# ============================================================
# EXCEL EXPORT HELPERS
# ============================================================

safe_sheet_name <- function(name) {
  name <- as.character(name)
  invalid_chars <- c("\\", "/", "?", "*", "[", "]", ":")
  for (ch in invalid_chars) {
    name <- gsub(ch, "_", name, fixed = TRUE)
  }
  name <- trimws(name)
  name <- substr(name, 1, 31)
  ifelse(nchar(name) == 0, "Sheet", name)
}

write_table_sheet <- function(wb, sheet_name, df) {
  sheet_name <- safe_sheet_name(sheet_name)
  openxlsx::addWorksheet(wb, sheet_name)
  df <- make_display_safe(as.data.frame(df))
  openxlsx::writeData(wb, sheet_name, df)
  if (ncol(df) > 0) {
    openxlsx::setColWidths(wb, sheet_name, cols = 1:ncol(df), widths = "auto")
    header_style <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#D9EAF7", border = "Bottom")
    openxlsx::addStyle(wb, sheet_name, header_style, rows = 1, cols = 1:ncol(df), gridExpand = TRUE)
    openxlsx::freezePane(wb, sheet_name, firstRow = TRUE)
  }
}

export_workbook <- function(file, metadata_df, filtered_df, univ_df, bivar_df, bivar_wide_df, univ_chart_df, bivar_chart_df,
                            assoc_observed_df = NULL, assoc_expected_df = NULL, assoc_residuals_df = NULL,
                            assoc_chisq_df = NULL, assoc_nominal_df = NULL, assoc_ordinal_df = NULL) {
  wb <- openxlsx::createWorkbook()
  write_table_sheet(wb, "Export Info", metadata_df)
  write_table_sheet(wb, "Filtered Data", filtered_df)
  write_table_sheet(wb, "Univariate Table", univ_df)
  write_table_sheet(wb, "Bivariate Table Long", bivar_df)
  write_table_sheet(wb, "Bivariate Table Wide", bivar_wide_df)
  write_table_sheet(wb, "Univariate Chart Data", univ_chart_df)
  write_table_sheet(wb, "Bivariate Chart Data", bivar_chart_df)
  if (!is.null(assoc_observed_df)) write_table_sheet(wb, "Assoc Observed Table", assoc_observed_df)
  if (!is.null(assoc_expected_df)) write_table_sheet(wb, "Assoc Expected Table", assoc_expected_df)
  if (!is.null(assoc_residuals_df)) write_table_sheet(wb, "Assoc Std Residuals", assoc_residuals_df)
  if (!is.null(assoc_chisq_df)) write_table_sheet(wb, "Assoc Chi-Square", assoc_chisq_df)
  if (!is.null(assoc_nominal_df)) write_table_sheet(wb, "Assoc Nominal Measures", assoc_nominal_df)
  if (!is.null(assoc_ordinal_df)) write_table_sheet(wb, "Assoc Ordinal Measures", assoc_ordinal_df)
  openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
}

export_excel_static_file <- function(export_function, prefix) {
  export_dir <- make_static_export_dir()
  filename <- make_export_filename(prefix, "xlsx")
  out_file <- file.path(export_dir, filename)
  result <- tryCatch({
    export_function(out_file)
    if (!file.exists(out_file) || file.info(out_file)$size <= 0) {
      stop("The Excel file was not created.")
    }
    list(
      ok = TRUE,
      message = "Excel file has been generated successfully.",
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  }, error = function(e) {
    list(
      ok = FALSE,
      message = paste("Excel export failed:", conditionMessage(e)),
      file = normalizePath(out_file, winslash = "/", mustWork = FALSE),
      href = make_static_href(filename),
      filename = filename
    )
  })
  result
}

# ============================================================
# UI
# ============================================================

legend_choices <- c("Right", "Left", "Top", "Bottom", "None / Hide legend")
facet_scale_choices <- c(
  "Same scale across panels" = "fixed",
  "Free Y scale by panel" = "free_y",
  "Free X scale by panel" = "free_x",
  "Free X and Y scale" = "free"
)

ui <- dashboardPage(
  dashboardHeader(title = APP_NAME, titleWidth = "100%"),
  dashboardSidebar(disable = TRUE),
  dashboardBody(
    tags$head(
      tags$style(HTML("\n        .content-wrapper, .right-side { background-color: #f7f9fb; }\n        .box { border-radius: 10px; }\n        .statcal-title { font-size: 24px; font-weight: 700; color: #1F4E79; }\n        .statcal-subtitle { font-size: 18px; font-weight: 600; color: #333333; }\n        .statcal-note { line-height: 1.6; text-align: justify; }\n        .small-note { font-size: 12px; color: #666666; }\n      "))
    ),
    fluidRow(
      box(
        width = 12, status = "primary", solidHeader = TRUE,
        title = "STATCAL ONLINE for Categorical Frequency Distribution and Bivariate Bar Chart Analyzer",
        div(class = "statcal-title", APP_TITLE),
        div(class = "statcal-subtitle", APP_UPDATED),
        tags$p(class = "statcal-note",
               "This R Shiny application is designed to analyze categorical variables through univariate and bivariate frequency-percentage tables. Users can create flexible publication-ready bar charts, organize charts into panels, customize colors and text sizes, select scientific background themes, and export charts or tables for academic reporting."
        ),
        tags$p(
          tags$b("Website: "), tags$a(href = WEBSITE_URL, target = "_blank", WEBSITE_URL), tags$br(),
          tags$b("STATCAL ONLINE Page: "), tags$a(href = STATCAL_ONLINE_URL, target = "_blank", STATCAL_ONLINE_URL), tags$br(),
          tags$b("Training Data: "), tags$a(href = URL_DATA_TRAINING, target = "_blank", "Open Google Drive Folder")
        )
      )
    ),
    tabsetPanel(
      id = "main_tabs",
      tabPanel(
        "1. Data & Settings",
        br(),
        fluidRow(
          box(width = 5, title = "Data Input", status = "primary", solidHeader = TRUE,
              fileInput("uploaded_file", "Upload Excel file", accept = c(".xlsx", ".xls")),
              uiOutput("sheet_ui"),
              tags$p(class = "small-note", "If no file is uploaded, the application uses the sample data variabel kategori.xlsx file.")),
          box(width = 7, title = "Variables and Filters", status = "primary", solidHeader = TRUE,
              uiOutput("category_vars_ui"),
              selectizeInput("filter_vars", "Optional filter variables", choices = NULL, selected = NULL, multiple = TRUE),
              uiOutput("filter_controls_ui"))
        ),
        fluidRow(
          valueBoxOutput("metric_original_rows", width = 3),
          valueBoxOutput("metric_filtered_rows", width = 3),
          valueBoxOutput("metric_columns", width = 3),
          valueBoxOutput("metric_selected_vars", width = 3)
        ),
        fluidRow(
          box(width = 12, title = "Dataset Preview", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(DTOutput("data_preview")))
        )
      ),
      tabPanel(
        "2. Univariate Table",
        br(),
        fluidRow(
          box(width = 8, title = "Univariate Table Settings", status = "primary", solidHeader = TRUE,
              uiOutput("univar_table_vars_ui"),
              textInput("univar_category_order", "Optional category order (comma-separated)", value = ""),
              tags$p(class = "small-note", "Leave blank to use category order detected from the data.")),
          box(width = 4, title = "Decimal Digits", status = "primary", solidHeader = TRUE,
              sliderInput("univar_digits", "Decimal digits", min = 0, max = 8, value = 2, step = 1))
        ),
        fluidRow(
          box(width = 12, title = "Univariate Frequency and Percentage Table", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(DTOutput("univar_table")))
        )
      ),
      tabPanel(
        "3. Bivariate Table",
        br(),
        fluidRow(
          box(width = 4, title = "Dependent Variable", status = "primary", solidHeader = TRUE,
              uiOutput("bivar_dependent_ui"),
              textInput("bivar_dependent_order", "Optional dependent category order", value = "")),
          box(width = 5, title = "Independent Variables", status = "primary", solidHeader = TRUE,
              uiOutput("bivar_independent_ui"),
              textInput("bivar_independent_order", "Optional independent category order", value = "")),
          box(width = 3, title = "Table Settings", status = "primary", solidHeader = TRUE,
              sliderInput("bivar_digits", "Decimal digits", min = 0, max = 8, value = 2, step = 1),
              selectInput("bivar_table_format", "Displayed table format", choices = c("Long", "Wide"), selected = "Long"))
        ),
        fluidRow(
          box(width = 12, title = "Bivariate Frequency and Percentage Table", status = "warning", solidHeader = TRUE,
              tags$p("Percentages are computed within each independent-variable category."),
              shinycssloaders::withSpinner(DTOutput("bivar_table")))
        )
      ),
      tabPanel(
        "4. Univariate Bar Chart",
        br(),
        fluidRow(
          box(width = 3, title = "Chart Data", status = "primary", solidHeader = TRUE,
              uiOutput("univar_chart_vars_ui"),
              selectInput("univar_chart_metric", "Bar height", choices = c("Frequency", "Percentage"), selected = "Percentage"),
              selectInput("univar_label_mode", "Show information on bars", choices = c("None", "Frequency", "Percentage", "Frequency and Percentage"), selected = "Percentage"),
              selectInput("univar_orientation", "Bar orientation", choices = c("Vertical", "Horizontal"), selected = "Vertical")),
          box(width = 3, title = "Panel and Scale", status = "primary", solidHeader = TRUE,
              sliderInput("univar_panel_cols", "Panel columns", min = 1, max = 10, value = 2, step = 1),
              selectInput("univar_facet_scales", "Panel axis scale", choices = facet_scale_choices, selected = "free_x"),
              selectInput("univar_legend_position", "Legend position", choices = legend_choices, selected = "Right")),
          box(width = 3, title = "Bar Style", status = "primary", solidHeader = TRUE,
              selectInput("univar_theme", "Background theme", choices = names(THEMES), selected = "White Publication"),
              sliderInput("univar_bar_width", "Bar width", min = 0.25, max = 1.00, value = 0.72, step = 0.01),
              sliderInput("univar_chart_height_px", "Preview chart height (px)", min = 350, max = 1200, value = 700, step = 50),
              sliderInput("univar_x_text_angle", "X-axis text angle", min = 0, max = 90, value = 35, step = 5)),
          box(width = 3, title = "Title and Labels", status = "primary", solidHeader = TRUE,
              textInput("univar_chart_title", "Title", value = "Univariate Categorical Distribution"),
              textInput("univar_chart_subtitle", "Subtitle", value = "Frequency or percentage distribution by categorical variable"),
              textInput("univar_x_axis_title", "Category axis title", value = "Category"),
              textInput("univar_legend_title", "Legend title", value = "Category"),
              checkboxInput("univar_show_labels", "Show text labels", value = TRUE),
              sliderInput("univar_label_size", "Label text size", min = 2, max = 8, value = 3.2, step = 0.2),
              textInput("univar_label_color", "Label text color", value = "#111111"))
        ),
        fluidRow(
          box(width = 8, title = "Flexible Text Size Settings", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              fluidRow(
                column(2, sliderInput("univar_title_size", "Title", 8, 34, 16, 1)),
                column(2, sliderInput("univar_subtitle_size", "Subtitle", 6, 26, 11, 1)),
                column(2, sliderInput("univar_axis_title_size", "Axis title", 6, 24, 11, 1)),
                column(2, sliderInput("univar_axis_text_size", "Axis text", 5, 22, 9, 1)),
                column(2, sliderInput("univar_panel_title_size", "Panel title", 6, 26, 11, 1)),
                column(2, sliderInput("univar_legend_text_size", "Legend text", 5, 22, 9, 1))
              ),
              fluidRow(
                column(2, sliderInput("univar_legend_title_size", "Legend title", 5, 24, 10, 1))
              )),
          box(width = 4, title = "Manual Colors by Category", status = "primary", solidHeader = TRUE,
              tags$p(class = "small-note", "Use HEX color codes, for example #2166AC."),
              uiOutput("univar_color_settings_ui"))
        ),
        fluidRow(
          box(width = 12, title = "Publication-Ready Univariate Bar Chart", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(uiOutput("univar_plot_ui")))
        ),
        fluidRow(
          box(width = 12, title = "Univariate Chart Data", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              shinycssloaders::withSpinner(DTOutput("univar_chart_data_table")))
        )
      ),
      tabPanel(
        "5. Bivariate Bar Chart",
        br(),
        fluidRow(
          box(width = 3, title = "Chart Data", status = "primary", solidHeader = TRUE,
              selectInput("bivar_chart_metric", "Bar height", choices = c("Frequency", "Percentage"), selected = "Percentage"),
              selectInput("bivar_label_mode", "Show information on bars", choices = c("None", "Frequency", "Percentage", "Frequency and Percentage"), selected = "Percentage"),
              selectInput("bivar_bar_position", "Bar position", choices = c("Stacked", "Grouped"), selected = "Stacked"),
              selectInput("bivar_orientation", "Bar orientation", choices = c("Vertical", "Horizontal"), selected = "Vertical")),
          box(width = 3, title = "Panel and Scale", status = "primary", solidHeader = TRUE,
              sliderInput("bivar_panel_cols", "Panel columns", min = 1, max = 10, value = 2, step = 1),
              selectInput("bivar_facet_scales", "Panel axis scale", choices = facet_scale_choices, selected = "free_x"),
              selectInput("bivar_legend_position", "Legend position", choices = legend_choices, selected = "Right")),
          box(width = 3, title = "Bar Style", status = "primary", solidHeader = TRUE,
              selectInput("bivar_theme", "Background theme", choices = names(THEMES), selected = "White Publication"),
              sliderInput("bivar_bar_width", "Bar width", min = 0.25, max = 1.00, value = 0.72, step = 0.01),
              sliderInput("bivar_chart_height_px", "Preview chart height (px)", min = 350, max = 1200, value = 700, step = 50),
              sliderInput("bivar_x_text_angle", "X-axis text angle", min = 0, max = 90, value = 35, step = 5)),
          box(width = 3, title = "Title and Labels", status = "primary", solidHeader = TRUE,
              textInput("bivar_chart_title", "Title", value = "Bivariate Categorical Distribution"),
              textInput("bivar_chart_subtitle", "Subtitle", value = "Distribution of dependent categories by independent variables"),
              textInput("bivar_x_axis_title", "Independent Category title", value = "Independent Category"),
              textInput("bivar_legend_title", "Dependent Category / legend title", value = "Dependent Category"),
              checkboxInput("bivar_show_labels", "Show text labels", value = TRUE),
              sliderInput("bivar_label_size", "Label text size", min = 2, max = 8, value = 3.0, step = 0.2),
              textInput("bivar_label_color", "Label text color", value = "#111111"),
              sliderInput("bivar_label_vjust", "Label position inside stack", min = 0.05, max = 0.95, value = 0.50, step = 0.05))
        ),
        fluidRow(
          box(width = 8, title = "Flexible Text Size Settings", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              fluidRow(
                column(2, sliderInput("bivar_title_size", "Title", 8, 34, 16, 1)),
                column(2, sliderInput("bivar_subtitle_size", "Subtitle", 6, 26, 11, 1)),
                column(2, sliderInput("bivar_axis_title_size", "Axis title", 6, 24, 11, 1)),
                column(2, sliderInput("bivar_axis_text_size", "Axis text", 5, 22, 9, 1)),
                column(2, sliderInput("bivar_panel_title_size", "Panel title", 6, 26, 11, 1)),
                column(2, sliderInput("bivar_legend_text_size", "Legend text", 5, 22, 9, 1))
              ),
              fluidRow(
                column(2, sliderInput("bivar_legend_title_size", "Legend title", 5, 24, 10, 1))
              )),
          box(width = 4, title = "Manual Colors by Dependent Category", status = "primary", solidHeader = TRUE,
              tags$p(class = "small-note", "Use HEX color codes, for example #2166AC."),
              uiOutput("bivar_color_settings_ui"))
        ),
        fluidRow(
          box(width = 12, title = "Publication-Ready Bivariate Bar Chart", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(uiOutput("bivar_plot_ui")))
        ),
        fluidRow(
          box(width = 12, title = "Bivariate Chart Data", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              shinycssloaders::withSpinner(DTOutput("bivar_chart_data_table")))
        )
      ),
      tabPanel(
        "6. Categorical Association Tests",
        br(),
        fluidRow(
          box(width = 4, title = "Variables", status = "primary", solidHeader = TRUE,
              uiOutput("assoc_independent_ui"),
              uiOutput("assoc_dependent_ui"),
              textInput("assoc_independent_order", "Optional Independent Category order", value = ""),
              textInput("assoc_dependent_order", "Optional Dependent Category order", value = ""),
              tags$p(class = "small-note", "For Gamma and Somers' d, write category order from lowest to highest when the categories are ordinal.")),
          box(width = 4, title = "Chi-square Settings", status = "primary", solidHeader = TRUE,
              selectInput("assoc_chisq_approach", "Chi-square / exact approach",
                          choices = c("Asymptotic", "Monte Carlo", "Exact"), selected = "Asymptotic"),
              numericInput("assoc_monte_carlo_B", "Monte Carlo replications", value = 2000, min = 100, max = 100000, step = 100),
              sliderInput("assoc_digits", "Decimal digits", min = 0, max = 8, value = 4, step = 1),
              tags$p(class = "small-note", "Exact uses Fisher's exact test for contingency tables. Monte Carlo uses simulated p-value for Pearson chi-square.")),
          box(width = 4, title = "Export", status = "success", solidHeader = TRUE,
              actionButton("generate_assoc_excel", "Generate Association Tests Excel", icon = icon("file-excel")),
              br(), br(), uiOutput("assoc_excel_static_download_ui"),
              br(), downloadButton("download_assoc_excel_fallback", "Fallback Download Association Excel"))
        ),
        fluidRow(
          box(width = 12, title = "Observed Contingency Table", status = "warning", solidHeader = TRUE,
              shinycssloaders::withSpinner(DTOutput("assoc_observed_table")))
        ),
        fluidRow(
          box(width = 6, title = "Chi-square / Exact Test Result", status = "info", solidHeader = TRUE,
              shinycssloaders::withSpinner(DTOutput("assoc_chisq_table"))),
          box(width = 6, title = "Nominal Association Measures", status = "info", solidHeader = TRUE,
              shinycssloaders::withSpinner(DTOutput("assoc_nominal_table")))
        ),
        fluidRow(
          box(width = 12, title = "Ordinal Association Measures", status = "primary", solidHeader = TRUE,
              tags$p("Gamma and Somers' d use the manual category order. P-values are approximate normal-test p-values."),
              shinycssloaders::withSpinner(DTOutput("assoc_ordinal_table")))
        ),
        fluidRow(
          box(width = 6, title = "Expected Counts", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              shinycssloaders::withSpinner(DTOutput("assoc_expected_table"))),
          box(width = 6, title = "Standardized Residuals", status = "info", solidHeader = TRUE, collapsible = TRUE, collapsed = TRUE,
              shinycssloaders::withSpinner(DTOutput("assoc_residuals_table")))
        )
      ),
      tabPanel(
        "7. Export",
        br(),
        fluidRow(
          box(width = 12, title = "Export Settings", status = "primary", solidHeader = TRUE,
              fluidRow(
                column(4, selectInput("export_dpi", "PNG resolution / DPI", choices = c(300, 600, 900, 1200, 1500), selected = 1200)),
                column(4, numericInput("export_width", "Export width (inches)", value = 8, min = 4, max = 30, step = 0.5)),
                column(4, numericInput("export_height", "Export height (inches)", value = 6, min = 3, max = 30, step = 0.5))
              ),
              tags$p(tags$b("Default DPI: "), "1200 DPI for publication-ready output."))
        ),
        fluidRow(
          box(width = 4, title = "Univariate Bar Chart PNG", status = "warning", solidHeader = TRUE,
              actionButton("generate_univar_png", "Generate Univariate Chart PNG", icon = icon("image")),
              br(), br(), uiOutput("univar_static_download_ui")),
          box(width = 4, title = "Bivariate Bar Chart PNG", status = "warning", solidHeader = TRUE,
              actionButton("generate_bivar_png", "Generate Bivariate Chart PNG", icon = icon("image")),
              br(), br(), uiOutput("bivar_static_download_ui")),
          box(width = 4, title = "Tables Excel", status = "success", solidHeader = TRUE,
              actionButton("generate_excel", "Generate Frequency Tables Excel", icon = icon("file-excel")),
              br(), br(), uiOutput("excel_static_download_ui"),
              br(), downloadButton("download_excel_fallback", "Fallback Download Excel"))
        )
      )
    )
  )
)

# ============================================================
# SERVER
# ============================================================

server <- function(input, output, session) {
  univar_export_result <- reactiveVal(NULL)
  bivar_export_result <- reactiveVal(NULL)
  excel_export_result <- reactiveVal(NULL)
  assoc_excel_export_result <- reactiveVal(NULL)
  
  output$univar_static_download_ui <- renderUI({
    static_export_link_ui(univar_export_result(), "Download Univariate Chart PNG", "Open PNG in new tab", preview_image = TRUE)
  })
  output$bivar_static_download_ui <- renderUI({
    static_export_link_ui(bivar_export_result(), "Download Bivariate Chart PNG", "Open PNG in new tab", preview_image = TRUE)
  })
  output$excel_static_download_ui <- renderUI({
    static_export_link_ui(excel_export_result(), "Download Frequency Tables Excel", "Open Excel file in new tab", preview_image = FALSE)
  })
  output$assoc_excel_static_download_ui <- renderUI({
    static_export_link_ui(assoc_excel_export_result(), "Download Association Tests Excel", "Open Excel file in new tab", preview_image = FALSE)
  })
  
  current_excel_path <- reactive({
    if (!is.null(input$uploaded_file)) {
      input$uploaded_file$datapath
    } else if (file.exists(SAMPLE_DATA_PATH)) {
      SAMPLE_DATA_PATH
    } else {
      NULL
    }
  })
  
  output$sheet_ui <- renderUI({
    path <- current_excel_path()
    if (is.null(path)) {
      return(helpText("Please upload an Excel file to start the analysis."))
    }
    sheets <- readxl::excel_sheets(path)
    selectInput("sheet_name", "Worksheet", choices = sheets, selected = sheets[1])
  })
  
  data_raw <- reactive({
    path <- current_excel_path()
    req(path)
    sheets <- readxl::excel_sheets(path)
    sheet <- input$sheet_name
    if (is.null(sheet) || !(sheet %in% sheets)) sheet <- sheets[1]
    clean_dataframe(readxl::read_excel(path, sheet = sheet))
  })
  
  category_columns <- reactive({
    names(data_raw())
  })
  
  selected_category_vars <- reactive({
    cols <- category_columns()
    if (is.null(input$category_vars) || length(input$category_vars) == 0) cols else input$category_vars
  })
  
  output$category_vars_ui <- renderUI({
    cols <- category_columns()
    preferred <- intersect(c("Pendidikan", "Jenis Kelamin", "Golongan Darah", "Pekerjaan"), cols)
    selected <- if (length(preferred) > 0) preferred else cols
    selectizeInput("category_vars", "Select categorical variables for analysis", choices = cols, selected = selected, multiple = TRUE)
  })
  
  observe({
    cols <- category_columns()
    updateSelectizeInput(session, "filter_vars", choices = cols, selected = character(0), server = TRUE)
  })
  
  output$filter_controls_ui <- renderUI({
    df <- data_raw()
    fvars <- input$filter_vars
    if (is.null(fvars) || length(fvars) == 0) {
      return(tags$p(class = "small-note", "No filters selected. The full dataset is used."))
    }
    tagList(lapply(fvars, function(v) {
      vals <- sorted_unique_values(df[[v]])
      selectizeInput(paste0("filter_", safe_id(v)), paste0("Filter: ", v), choices = vals, selected = vals, multiple = TRUE)
    }))
  })
  
  filtered_data <- reactive({
    df <- data_raw()
    fvars <- input$filter_vars
    if (!is.null(fvars) && length(fvars) > 0) {
      for (v in fvars) {
        input_id <- paste0("filter_", safe_id(v))
        chosen <- input[[input_id]]
        if (!is.null(chosen) && length(chosen) > 0) {
          df <- df[clean_category_value(df[[v]]) %in% chosen, , drop = FALSE]
        }
      }
    }
    df
  })
  
  output$metric_original_rows <- renderValueBox({
    valueBox(nrow(data_raw()), "Original rows", icon = icon("table"), color = "blue")
  })
  output$metric_filtered_rows <- renderValueBox({
    valueBox(nrow(filtered_data()), "Rows after filtering", icon = icon("filter"), color = "green")
  })
  output$metric_columns <- renderValueBox({
    valueBox(ncol(data_raw()), "Columns", icon = icon("columns"), color = "yellow")
  })
  output$metric_selected_vars <- renderValueBox({
    valueBox(length(selected_category_vars()), "Selected variables", icon = icon("list"), color = "purple")
  })
  output$data_preview <- renderDT({
    DT::datatable(make_display_safe(filtered_data()), options = list(scrollX = TRUE, pageLength = 10))
  })
  
  output$univar_table_vars_ui <- renderUI({
    cols <- category_columns()
    selectizeInput("univar_table_vars", "Select variables", choices = cols, selected = selected_category_vars(), multiple = TRUE)
  })
  
  univar_table_data <- reactive({
    vars <- input$univar_table_vars
    if (is.null(vars) || length(vars) == 0) vars <- selected_category_vars()
    compute_univariate_distribution(filtered_data(), vars, input$univar_digits, input$univar_category_order)
  })
  
  output$univar_table <- renderDT({
    DT::datatable(make_display_safe(univar_table_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  
  output$bivar_dependent_ui <- renderUI({
    cols <- category_columns()
    selected <- if ("Jenis Kelamin" %in% cols) "Jenis Kelamin" else cols[1]
    selectInput("bivar_dependent", "Dependent variable", choices = cols, selected = selected)
  })
  
  output$bivar_independent_ui <- renderUI({
    cols <- category_columns()
    dep <- input$bivar_dependent
    choices <- setdiff(cols, dep)
    preferred <- intersect(c("Pendidikan", "Golongan Darah", "Pekerjaan"), choices)
    selected <- if (length(preferred) > 0) preferred else choices[seq_len(min(2, length(choices)))]
    selectizeInput("bivar_independent", "Independent variables", choices = choices, selected = selected, multiple = TRUE)
  })
  
  bivar_table_data <- reactive({
    req(input$bivar_dependent, input$bivar_independent)
    compute_bivariate_distribution(
      filtered_data(),
      dependent_var = input$bivar_dependent,
      independent_vars = input$bivar_independent,
      digits = input$bivar_digits,
      dependent_order = input$bivar_dependent_order,
      independent_order = input$bivar_independent_order
    )
  })
  
  bivar_wide_table_data <- reactive({
    make_bivariate_wide_table(bivar_table_data())
  })
  
  output$bivar_table <- renderDT({
    table_to_show <- if (input$bivar_table_format == "Wide") bivar_wide_table_data() else bivar_table_data()
    DT::datatable(make_display_safe(table_to_show), options = list(scrollX = TRUE, pageLength = 15))
  })
  
  output$univar_chart_vars_ui <- renderUI({
    cols <- category_columns()
    selectizeInput("univar_chart_vars", "Select variables for chart panels", choices = cols, selected = selected_category_vars(), multiple = TRUE)
  })
  
  univar_chart_data <- reactive({
    vars <- input$univar_chart_vars
    if (is.null(vars) || length(vars) == 0) vars <- selected_category_vars()
    df <- compute_univariate_distribution(filtered_data(), vars, input$univar_digits, input$univar_category_order)
    prepare_univariate_chart_data(df, input$univar_category_order)
  })
  
  output$univar_color_settings_ui <- renderUI({
    df <- univar_chart_data()
    cats <- unique(as.character(df$Category))
    tagList(lapply(seq_along(cats), function(i) {
      textInput(paste0("univar_color_", safe_id(cats[i])), paste0("Color for ", cats[i]), value = DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1])
    }))
  })
  
  univar_plot_object <- reactive({
    cats <- unique(as.character(univar_chart_data()$Category))
    create_univariate_bar_plot(
      chart_df = univar_chart_data(),
      manual_colors = get_manual_colors(cats, input, "univar_color_"),
      title = input$univar_chart_title,
      subtitle = input$univar_chart_subtitle,
      y_metric = input$univar_chart_metric,
      label_mode = input$univar_label_mode,
      orientation = input$univar_orientation,
      theme_name = input$univar_theme,
      bar_width = input$univar_bar_width,
      show_labels = input$univar_show_labels,
      label_size = input$univar_label_size,
      label_color = input$univar_label_color,
      panel_cols = input$univar_panel_cols,
      facet_scales = input$univar_facet_scales,
      legend_position = input$univar_legend_position,
      title_size = input$univar_title_size,
      subtitle_size = input$univar_subtitle_size,
      axis_title_size = input$univar_axis_title_size,
      axis_text_size = input$univar_axis_text_size,
      panel_title_size = input$univar_panel_title_size,
      legend_title_size = input$univar_legend_title_size,
      legend_text_size = input$univar_legend_text_size,
      x_text_angle = input$univar_x_text_angle,
      legend_title = input$univar_legend_title,
      x_axis_title = input$univar_x_axis_title,
      digits = input$univar_digits
    )
  })
  
  output$univar_plot_ui <- renderUI({
    plotOutput("univar_bar_plot", height = paste0(safe_number(input$univar_chart_height_px, 700, 350, 1400), "px"))
  })
  output$univar_bar_plot <- renderPlot({
    univar_plot_object()
  })
  output$univar_chart_data_table <- renderDT({
    DT::datatable(make_display_safe(univar_chart_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  
  bivar_chart_data <- reactive({
    prepare_bivariate_chart_data(bivar_table_data(), input$bivar_dependent_order, input$bivar_independent_order)
  })
  
  output$bivar_color_settings_ui <- renderUI({
    df <- bivar_chart_data()
    cats <- unique(as.character(df$`Dependent Category`))
    tagList(lapply(seq_along(cats), function(i) {
      textInput(paste0("bivar_color_", safe_id(cats[i])), paste0("Color for ", cats[i]), value = DEFAULT_COLORS[((i - 1) %% length(DEFAULT_COLORS)) + 1])
    }))
  })
  
  bivar_plot_object <- reactive({
    cats <- unique(as.character(bivar_chart_data()$`Dependent Category`))
    create_bivariate_bar_plot(
      chart_df = bivar_chart_data(),
      manual_colors = get_manual_colors(cats, input, "bivar_color_"),
      title = input$bivar_chart_title,
      subtitle = input$bivar_chart_subtitle,
      y_metric = input$bivar_chart_metric,
      label_mode = input$bivar_label_mode,
      bar_position = input$bivar_bar_position,
      orientation = input$bivar_orientation,
      theme_name = input$bivar_theme,
      bar_width = input$bivar_bar_width,
      show_labels = input$bivar_show_labels,
      label_size = input$bivar_label_size,
      label_color = input$bivar_label_color,
      label_vjust = input$bivar_label_vjust,
      panel_cols = input$bivar_panel_cols,
      facet_scales = input$bivar_facet_scales,
      legend_position = input$bivar_legend_position,
      title_size = input$bivar_title_size,
      subtitle_size = input$bivar_subtitle_size,
      axis_title_size = input$bivar_axis_title_size,
      axis_text_size = input$bivar_axis_text_size,
      panel_title_size = input$bivar_panel_title_size,
      legend_title_size = input$bivar_legend_title_size,
      legend_text_size = input$bivar_legend_text_size,
      x_text_angle = input$bivar_x_text_angle,
      legend_title = input$bivar_legend_title,
      x_axis_title = input$bivar_x_axis_title,
      digits = input$bivar_digits
    )
  })
  
  output$bivar_plot_ui <- renderUI({
    plotOutput("bivar_bar_plot", height = paste0(safe_number(input$bivar_chart_height_px, 700, 350, 1400), "px"))
  })
  output$bivar_bar_plot <- renderPlot({
    bivar_plot_object()
  })
  output$bivar_chart_data_table <- renderDT({
    DT::datatable(make_display_safe(bivar_chart_data()), options = list(scrollX = TRUE, pageLength = 15))
  })
  
  
  # ------------------ Categorical association tests ------------------
  output$assoc_independent_ui <- renderUI({
    cols <- category_columns()
    selected <- if (length(input$bivar_independent) > 0) input$bivar_independent[1] else cols[1]
    selectInput("assoc_independent_var", "Independent categorical variable", choices = cols, selected = selected)
  })
  output$assoc_dependent_ui <- renderUI({
    cols <- category_columns()
    indep <- input$assoc_independent_var
    choices <- setdiff(cols, indep)
    selected <- if (!is.null(input$bivar_dependent) && input$bivar_dependent %in% choices) input$bivar_dependent else choices[1]
    selectInput("assoc_dependent_var", "Dependent categorical variable", choices = choices, selected = selected)
  })
  
  assoc_contingency_table <- reactive({
    cols <- category_columns()
    row_var <- input$assoc_independent_var
    if (is.null(row_var) || !(row_var %in% cols)) {
      row_var <- if (!is.null(input$bivar_independent) && length(input$bivar_independent) > 0) input$bivar_independent[1] else cols[1]
    }
    col_var <- input$assoc_dependent_var
    if (is.null(col_var) || !(col_var %in% cols) || identical(col_var, row_var)) {
      candidate_cols <- setdiff(cols, row_var)
      col_var <- if (!is.null(input$bivar_dependent) && input$bivar_dependent %in% candidate_cols) input$bivar_dependent else candidate_cols[1]
    }
    validate(need(!is.na(row_var) && !is.na(col_var), "Please select two different categorical variables."))
    make_contingency_table(
      filtered_data(),
      row_var = row_var,
      col_var = col_var,
      row_order = input$assoc_independent_order,
      col_order = input$assoc_dependent_order
    )
  })
  
  assoc_chisq_data <- reactive({
    compute_chi_square_test(
      assoc_contingency_table(),
      approach = input$assoc_chisq_approach,
      monte_carlo_B = input$assoc_monte_carlo_B,
      digits = input$assoc_digits
    )
  })
  assoc_nominal_data <- reactive({
    p_value <- assoc_chisq_data()[["p-value"]][1]
    compute_nominal_association_measures(assoc_contingency_table(), chisq_p_value = p_value, digits = input$assoc_digits)
  })
  assoc_ordinal_data <- reactive({
    compute_ordinal_association_measures(assoc_contingency_table(), digits = input$assoc_digits)
  })
  assoc_observed_data <- reactive({
    as_table_df(assoc_contingency_table())
  })
  assoc_expected_data <- reactive({
    compute_expected_table(assoc_contingency_table(), digits = input$assoc_digits)
  })
  assoc_residuals_data <- reactive({
    compute_standardized_residuals_table(assoc_contingency_table(), digits = input$assoc_digits)
  })
  
  output$assoc_observed_table <- renderDT({
    DT::datatable(make_display_safe(assoc_observed_data()), options = list(scrollX = TRUE, pageLength = 15), rownames = FALSE)
  })
  output$assoc_chisq_table <- renderDT({
    DT::datatable(make_display_safe(assoc_chisq_data()), options = list(scrollX = TRUE, pageLength = 10), rownames = FALSE)
  })
  output$assoc_nominal_table <- renderDT({
    DT::datatable(make_display_safe(assoc_nominal_data()), options = list(scrollX = TRUE, pageLength = 10), rownames = FALSE)
  })
  output$assoc_ordinal_table <- renderDT({
    DT::datatable(make_display_safe(assoc_ordinal_data()), options = list(scrollX = TRUE, pageLength = 10), rownames = FALSE)
  })
  output$assoc_expected_table <- renderDT({
    DT::datatable(make_display_safe(assoc_expected_data()), options = list(scrollX = TRUE, pageLength = 15), rownames = FALSE)
  })
  output$assoc_residuals_table <- renderDT({
    DT::datatable(make_display_safe(assoc_residuals_data()), options = list(scrollX = TRUE, pageLength = 15), rownames = FALSE)
  })
  
  export_association_excel <- function(file) {
    wb <- openxlsx::createWorkbook()
    write_table_sheet(wb, "Observed Table", assoc_observed_data())
    write_table_sheet(wb, "Expected Counts", assoc_expected_data())
    write_table_sheet(wb, "Standardized Residuals", assoc_residuals_data())
    write_table_sheet(wb, "Chi-Square Exact Test", assoc_chisq_data())
    write_table_sheet(wb, "Nominal Measures", assoc_nominal_data())
    write_table_sheet(wb, "Ordinal Measures", assoc_ordinal_data())
    openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  }
  
  observeEvent(input$generate_assoc_excel, {
    result <- export_excel_static_file(
      export_function = function(path) export_association_excel(path),
      prefix = "statcal_online_categorical_association_tests"
    )
    assoc_excel_export_result(result)
  })
  
  output$download_assoc_excel_fallback <- downloadHandler(
    filename = function() paste0("statcal_online_categorical_association_tests_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx"),
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    content = function(file) {
      export_association_excel(file)
    }
  )
  
  build_export_metadata <- reactive({
    data.frame(
      Item = c(
        "Application", "Export Time", "Rows After Filtering", "Selected Variables",
        "Bivariate Dependent Variable", "Bivariate Independent Variables", "Training Data URL"
      ),
      Value = c(
        APP_TITLE, as.character(Sys.time()), as.character(nrow(filtered_data())),
        paste(selected_category_vars(), collapse = ", "),
        ifelse(is.null(input$bivar_dependent), "", input$bivar_dependent),
        ifelse(is.null(input$bivar_independent), "", paste(input$bivar_independent, collapse = ", ")),
        URL_DATA_TRAINING
      ),
      stringsAsFactors = FALSE
    )
  })
  
  export_current_excel <- function(file) {
    export_workbook(
      file = file,
      metadata_df = build_export_metadata(),
      filtered_df = filtered_data(),
      univ_df = univar_table_data(),
      bivar_df = bivar_table_data(),
      bivar_wide_df = bivar_wide_table_data(),
      univ_chart_df = univar_chart_data(),
      bivar_chart_df = bivar_chart_data(),
      assoc_observed_df = assoc_observed_data(),
      assoc_expected_df = assoc_expected_data(),
      assoc_residuals_df = assoc_residuals_data(),
      assoc_chisq_df = assoc_chisq_data(),
      assoc_nominal_df = assoc_nominal_data(),
      assoc_ordinal_df = assoc_ordinal_data()
    )
  }
  
  observeEvent(input$generate_univar_png, {
    result <- export_ggplot_static_png(
      plot_function = function() univar_plot_object(),
      prefix = "statcal_online_categorical_univariate_bar_chart",
      width = input$export_width,
      height = input$export_height,
      dpi = input$export_dpi,
      bg = safe_theme_bg(input$univar_theme)
    )
    univar_export_result(result)
  })
  
  observeEvent(input$generate_bivar_png, {
    result <- export_ggplot_static_png(
      plot_function = function() bivar_plot_object(),
      prefix = "statcal_online_categorical_bivariate_bar_chart",
      width = input$export_width,
      height = input$export_height,
      dpi = input$export_dpi,
      bg = safe_theme_bg(input$bivar_theme)
    )
    bivar_export_result(result)
  })
  
  observeEvent(input$generate_excel, {
    result <- export_excel_static_file(
      export_function = function(path) export_current_excel(path),
      prefix = "statcal_online_categorical_frequency_tables"
    )
    excel_export_result(result)
  })
  
  output$download_excel_fallback <- downloadHandler(
    filename = function() paste0("statcal_online_categorical_frequency_tables_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx"),
    contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    content = function(file) {
      export_current_excel(file)
    }
  )
}

shinyApp(ui = ui, server = server)
