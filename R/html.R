# html.R - structured table data -> minimal HTML

# HELPERS

htmlEscape <- function(text) {
  text <- gsub("&", "&amp;", text, fixed = TRUE)
  text <- gsub("<", "&lt;",  text, fixed = TRUE)
  text <- gsub(">", "&gt;",  text, fixed = TRUE)
  text
}

# Convert twips to px (96 dpi: 1 inch = 1440 twips = 96 px => 1 twip = 96/1440)
twips_to_px <- function(twips) round(twips * 96 / 1440)

# Build a single <tr> string for a header row
# is_last_header: whether to add bottom border to cells
html_header_row <- function(row, is_last_header) {
  cells <- vapply(seq_along(row$cells), function(ci) {
    cell <- row$cells[[ci]]
    if (isTRUE(cell$colspan == 0L)) return("")   # continuation cell, skip
    span_attr <- if (!is.null(cell$colspan) && cell$colspan > 1L) {
      sprintf(" colspan=\"%d\"", cell$colspan)
    } else ""
    has_text <- nchar(trimws(cell$text)) > 0L
    border   <- if (is_last_header || has_text) "border-bottom:2px solid black;" else ""
    raw_align <- if (!is.null(cell$align) && !is.na(cell$align)) cell$align else if (ci == 1L) "left" else "center"
    align    <- paste0("text-align:", raw_align)
    style    <- sprintf(" style=\"%spadding:0;%s\"", border, align)
    sprintf("<th%s%s><b>%s</b></th>", span_attr, style, htmlEscape(cell$text))
  }, character(1))
  paste0("<tr>", paste(cells, collapse = ""), "</tr>")
}

# Build a single <tr> string for a data row
# n_cols: total columns (for alignment: col 1 left, rest centre)
# is_last_data: whether to add bottom border to cells
html_data_row <- function(row, n_cols, is_last_data) {
  border_style <- if (is_last_data) "border-bottom:2px solid black;" else ""
  cells <- lapply(seq_along(row$cells), function(ci) {
    cell <- row$cells[[ci]]
    raw_align <- if (!is.null(cell$align) && !is.na(cell$align)) cell$align else if (ci == 1L) "left" else "center"
    align <- paste0("text-align:", raw_align)
    style <- trimws(paste0(border_style, "padding:0;", align), which = "left")
    sprintf("<td style=\"%s\">%s</td>", style, htmlEscape(cell$text))
  })
  paste0("<tr>", paste(cells, collapse = ""), "</tr>")
}

# CORE BUILDER

# Build HTML from a combined table object (output of combine_pages())
# cols: integer vector of column indices to include (NULL = all)
# row_start / row_end: 1-based data row range (NULL = all)
build_html <- function(combined, cols = NULL, row_start = NULL, row_end = NULL) {
  header_rows      <- combined$header_rows
  data_rows        <- combined$data_rows
  col_widths_twips <- combined$col_widths_twips

  n_cols_total <- length(col_widths_twips)
  if (n_cols_total == 0L && length(header_rows) > 0L) {
    n_cols_total <- length(header_rows[[1]]$cells)
  }

  cols      <- resolve_cols(cols, n_cols_total, header_rows)
  data_rows <- slice_rows(data_rows, row_start, row_end)

  # Helper: filter cells in a row to selected cols
  filter_cells <- function(row) {
    row$cells <- row$cells[cols]
    row
  }

  header_rows <- lapply(header_rows, filter_cells)
  data_rows   <- lapply(data_rows,   filter_cells)

  n_cols <- length(cols)

  # Column widths
  widths_px <- twips_to_px(col_widths_twips[cols])
  col_tags  <- paste(sprintf("<col style=\"width:%dpx\">", widths_px), collapse = "")
  colgroup  <- sprintf("<colgroup>%s</colgroup>", col_tags)

  # <thead>
  thead_rows <- lapply(seq_along(header_rows), function(i) {
    html_header_row(header_rows[[i]], is_last_header = (i == length(header_rows)))
  })
  thead <- sprintf("<thead>%s</thead>", paste(thead_rows, collapse = ""))

  # <tbody>
  n_data_rows  <- length(data_rows)
  row_limit    <- 200L
  truncated    <- n_data_rows > row_limit
  display_rows <- if (truncated) data_rows[seq_len(row_limit)] else data_rows

  tbody_rows <- lapply(seq_along(display_rows), function(i) {
    html_data_row(display_rows[[i]], n_cols, is_last_data = (!truncated && i == length(display_rows)))
  })

  if (truncated) {
    notice <- sprintf(
      "<tr><td colspan=\"%d\" style=\"padding:4px 0;text-align:center;font-style:italic;color:#888;border-top:1px solid #ccc;\">%d rows not shown - table truncated for preview</td></tr>",
      n_cols, n_data_rows - row_limit
    )
    tbody_rows <- c(tbody_rows, list(notice))
  }

  tbody <- sprintf("<tbody>%s</tbody>", paste(tbody_rows, collapse = ""))

  # Assemble - font defined once on the table element
  sprintf(
    "<table style=\"font-family:'Times New Roman',Times,serif;font-size:10pt;border-collapse:collapse\">%s%s%s</table>",
    colgroup, thead, tbody
  )
}

# SELECTION PANE BUILDER

# Build HTML for the interactive "selection" pane.
# Adds data-col / data-row / data-rowtype attributes for JS click handling.
# excluded_cols / excluded_rows / excluded_header_rows: integer vectors (1-based)
# - matching elements are rendered at opacity 0.3.
build_html_selection <- function(combined,
                                 excluded_cols        = integer(),
                                 excluded_rows        = integer(),
                                 excluded_header_rows = integer()) {
  header_rows      <- combined$header_rows
  data_rows        <- combined$data_rows
  col_widths_twips <- combined$col_widths_twips

  n_cols_total <- length(col_widths_twips)
  if (n_cols_total == 0L && length(header_rows) > 0L)
    n_cols_total <- length(header_rows[[1]]$cells)

  if (is.null(excluded_cols))        excluded_cols        <- integer()
  if (is.null(excluded_rows))        excluded_rows        <- integer()
  if (is.null(excluded_header_rows)) excluded_header_rows <- integer()

  widths_px <- twips_to_px(col_widths_twips)
  col_tags  <- paste(sprintf("<col style=\"width:%dpx\">", widths_px), collapse = "")
  colgroup  <- sprintf("<colgroup>%s</colgroup>", col_tags)

  # Header rows
  thead_rows <- lapply(seq_along(header_rows), function(ri) {
    row      <- header_rows[[ri]]
    is_last  <- ri == length(header_rows)
    row_excl <- ri %in% excluded_header_rows

    phys_col <- 1L
    cells <- character(length(row$cells))
    for (ci in seq_along(row$cells)) {
      cell <- row$cells[[ci]]
      if (isTRUE(cell$colspan == 0L)) {
        cells[ci] <- ""
        next
      }
      span    <- if (!is.null(cell$colspan) && cell$colspan > 1L) cell$colspan else 1L
      spanned <- seq(phys_col, phys_col + span - 1L)
      col_excl <- any(spanned %in% excluded_cols)

      data_col_val <- if (span > 1L) {
        paste0("[", paste(spanned, collapse = ","), "]")
      } else {
        as.character(phys_col)
      }
      phys_col <- phys_col + span

      span_attr <- if (span > 1L) sprintf(" colspan=\"%d\"", span) else ""
      has_text  <- nchar(trimws(cell$text)) > 0L
      border    <- if (is_last || has_text) "border-bottom:2px solid black;" else ""
      raw_align <- if (!is.null(cell$align) && !is.na(cell$align)) cell$align else
                   if (ci == 1L) "left" else "center"
      opacity   <- if (col_excl || row_excl) "opacity:0.3;" else ""
      style     <- sprintf(
        " style=\"%s%spadding:0;text-align:%s;cursor:pointer\"",
        border, opacity, raw_align
      )
      cells[ci] <- sprintf(
        "<th%s%s data-col='%s'><b>%s</b></th>",
        span_attr, style, data_col_val, htmlEscape(cell$text)
      )
    }

    row_opacity <- if (row_excl) " style=\"opacity:0.3\"" else ""
    sprintf(
      "<tr data-row=\"%d\" data-rowtype=\"header\"%s>%s</tr>",
      ri, row_opacity, paste(cells, collapse = "")
    )
  })
  thead <- sprintf("<thead>%s</thead>", paste(thead_rows, collapse = ""))

  # Data rows
  n_data_rows  <- length(data_rows)
  row_limit    <- 200L
  truncated    <- n_data_rows > row_limit
  display_rows <- if (truncated) data_rows[seq_len(row_limit)] else data_rows

  tbody_rows <- lapply(seq_along(display_rows), function(ri) {
    row      <- display_rows[[ri]]
    is_last  <- !truncated && ri == length(display_rows)
    row_excl <- ri %in% excluded_rows
    border_style <- if (is_last) "border-bottom:2px solid black;" else ""

    phys_col <- 1L
    cells <- character(length(row$cells))
    for (ci in seq_along(row$cells)) {
      cell     <- row$cells[[ci]]
      col_excl <- phys_col %in% excluded_cols
      opacity  <- if (col_excl || row_excl) "opacity:0.3;" else ""
      raw_align <- if (!is.null(cell$align) && !is.na(cell$align)) cell$align else
                   if (ci == 1L) "left" else "center"
      style <- sprintf(
        " style=\"%s%spadding:0;text-align:%s\"",
        border_style, opacity, raw_align
      )
      dc       <- as.character(phys_col)
      phys_col <- phys_col + 1L
      cells[ci] <- sprintf(
        "<td%s data-col='%s'>%s</td>", style, dc, htmlEscape(cell$text)
      )
    }

    row_style <- sprintf(
      " style=\"cursor:pointer%s\"",
      if (row_excl) ";opacity:0.3" else ""
    )
    sprintf(
      "<tr data-row=\"%d\" data-rowtype=\"data\"%s>%s</tr>",
      ri, row_style, paste(cells, collapse = "")
    )
  })

  if (truncated) {
    notice <- sprintf(
      "<tr><td colspan=\"%d\" style=\"padding:4px 0;text-align:center;font-style:italic;color:#888;border-top:1px solid #ccc;\">%d rows not shown - table truncated for preview</td></tr>",
      n_cols_total, n_data_rows - row_limit
    )
    tbody_rows <- c(tbody_rows, list(notice))
  }

  tbody <- sprintf("<tbody>%s</tbody>", paste(tbody_rows, collapse = ""))

  sprintf(
    paste0("<table style=\"font-family:'Times New Roman',Times,serif;",
           "font-size:10pt;border-collapse:collapse;user-select:none\">",
           "%s%s%s</table>"),
    colgroup, thead, tbody
  )
}

# PUBLIC FUNCTIONS (called from app.R)

#' Generate selection-pane HTML (interactive, data attributes, greying)
#'
#' @param path Path to the .rtf file
#' @param excluded_cols Integer vector of 1-based column indices to grey out
#' @param excluded_rows Integer vector of 1-based data row indices to grey out
#' @param excluded_header_rows Integer vector of 1-based header row indices
#' @param parameters Character vector of parameter values to keep (NULL = all)
#' @param timelines Character vector of timeline labels to keep (NULL = all)
#' @return HTML string
get_table_html_selection <- function(path,
                                     excluded_cols        = NULL,
                                     excluded_rows        = NULL,
                                     excluded_header_rows = NULL,
                                     parameters           = NULL,
                                     timelines            = NULL,
                                     pages                = NULL) {
  if (is.null(pages)) pages <- parse_rtf(path)
  combined <- pages |>
    filter_pages(parameters) |>
    filter_timelines(timelines) |>
    combine_pages()
  build_html_selection(
    combined,
    excluded_cols        = excluded_cols        %||% integer(),
    excluded_rows        = excluded_rows        %||% integer(),
    excluded_header_rows = excluded_header_rows %||% integer()
  )
}

#' Generate output-pane HTML (excluded cols/rows fully hidden)
#'
#' @param path Path to the .rtf file
#' @param excluded_cols Integer vector of 1-based column indices to remove
#' @param excluded_rows Integer vector of 1-based data row indices to remove
#' @param excluded_header_rows Integer vector of 1-based header row indices
#' @param parameters Character vector of parameter values to keep (NULL = all)
#' @param timelines Character vector of timeline labels to keep (NULL = all)
#' @return HTML string
get_table_html_output <- function(path,
                                  excluded_cols        = NULL,
                                  excluded_rows        = NULL,
                                  excluded_header_rows = NULL,
                                  parameters           = NULL,
                                  timelines            = NULL,
                                  pages                = NULL) {
  prep <- prepare_table(
    pages = pages, path = path,
    excluded_cols = excluded_cols, excluded_rows = excluded_rows,
    excluded_header_rows = excluded_header_rows,
    parameters = parameters, timelines = timelines
  )
  build_html(prep$combined, cols = prep$included_cols)
}

#' Generate full HTML preview for an RTF file
#'
#' @param path Path to the .rtf file
#' @return HTML string
get_table_html <- function(path, pages = NULL) {
  if (is.null(pages)) pages <- parse_rtf(path)
  combined <- combine_pages(pages)
  build_html(combined)
}
