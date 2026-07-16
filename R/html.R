# html.R - structured table data (blocks) -> minimal HTML

# HELPERS

htmlEscape <- function(text) {
  text <- gsub("&", "&amp;", text, fixed = TRUE)
  text <- gsub("<", "&lt;",  text, fixed = TRUE)
  text <- gsub(">", "&gt;",  text, fixed = TRUE)
  text
}

# Convert twips to px (96 dpi: 1 inch = 1440 twips = 96 px => 1 twip = 96/1440)
twips_to_px <- function(twips) round(twips * 96 / 1440)

# Resolve declared alignment with the positional fallback:
# column 1 is a label (left), the rest are data (center)
cell_align_or_default <- function(align, ci) {
  if (!is.na(align)) align else if (ci == 1L) "left" else "center"
}

# Build a single <tr> string for row `r` of a header block
# is_last_header: whether to add bottom border to cells
html_header_row <- function(b, r, is_last_header) {
  k <- block_ncol(b)
  cells <- character(k)
  for (ci in seq_len(k)) {
    if (!b$present[r, ci]) next
    span <- b$colspan[r, ci]
    if (span == 0L) next                        # continuation cell, skip
    span_attr <- if (span > 1L) sprintf(" colspan=\"%d\"", span) else ""
    has_text <- nzchar(trimws(b$text[r, ci]))
    border   <- if (is_last_header || has_text) "border-bottom:2px solid black;" else ""
    align    <- paste0("text-align:", cell_align_or_default(b$align[r, ci], ci))
    style    <- sprintf(" style=\"%spadding:0;%s\"", border, align)
    cells[ci] <- sprintf("<th%s%s><b>%s</b></th>",
                         span_attr, style, htmlEscape(b$text[r, ci]))
  }
  paste0("<tr>", paste(cells, collapse = ""), "</tr>")
}

# Build a single <tr> string for row `r` of a data block
# is_last_data: whether to add bottom border to cells
html_data_row <- function(b, r, is_last_data) {
  border_style <- if (is_last_data) "border-bottom:2px solid black;" else ""
  k <- block_ncol(b)
  cells <- character(k)
  for (ci in seq_len(k)) {
    if (!b$present[r, ci]) next
    span <- b$colspan[r, ci]
    if (span == 0L) next                        # merge continuation, skip
    span_attr <- if (span > 1L) sprintf(" colspan=\"%d\"", span) else ""
    align <- paste0("text-align:", cell_align_or_default(b$align[r, ci], ci))
    style <- trimws(paste0(border_style, "padding:0;", align), which = "left")
    cells[ci] <- sprintf("<td%s style=\"%s\">%s</td>",
                         span_attr, style, htmlEscape(b$text[r, ci]))
  }
  paste0("<tr>", paste(cells, collapse = ""), "</tr>")
}

# CORE BUILDER

# Build HTML from a combined table object (output of combine_pages())
# cols: integer vector of column indices to include (NULL = all)
# row_start / row_end: 1-based data row range (NULL = all)
build_html <- function(combined, cols = NULL, row_start = NULL, row_end = NULL) {
  header           <- combined$header
  data             <- combined$data
  col_widths_twips <- combined$col_widths_twips

  n_cols_total <- length(col_widths_twips)
  if (n_cols_total == 0L && block_nrow(header) > 0L) {
    n_cols_total <- sum(header$present[1L, ])
  }

  cols <- resolve_cols(cols, n_cols_total, header)
  data <- block_rows(data, slice_range(block_nrow(data), row_start, row_end))

  # Filter to selected columns
  header <- block_cols(header, cols)
  data   <- block_cols(data, cols)

  n_cols <- length(cols)

  # Column widths
  widths_px <- twips_to_px(col_widths_twips[cols])
  col_tags  <- paste(sprintf("<col style=\"width:%dpx\">", widths_px), collapse = "")
  colgroup  <- sprintf("<colgroup>%s</colgroup>", col_tags)

  # <thead>
  n_hdr <- block_nrow(header)
  thead_rows <- vapply(seq_len(n_hdr), function(i) {
    html_header_row(header, i, is_last_header = (i == n_hdr))
  }, character(1))
  thead <- sprintf("<thead>%s</thead>", paste(thead_rows, collapse = ""))

  # <tbody>
  n_data_rows <- block_nrow(data)
  row_limit   <- 200L
  truncated   <- n_data_rows > row_limit
  n_display   <- if (truncated) row_limit else n_data_rows

  tbody_rows <- vapply(seq_len(n_display), function(i) {
    html_data_row(data, i, is_last_data = (!truncated && i == n_display))
  }, character(1))

  if (truncated) {
    notice <- sprintf(
      "<tr><td colspan=\"%d\" style=\"padding:4px 0;text-align:center;font-style:italic;color:#888;border-top:1px solid #ccc;\">%d rows not shown - table truncated for preview</td></tr>",
      n_cols, n_data_rows - row_limit
    )
    tbody_rows <- c(tbody_rows, notice)
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
# Data rows carry their stable row ID in data-row (survives filter changes);
# header rows are positional.
# excluded_cols / excluded_header_rows: integer vectors (1-based positions);
# excluded_rows: stable row IDs. Matching elements render at opacity 0.3.
build_html_selection <- function(combined,
                                 excluded_cols        = integer(),
                                 excluded_rows        = integer(),
                                 excluded_header_rows = integer()) {
  header           <- combined$header
  data             <- combined$data
  col_widths_twips <- combined$col_widths_twips

  n_cols_total <- length(col_widths_twips)
  if (n_cols_total == 0L && block_nrow(header) > 0L) {
    n_cols_total <- sum(header$present[1L, ])
  }

  if (is.null(excluded_cols))        excluded_cols        <- integer()
  if (is.null(excluded_rows))        excluded_rows        <- integer()
  if (is.null(excluded_header_rows)) excluded_header_rows <- integer()

  widths_px <- twips_to_px(col_widths_twips)
  col_tags  <- paste(sprintf("<col style=\"width:%dpx\">", widths_px), collapse = "")
  colgroup  <- sprintf("<colgroup>%s</colgroup>", col_tags)

  # Header rows
  n_hdr <- block_nrow(header)
  k_hdr <- block_ncol(header)
  thead_rows <- vapply(seq_len(n_hdr), function(ri) {
    is_last  <- ri == n_hdr
    row_excl <- ri %in% excluded_header_rows

    phys_col <- 1L
    cells <- character(k_hdr)
    for (ci in seq_len(k_hdr)) {
      if (!header$present[ri, ci]) next
      span <- header$colspan[ri, ci]
      if (span == 0L) next
      if (span < 1L) span <- 1L
      spanned  <- seq(phys_col, phys_col + span - 1L)
      col_excl <- any(spanned %in% excluded_cols)

      data_col_val <- if (span > 1L) {
        paste0("[", paste(spanned, collapse = ","), "]")
      } else {
        as.character(phys_col)
      }
      phys_col <- phys_col + span

      span_attr <- if (span > 1L) sprintf(" colspan=\"%d\"", span) else ""
      has_text  <- nzchar(trimws(header$text[ri, ci]))
      border    <- if (is_last || has_text) "border-bottom:2px solid black;" else ""
      raw_align <- cell_align_or_default(header$align[ri, ci], ci)
      opacity   <- if (col_excl || row_excl) "opacity:0.3;" else ""
      style     <- sprintf(
        " style=\"%s%spadding:0;text-align:%s;cursor:pointer\"",
        border, opacity, raw_align
      )
      cells[ci] <- sprintf(
        "<th%s%s data-col='%s'><b>%s</b></th>",
        span_attr, style, data_col_val, htmlEscape(header$text[ri, ci])
      )
    }

    row_opacity <- if (row_excl) " style=\"opacity:0.3\"" else ""
    sprintf(
      "<tr data-row=\"%d\" data-rowtype=\"header\"%s>%s</tr>",
      ri, row_opacity, paste(cells, collapse = "")
    )
  }, character(1))
  thead <- sprintf("<thead>%s</thead>", paste(thead_rows, collapse = ""))

  # Data rows
  n_data_rows <- block_nrow(data)
  k_dat       <- block_ncol(data)
  row_limit   <- 200L
  truncated   <- n_data_rows > row_limit
  n_display   <- if (truncated) row_limit else n_data_rows

  tbody_rows <- vapply(seq_len(n_display), function(ri) {
    is_last <- !truncated && ri == n_display
    # Data rows are identified by stable row ID (survives filter changes);
    # rows without one fall back to their display position
    rid      <- data$row_id[ri]
    if (is.na(rid)) rid <- ri
    row_excl <- rid %in% excluded_rows
    border_style <- if (is_last) "border-bottom:2px solid black;" else ""

    phys_col <- 1L
    cells <- character(k_dat)
    for (ci in seq_len(k_dat)) {
      if (!data$present[ri, ci]) next
      span <- data$colspan[ri, ci]
      if (span == 0L) next                     # merge continuation, skip
      if (span < 1L) span <- 1L
      spanned  <- seq(phys_col, phys_col + span - 1L)
      col_excl <- any(spanned %in% excluded_cols)

      dc <- if (span > 1L) {
        paste0("[", paste(spanned, collapse = ","), "]")
      } else {
        as.character(phys_col)
      }
      phys_col <- phys_col + span

      span_attr <- if (span > 1L) sprintf(" colspan=\"%d\"", span) else ""
      opacity   <- if (col_excl || row_excl) "opacity:0.3;" else ""
      raw_align <- cell_align_or_default(data$align[ri, ci], ci)
      style <- sprintf(
        " style=\"%s%spadding:0;text-align:%s\"",
        border_style, opacity, raw_align
      )
      cells[ci] <- sprintf(
        "<td%s%s data-col='%s'>%s</td>",
        span_attr, style, dc, htmlEscape(data$text[ri, ci])
      )
    }

    row_style <- sprintf(
      " style=\"cursor:pointer%s\"",
      if (row_excl) ";opacity:0.3" else ""
    )
    sprintf(
      "<tr data-row=\"%d\" data-rowtype=\"data\"%s>%s</tr>",
      rid, row_style, paste(cells, collapse = "")
    )
  }, character(1))

  if (truncated) {
    notice <- sprintf(
      "<tr><td colspan=\"%d\" style=\"padding:4px 0;text-align:center;font-style:italic;color:#888;border-top:1px solid #ccc;\">%d rows not shown - table truncated for preview</td></tr>",
      n_cols_total, n_data_rows - row_limit
    )
    tbody_rows <- c(tbody_rows, notice)
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
#' @param excluded_rows Integer vector of stable data-row IDs to grey out
#' @param excluded_header_rows Integer vector of 1-based header row indices
#' @param parameters Character vector of parameter values to keep (NULL = all)
#' @param timelines Character vector of timeline labels to keep (NULL = all)
#' @param pages Pre-parsed RTF pages (output of parse_rtf()); parsed from path if NULL
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
#' @param excluded_rows Integer vector of stable data-row IDs to remove
#' @param excluded_header_rows Integer vector of 1-based header row indices
#' @param parameters Character vector of parameter values to keep (NULL = all)
#' @param timelines Character vector of timeline labels to keep (NULL = all)
#' @param pages Pre-parsed RTF pages (output of parse_rtf()); parsed from path if NULL
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
#' @param pages Pre-parsed RTF pages (output of parse_rtf()); parsed from path if NULL
#' @return HTML string
get_table_html <- function(path, pages = NULL) {
  if (is.null(pages)) pages <- parse_rtf(path)
  combined <- combine_pages(pages)
  build_html(combined)
}
