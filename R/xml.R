# xml.R - structured table data -> minimal Word XML (OOXML / w: namespace)

# HELPERS

xml_escape <- function(text) {
  text <- gsub("&",  "&amp;",  text, fixed = TRUE)
  text <- gsub("<",  "&lt;",   text, fixed = TRUE)
  text <- gsub(">",  "&gt;",   text, fixed = TRUE)
  text <- gsub("\"", "&quot;", text, fixed = TRUE)
  text
}

# Returns a <w:tcBorders> block declaring only the specified sides as single lines.
# If sides is empty, returns "" - no border declaration needed (Word default = none).
cell_borders_xml <- function(sides = character()) {
  if (length(sides) == 0L) return("")
  tags <- vapply(sides, function(s) {
    sprintf("<w:%s w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>", s)
  }, character(1))
  sprintf("<w:tcBorders>%s</w:tcBorders>", paste(tags, collapse = ""))
}

# Run properties: font + size (+ optional bold)
run_pr_xml <- function(bold = FALSE) {
  b_tag <- if (bold) "<w:b/>" else ""
  sprintf(
    "<w:rPr>%s<w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\"/><w:sz w:val=\"20\"/><w:szCs w:val=\"20\"/></w:rPr>",
    b_tag
  )
}

# ROW BUILDERS

# Build a <w:tr> for a header row
# is_last_header: add bottom border to cells
xml_header_row <- function(row, is_last_header) {
  run_pr <- run_pr_xml(bold = TRUE)

  cells <- lapply(seq_along(row$cells), function(ci) {
    cell <- row$cells[[ci]]
    if (isTRUE(cell$colspan == 0L)) return("")  # continuation, omit

    grid_span <- if (!is.null(cell$colspan) && cell$colspan > 1L) {
      sprintf("<w:gridSpan w:val=\"%d\"/>", cell$colspan)
    } else ""

    # Last header row: always border. Upper rows: border only if cell has text.
    has_text <- nchar(trimws(cell$text)) > 0L
    sides    <- if (is_last_header || has_text) "bottom" else character()
    borders  <- cell_borders_xml(sides)

    align <- if (!is.na(cell$align)) cell$align else if (ci == 1L) "left" else "center"

    tc_pr <- sprintf(
      "<w:tcPr>%s%s<w:jc w:val=\"%s\"/></w:tcPr>",
      grid_span, borders, align
    )

    para_pr <- sprintf(
      "<w:pPr><w:jc w:val=\"%s\"/><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>",
      align
    )
    run  <- sprintf("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>", run_pr, xml_escape(cell$text))
    para <- sprintf("<w:p>%s%s</w:p>", para_pr, run)

    sprintf("<w:tc>%s%s</w:tc>", tc_pr, para)
  })

  tr_pr <- "<w:trPr><w:tblHeader/></w:trPr>"
  sprintf("<w:tr>%s%s</w:tr>", tr_pr, paste(cells, collapse = ""))
}

# Build a <w:tr> for a data row
# col_index: 1-based index of each cell (for alignment: 1 = left, rest = centre)
# is_last_data: add bottom border to cells
xml_data_row <- function(row, is_last_data) {
  borders <- cell_borders_xml(if (is_last_data) "bottom" else character())
  run_pr  <- run_pr_xml(bold = FALSE)

  cells <- lapply(seq_along(row$cells), function(ci) {
    cell  <- row$cells[[ci]]
    align <- if (!is.na(cell$align)) cell$align else if (ci == 1L) "left" else "center"

    tc_pr   <- sprintf("<w:tcPr>%s<w:jc w:val=\"%s\"/></w:tcPr>", borders, align)
    para_pr <- sprintf(
      "<w:pPr><w:jc w:val=\"%s\"/><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>",
      align
    )
    run     <- sprintf("<w:r>%s<w:t xml:space=\"preserve\">%s</w:t></w:r>", run_pr, xml_escape(cell$text))
    para    <- sprintf("<w:p>%s%s</w:p>", para_pr, run)

    sprintf("<w:tc>%s%s</w:tc>", tc_pr, para)
  })

  sprintf("<w:tr>%s</w:tr>", paste(cells, collapse = ""))
}

# CORE BUILDER

# Build OOXML <w:tbl> string from a combined table object
# cols: integer vector of 1-based column indices to include (NULL = all)
# row_start / row_end: 1-based data row range (NULL = all)
# text_width_twips: if supplied, scale column widths proportionally to this total
build_xml <- function(combined, cols = NULL, row_start = NULL, row_end = NULL,
                      text_width_twips = NULL) {
  header_rows      <- combined$header_rows
  data_rows        <- combined$data_rows
  col_widths_twips <- combined$col_widths_twips

  n_cols_total <- length(col_widths_twips)
  if (n_cols_total == 0L && length(header_rows) > 0L) {
    n_cols_total <- length(header_rows[[1]]$cells)
  }

  cols      <- resolve_cols(cols, n_cols_total, header_rows)
  data_rows <- slice_rows(data_rows, row_start, row_end)

  # Filter cells to selected columns
  filter_cells <- function(row) {
    row$cells <- row$cells[cols]
    row
  }
  header_rows <- lapply(header_rows, filter_cells)
  data_rows   <- lapply(data_rows,   filter_cells)

  # Column widths for the selected columns
  selected_widths <- col_widths_twips[cols]
  total_width     <- sum(selected_widths)

  # Scale proportionally to fill text_width_twips, preserving column ratios.
  # If no target width supplied (or source total is zero), use widths as-is.
  if (!is.null(text_width_twips) && total_width > 0L) {
    scale          <- text_width_twips / total_width
    selected_widths <- pmax(1L, round(selected_widths * scale))
    # Absorb any rounding remainder into the widest column
    diff <- text_width_twips - sum(selected_widths)
    if (diff != 0L) {
      widest <- which.max(selected_widths)
      selected_widths[widest] <- selected_widths[widest] + diff
    }
    total_width <- text_width_twips
  }

  # tblGrid - one gridCol per selected column (scaled widths in twips)
  grid_cols <- paste(
    sprintf("<w:gridCol w:w=\"%d\"/>", selected_widths),
    collapse = ""
  )
  tbl_grid <- sprintf("<w:tblGrid>%s</w:tblGrid>", grid_cols)

  # tblPr - explicit fixed width matching the text area
  tbl_pr <- sprintf(
    "<w:tblPr><w:tblW w:w=\"%d\" w:type=\"dxa\"/></w:tblPr>",
    total_width
  )

  # Header rows
  hdr_xml <- paste(lapply(seq_along(header_rows), function(i) {
    xml_header_row(header_rows[[i]], is_last_header = (i == length(header_rows)))
  }), collapse = "")

  # Data rows
  dat_xml <- paste(lapply(seq_along(data_rows), function(i) {
    xml_data_row(data_rows[[i]], is_last_data = (i == length(data_rows)))
  }), collapse = "")

  sprintf("<w:tbl>%s%s%s%s</w:tbl>", tbl_pr, tbl_grid, hdr_xml, dat_xml)
}

# PUBLIC FUNCTION (called from app.R / docx.R)

#' Generate Word XML for an RTF table
#'
#' @param path Path to the .rtf file
#' @param excluded_cols Integer vector of 1-based column indices to exclude (NULL = none)
#' @param excluded_rows Integer vector of 1-based data row indices to exclude (NULL = none)
#' @param excluded_header_rows Integer vector of 1-based header row indices to exclude
#' @param parameters Character vector of parameter values to keep (NULL = all)
#' @param timelines Character vector of timeline labels to keep (NULL = all)
#' @param text_width_twips Target table width in twips (NULL = use RTF widths)
#' @return Character string containing a <w:tbl> XML fragment
get_table_xml <- function(path,
                           excluded_cols        = NULL,
                           excluded_rows        = NULL,
                           excluded_header_rows = NULL,
                           parameters           = NULL,
                           timelines            = NULL,
                           text_width_twips     = NULL,
                           pages                = NULL) {
  prep <- prepare_table(
    pages = pages, path = path,
    excluded_cols = excluded_cols, excluded_rows = excluded_rows,
    excluded_header_rows = excluded_header_rows,
    parameters = parameters, timelines = timelines
  )
  build_xml(prep$combined, cols = prep$included_cols,
            text_width_twips = text_width_twips)
}
