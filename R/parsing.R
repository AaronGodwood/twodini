# parsing.R - RTF -> structured R data

# 1. LOW-LEVEL RTF HELPERS

# Read raw bytes and convert to UTF-8 string, handling CP1252/Latin-1
rtf_read_raw <- function(path) {
  lines <- tryCatch(
    readLines(path, encoding = "UTF-8",   warn = FALSE),
    error = function(e)
    readLines(path, encoding = "latin1",  warn = FALSE)
  )
  text <- paste(lines, collapse = "\n")
  # Normalise any residual \r\n or \r
  text <- gsub("\r\n", "\n", text, fixed = TRUE)
  gsub("\r",   "\n", text, fixed = TRUE)
}

# Resolve RTF escape sequences in a text string:
#   \'xx  -> the character for hex xx (in latin1, then to UTF-8)
#   \uN   -> Unicode code point N (followed by a fallback char we skip)
rtf_unescape <- function(text) {
  repeat {
    m <- regexpr("\\\\u(-?[0-9]+)\\??.", text, perl = TRUE)
    if (m == -1L) break
    n <- as.integer(sub("\\\\u(-?[0-9]+).*", "\\1",
                        substr(text, m, m + attr(m, "match.length") - 1L)))
    if (n < 0L) n <- n + 65536L
    text <- paste0(substr(text, 1L, m - 1L),
                   intToUtf8(n),
                   substr(text, m + attr(m, "match.length"), nchar(text)))
  }

  repeat {
    m <- regexpr("\\\\'([0-9a-fA-F]{2})", text, perl = TRUE)
    if (m == -1L) break
    hex <- substr(text, m + 2L, m + 3L)
    ch  <- rawToChar(as.raw(strtoi(hex, 16L)))
    ch  <- iconv(ch, from = "latin1", to = "UTF-8", sub = "?")
    text <- paste0(substr(text, 1L, m - 1L),
                   ch,
                   substr(text, m + 4L, nchar(text)))
  }

  text
}

# 2. PAGE SPLITTING

# Split RTF text into one character element per page.
# SAS uses \sectd to open each section and \sect to close it.
rtf_split_pages <- function(text) {
  # Split on bare \sect (not \sectd or any other longer \sect* word)
  parts <- strsplit(text, "\\\\sect(?![a-zA-Z])", perl = TRUE)[[1]]

  # Keep only chunks that contain \sectd (i.e. are real pages)
  parts <- parts[grepl("\\\\sectd", parts, perl = TRUE)]

  parts
}

# 3. GROUP EXTRACTION (brace-balanced)

# Find the content of the FIRST group whose opening is tagged with `tag`
# e.g. tag = "\\header" finds {\\header ...}
# Returns the inner text (without the outer braces) or NA if not found.
extract_group <- function(text, tag) {
  # Locate the tag
  m <- regexpr(paste0("\\", tag, "(?![a-zA-Z])"), text, perl = TRUE)
  if (m == -1L) return(NA_character_)

  start <- m[1]
  # Walk back to find the opening brace for this group
  # In RTF the pattern is always {<tag> ...} so the brace is just before tag
  # but sometimes there's no space: {\header
  brace_pos <- NA_integer_
  for (i in seq(start - 1L, max(1L, start - 5L), by = -1L)) {
    if (substr(text, i, i) == "{") { brace_pos <- i; break }
  }
  if (is.na(brace_pos)) return(NA_character_)

  # Walk forward counting braces until we close the group
  depth <- 0L
  chars <- strsplit(substr(text, brace_pos, nchar(text)), "")[[1]]
  end_pos <- brace_pos
  for (i in seq_along(chars)) {
    ch <- chars[i]
    if (ch == "{")  depth <- depth + 1L
    if (ch == "}") {
      depth <- depth - 1L
      if (depth == 0L) { end_pos <- brace_pos + i - 1L; break }
    }
  }

  substr(text, brace_pos + 1L, end_pos - 1L)
}

# Remove all groups tagged with `tag` from text
remove_group <- function(text, tag) {
  repeat {
    m <- regexpr(paste0("\\", tag, "(?![a-zA-Z])"), text, perl = TRUE)
    if (m == -1L) break

    start <- m[1]
    brace_pos <- NA_integer_
    for (i in seq(start - 1L, max(1L, start - 5L), by = -1L)) {
      if (substr(text, i, i) == "{") { brace_pos <- i; break }
    }
    if (is.na(brace_pos)) break

    depth <- 0L
    chars <- strsplit(substr(text, brace_pos, nchar(text)), "")[[1]]
    end_pos <- brace_pos
    for (i in seq_along(chars)) {
      ch <- chars[i]
      if (ch == "{")  depth <- depth + 1L
      if (ch == "}") {
        depth <- depth - 1L
        if (depth == 0L) { end_pos <- brace_pos + i - 1L; break }
      }
    }

    text <- paste0(substr(text, 1L, brace_pos - 1L),
                   substr(text, end_pos + 1L, nchar(text)))
  }
  text
}

# 4. TABLE PARSING

# Given RTF text for one section (header or body), parse all rows.
# Returns a list of row objects:
#   list(is_header, cells = list(list(text, width_twips, colspan)))
parse_rtf_table <- function(section_text) {
  rows <- list()

  # Tokenise: find each \trowd ... \row block
  # We'll scan linearly through \trowd markers
  trowd_positions <- gregexpr("\\\\trowd(?![a-zA-Z])", section_text, perl = TRUE)[[1]]
  if (identical(trowd_positions, -1L)) return(rows)

  row_positions   <- gregexpr("\\\\row(?![a-zA-Z])",   section_text, perl = TRUE)[[1]]
  if (identical(row_positions, -1L))   return(rows)

  for (ti in seq_along(trowd_positions)) {
    row_start <- trowd_positions[ti]

    # Find the matching \row that comes after this \trowd
    row_end_candidates <- row_positions[row_positions > row_start]
    if (length(row_end_candidates) == 0L) next
    row_end <- row_end_candidates[1]

    row_text <- substr(section_text, row_start, row_end + 3L)  # include \row

    # --- is_header: \trhdr present? ---
    is_hdr <- grepl("\\\\trhdr(?![a-zA-Z])", row_text, perl = TRUE)

    # --- cell boundary positions (\cellxN) ---
    cellx_matches <- gregexpr("\\\\cellx([0-9]+)", row_text, perl = TRUE)
    cellx_vals <- as.integer(regmatches(row_text, cellx_matches)[[1]] |>
                               gsub("\\\\cellx", "", x = _))

    # --- merge flags per cell position ---
    # \clmgf = first of merge, \clmrg = continuation
    # These appear in the row definition before \cellxN values
    # We pair them by order with cellx_vals
    clmgf_pos <- gregexpr("\\\\clmgf(?![a-zA-Z])", row_text, perl = TRUE)[[1]]
    clmrg_pos <- gregexpr("\\\\clmrg(?![a-zA-Z])",  row_text, perl = TRUE)[[1]]
    has_clmgf <- !identical(clmgf_pos, -1L)
    has_clmrg <- !identical(clmrg_pos, -1L)

    # Build cell definition list (one entry per \cellx)
    n_cells_def <- length(cellx_vals)
    cell_defs <- vector("list", n_cells_def)
    for (ci in seq_len(n_cells_def)) {
      w <- if (ci == 1L) cellx_vals[1] else cellx_vals[ci] - cellx_vals[ci - 1L]
      cell_defs[[ci]] <- list(width_twips = w, is_merge_first = FALSE, is_merge_cont = FALSE)
    }

    # Tag merge cells: find \clmgf / \clmrg occurrences and their nearest following \cellx
    # Simpler approach: find all cell-def blocks between \trowd and first \cell
    # Each block starts at a \clmgf or \clmrg before its \cellxN
    if (has_clmgf || has_clmrg) {
      # Find positions of all \cellx in the row_text
      cx_pos <- gregexpr("\\\\cellx[0-9]+", row_text, perl = TRUE)[[1]]
      if (!identical(cx_pos, -1L)) {
        for (ci in seq_along(cx_pos)) {
          # What's between previous cellx (or \trowd) and this cellx?
          prev <- if (ci == 1L) 1L else cx_pos[ci - 1L]
          seg  <- substr(row_text, prev, cx_pos[ci])
          cell_defs[[ci]]$is_merge_first <- grepl("\\\\clmgf(?![a-zA-Z])", seg, perl = TRUE)
          cell_defs[[ci]]$is_merge_cont  <- grepl("\\\\clmrg(?![a-zA-Z])",  seg, perl = TRUE)
        }
      }
    }

    # --- extract cell contents (\cell boundaries) ---
    # Split on \cell to get cell content chunks
    cell_chunks <- strsplit(row_text, "\\\\cell(?![a-zA-Z])", perl = TRUE)[[1]]
    # Last chunk is after the final \cell (contains \row etc.), discard it
    if (length(cell_chunks) > 1L) cell_chunks <- cell_chunks[-length(cell_chunks)]

    n_cells_content <- length(cell_chunks)

    # --- bold fallback for header detection ---
    if (!is_hdr && n_cells_content > 0L) {
      non_empty <- cell_chunks[nchar(trimws(cell_chunks)) > 0L]
      if (length(non_empty) > 0L) {
        all_bold <- all(sapply(non_empty, function(ch) {
          # Has \b (bold on) and no \b0 (bold off) after it
          grepl("\\\\b(?![a-zA-Z0-9])", ch, perl = TRUE) &&
            !grepl("\\\\b0(?![a-zA-Z])", ch, perl = TRUE)
        }))
        if (all_bold) is_hdr <- TRUE
      }
    }

    # --- clean cell text ---
    cells <- vector("list", n_cells_content)
    for (ci in seq_len(n_cells_content)) {
      raw_cell <- cell_chunks[ci]

      # Extract alignment from RTF control words before stripping
      cell_align <- if (grepl("\\\\ql(?![a-zA-Z])", raw_cell, perl = TRUE)) {
        "left"
      } else if (grepl("\\\\qr(?![a-zA-Z])", raw_cell, perl = TRUE)) {
        "right"
      } else if (grepl("\\\\qc(?![a-zA-Z])", raw_cell, perl = TRUE)) {
        "center"
      } else {
        NA_character_
      }

      # Strip RTF control words and groups, leaving plain text
      cell_text <- rtf_cell_to_text(raw_cell)

      w_twips <- if (ci <= length(cell_defs)) cell_defs[[ci]]$width_twips else 0L
      is_cont  <- ci <= length(cell_defs) && isTRUE(cell_defs[[ci]]$is_merge_cont)
      is_first <- ci <= length(cell_defs) && isTRUE(cell_defs[[ci]]$is_merge_first)

      cells[[ci]] <- list(
        text           = cell_text,
        width_twips    = w_twips,
        align          = cell_align,
        is_merge_first = is_first,
        is_merge_cont  = is_cont
      )
    }

    # --- resolve merged cells: repeat first-cell text into continuations ---
    if (n_cells_content > 1L) {
      last_first_text <- ""
      for (ci in seq_len(n_cells_content)) {
        if (isTRUE(cells[[ci]]$is_merge_first)) {
          last_first_text <- cells[[ci]]$text
        } else if (isTRUE(cells[[ci]]$is_merge_cont)) {
          cells[[ci]]$text <- last_first_text
        }
      }
    }

    # --- compute colspan for each cell ---
    # A merged group: is_merge_first followed by N is_merge_cont cells -> colspan = N+1
    if (n_cells_content > 0L) {
      ci <- 1L
      while (ci <= n_cells_content) {
        if (isTRUE(cells[[ci]]$is_merge_first)) {
          span <- 1L
          j <- ci + 1L
          while (j <= n_cells_content && isTRUE(cells[[j]]$is_merge_cont)) {
            span <- span + 1L
            j <- j + 1L
          }
          cells[[ci]]$colspan <- span
          # Mark continuation cells with colspan 0 (to be skipped in output)
          for (k in seq(ci + 1L, length.out = span - 1L)) {
            if (k <= n_cells_content) cells[[k]]$colspan <- 0L
          }
          ci <- j
        } else {
          cells[[ci]]$colspan <- 1L
          ci <- ci + 1L
        }
      }
    }

    rows <- c(rows, list(list(is_header = is_hdr, cells = cells)))
  }

  rows
}

# Strip RTF markup from a cell's raw text, returning clean plain text
rtf_cell_to_text <- function(raw) {
  # Remove nested groups (e.g. field instructions, pictures)
  text <- raw

  # Remove {\*...} destination groups entirely
  text <- gsub("\\{\\\\\\*[^}]*\\}", "", text, perl = TRUE)

  # Iteratively strip innermost {...} groups: remove control words inside, keep plain text.
  # Two-pass per iteration: first strip control words inside the group, then remove braces.
  for (i in seq_len(20L)) {
    # Strip RTF control words inside innermost groups (no nested braces)
    new_text <- gsub("\\{((?:[^{}\\\\]|\\\\[a-zA-Z]+[-]?[0-9]*[ ]?)*)\\}",
                     "\\1", text, perl = TRUE)
    # Strip any remaining control words that were the only content
    new_text <- gsub("\\{\\\\[a-zA-Z]+[-]?[0-9]*[ ]?\\}", "", new_text, perl = TRUE)
    if (identical(new_text, text)) break
    text <- new_text
  }

  # Remove remaining RTF control words (\word or \word123)
  text <- gsub("\\\\[a-zA-Z]+[-]?[0-9]*\\s?", "", text, perl = TRUE)
  # Remove remaining control symbols (\<symbol>)
  text <- gsub("\\\\.", "", text, perl = TRUE)
  # Remove stray braces
  text <- gsub("[{}]", "", text, fixed = FALSE)

  # Apply unicode/hex unescaping on what remains
  text <- rtf_unescape(text)

  trimws(text)
}

# 5. PAGE PARSING

parse_page <- function(page_text) {
  # Extract header and footer groups, then body is what remains
  header_text <- extract_group(page_text, "\\header")
  footer_text <- extract_group(page_text, "\\footer")  # not used currently

  body_text <- page_text |>
    remove_group("\\header")  |>
    remove_group("\\footer")  |>
    remove_group("\\headerf") |>
    remove_group("\\footerf") |>
    remove_group("\\headerl") |>
    remove_group("\\headerr") |>
    remove_group("\\footerl") |>
    remove_group("\\footerr")

  # Parse header table rows
  header_rows <- if (!is.na(header_text)) parse_rtf_table(header_text) else list()

  # Extract parameter from lowest header row
  parameter <- extract_parameter(header_rows)

  # Parse body table rows
  body_rows <- parse_rtf_table(body_text)

  # Split body rows into header rows (is_header=TRUE) and data rows
  is_hdr <- vapply(body_rows, `[[`, logical(1), "is_header")
  list(
    parameter   = parameter,
    header_rows = body_rows[is_hdr],
    data_rows   = body_rows[!is_hdr]
  )
}

# Extract "Parameter: <value>" from the lowest row of the RTF header section
extract_parameter <- function(header_rows) {
  if (length(header_rows) == 0L) return(NA_character_)

  # Check rows from bottom up
  for (i in rev(seq_along(header_rows))) {
    row <- header_rows[[i]]
    for (cell in row$cells) {
      m <- regmatches(cell$text,
                      regexpr("(?i)^Parameter:\\s*(.+)$", cell$text, perl = TRUE))
      if (length(m) > 0L && nchar(m) > 0L) {
        return(trimws(sub("(?i)^Parameter:\\s*", "", m, perl = TRUE)))
      }
    }
  }
  NA_character_
}

# 6. TOP-LEVEL PARSE FUNCTION

#' Parse an RTF file into a list of page objects
#'
#' @param path Path to the .rtf file
#' @return List of page objects, each with:
#'   \item{parameter}{character or NA}
#'   \item{header_rows}{list of row objects}
#'   \item{data_rows}{list of row objects}
#'   Each row object: list(is_header, cells = list(list(text, width_twips, colspan, ...)))
parse_rtf <- function(path) {
  text  <- rtf_read_raw(path)
  pages <- rtf_split_pages(text)
  lapply(pages, parse_page)
}

# 7. FILTERING

#' Filter pages by parameter value
#'
#' Pages with NA parameter are always kept.
#' @param pages Output of parse_rtf()
#' @param parameters Character vector of parameter values to include
filter_pages <- function(pages, parameters) {
  if (is.null(parameters) || length(parameters) == 0L) return(pages)
  pages[vapply(pages, function(p) {
    is.na(p$parameter) || p$parameter %in% parameters
  }, logical(1))]
}

#' Get all unique timeline labels across pages
#'
#' A timeline label is a non-empty value in column 1 of a data row where
#' all other cells in that row are empty. This matches the SAS pattern of
#' "Week 1" etc. appearing as a spanning label row.
#' @param pages Output of parse_rtf()
#' @return Character vector of unique timeline labels
get_timelines <- function(pages) {
  labels <- character()
  for (page in pages) {
    for (row in page$data_rows) {
      if (length(row$cells) == 0L) next
      col1 <- trimws(row$cells[[1]]$text)
      if (nchar(col1) == 0L) next
      # Check all other cells are empty
      rest_empty <- all(vapply(row$cells[-1], function(cell) nchar(trimws(cell$text)) == 0L, logical(1)))
      if (rest_empty) labels <- c(labels, col1)
    }
  }
  unique(labels)
}

#' Get all unique parameter values across pages (excluding NA)
#'
#' @param pages Output of parse_rtf()
#' @return Character vector of unique parameter values
get_parameters <- function(pages) {
  params <- vapply(pages, `[[`, character(1), "parameter")
  unique(params[!is.na(params)])
}

#' Filter data rows within pages by timeline labels
#'
#' Rows belonging to an excluded timeline block are dropped.
#' Header rows are never affected.
#' @param pages Output of parse_rtf() (possibly already filtered by parameter)
#' @param timelines Character vector of timeline labels to INCLUDE
filter_timelines <- function(pages, timelines) {
  if (is.null(timelines) || length(timelines) == 0L) return(pages)

  all_labels <- get_timelines(pages)
  # If no timeline labels exist in the data, return unchanged
  if (length(all_labels) == 0L) return(pages)

  lapply(pages, function(page) {
    current_timeline <- NA_character_
    kept <- logical(length(page$data_rows))

    for (i in seq_along(page$data_rows)) {
      row  <- page$data_rows[[i]]
      col1 <- if (length(row$cells) > 0L) trimws(row$cells[[1]]$text) else ""

      # Check if this row is a timeline label row
      rest_empty <- length(row$cells) <= 1L ||
        all(vapply(row$cells[-1], function(cell) nchar(trimws(cell$text)) == 0L, logical(1)))

      if (nchar(col1) > 0L && rest_empty && col1 %in% all_labels) {
        current_timeline <- col1
      }

      # Keep if: no timeline context yet, or current timeline is in the include list
      kept[i] <- is.na(current_timeline) || current_timeline %in% timelines
    }

    page$data_rows <- page$data_rows[kept]
    page
  })
}

# 8. COMBINING PAGES

#' Combine multiple pages into a single flat table
#'
#' Header rows are taken from the first page only.
#' Data rows are concatenated across all pages.
#' @param pages Filtered list of page objects
#' @return list(header_rows, data_rows, col_widths_twips)
combine_pages <- function(pages) {
  if (length(pages) == 0L) {
    return(list(header_rows = list(), data_rows = list(), col_widths_twips = numeric()))
  }

  header_rows <- pages[[1]]$header_rows
  data_rows   <- unlist(lapply(pages, `[[`, "data_rows"), recursive = FALSE)

  # Column widths from first page's first data (or header) row
  ref_rows <- if (length(header_rows) > 0L) header_rows else data_rows
  col_widths_twips <- if (length(ref_rows) > 0L) {
    vapply(ref_rows[[1]]$cells, `[[`, numeric(1), "width_twips")
  } else {
    numeric()
  }

  list(
    header_rows      = header_rows,
    data_rows        = data_rows,
    col_widths_twips = col_widths_twips
  )
}

# 9. PUBLIC INFO FUNCTION (used by app.R)

#' Get summary info about an RTF file for the UI
#'
#' @param path Path to the .rtf file
#' @return list(n_cols, n_rows, col_names, parameters, timelines)
get_table_info <- function(path) {
  pages <- parse_rtf(path)
  combined <- combine_pages(pages)

  ref_row <- if (length(combined$header_rows) > 0L) {
    combined$header_rows[[1]]
  } else if (length(combined$data_rows) > 0L) {
    combined$data_rows[[1]]
  } else {
    NULL
  }

  n_cols    <- if (!is.null(ref_row)) length(ref_row$cells) else 0L
  col_names <- if (!is.null(ref_row)) {
    vapply(ref_row$cells, function(cell) cell$text, character(1))
  } else {
    character()
  }

  list(
    n_cols     = n_cols,
    n_rows     = length(combined$data_rows),
    col_names  = col_names,
    parameters = get_parameters(pages),
    timelines  = get_timelines(pages)
  )
}

# 10. IMAGE EXTRACTION

#' Check whether an RTF file contains an embedded PNG image
#'
#' @param path Path to the .rtf file
#' @return TRUE if a \\pngblip group is found, FALSE otherwise
is_image_rtf <- function(path) {
  # Read only enough to find the marker - images appear early
  # Use raw text scan rather than full parse for speed
  text <- rtf_read_raw(path)
  grepl("\\\\pngblip(?![a-zA-Z])", text, perl = TRUE)
}

#' Extract the first PNG image from an RTF file
#'
#' Locates the first {\\pict ... \\pngblip ... <hex>} group, decodes the
#' hex-encoded bytes to raw, and returns the image data plus its declared
#' dimensions in twips (\\picwgoal / \\pichgoal).
#'
#' @param path Path to the .rtf file
#' @return list(png_bytes = raw, width_twips = integer, height_twips = integer)
#'   or NULL if no PNG found
extract_png <- function(path) {
  text <- rtf_read_raw(path)

  # Find the \pict group containing \pngblip
  pict_pos <- regexpr("\\\\pict(?![a-zA-Z])", text, perl = TRUE)
  if (pict_pos == -1L) return(NULL)

  # Walk back to opening brace
  brace_pos <- NA_integer_
  for (i in seq(pict_pos - 1L, max(1L, pict_pos - 5L), by = -1L)) {
    if (substr(text, i, i) == "{") { brace_pos <- i; break }
  }
  if (is.na(brace_pos)) return(NULL)

  # Find the closing brace of the \pict group.
  # The hex data in a \pict group contains no braces, so we only need to count
  # braces in the short header region (control words before the hex blob).
  # Scan forward in small chunks until depth returns to 0.
  chunk_size <- 4096L
  depth      <- 0L
  end_pos    <- NA_integer_
  pos        <- brace_pos
  text_len   <- nchar(text)
  while (pos <= text_len) {
    chunk_end <- min(pos + chunk_size - 1L, text_len)
    chars     <- strsplit(substr(text, pos, chunk_end), "")[[1]]
    for (i in seq_along(chars)) {
      if (chars[i] == "{") depth <- depth + 1L
      if (chars[i] == "}") {
        depth <- depth - 1L
        if (depth == 0L) { end_pos <- pos + i - 1L; break }
      }
    }
    if (!is.na(end_pos)) break
    pos <- pos + chunk_size
  }
  if (is.na(end_pos)) return(NULL)

  pict_text <- substr(text, brace_pos, end_pos)

  # Must contain \pngblip
  if (!grepl("\\\\pngblip(?![a-zA-Z])", pict_text, perl = TRUE)) return(NULL)

  # Extract \picwgoal and \pichgoal
  w_match <- regmatches(pict_text, regexpr("\\\\picwgoal([0-9]+)", pict_text, perl = TRUE))
  h_match <- regmatches(pict_text, regexpr("\\\\pichgoal([0-9]+)", pict_text, perl = TRUE))
  width_twips  <- if (length(w_match) > 0L) as.integer(sub("\\\\picwgoal", "", w_match)) else NA_integer_
  height_twips <- if (length(h_match) > 0L) as.integer(sub("\\\\pichgoal", "", h_match)) else NA_integer_

  # Extract the hex data: it starts on the line after the last RTF control word.
  # RTF control words are \word or \wordN; the hex blob follows on its own line(s).
  # Strategy: find the position right after the last \controlword (with optional
  # numeric argument) in the pict group, then take the rest up to the closing brace.
  last_ctrl <- gregexpr("\\\\[a-zA-Z]+[-]?[0-9]*", pict_text, perl = TRUE)[[1]]
  if (identical(last_ctrl, -1L)) return(NULL)
  last_end <- last_ctrl[length(last_ctrl)] +
              attr(last_ctrl, "match.length")[length(last_ctrl)] - 1L
  hex_region <- substr(pict_text, last_end + 1L, nchar(pict_text))
  # Keep only hex characters, strip everything else (whitespace, braces)
  hex_clean <- gsub("[^0-9a-fA-F]", "", hex_region, perl = TRUE)
  if (nchar(hex_clean) == 0L) return(NULL)
  hex_match <- hex_clean  # reuse variable name for code below

  # Strip whitespace from hex string and decode to raw bytes
  hex_clean <- gsub("\\s", "", hex_match, perl = TRUE)
  if (nchar(hex_clean) %% 2L != 0L) return(NULL)

  n_bytes   <- nchar(hex_clean) %/% 2L
  byte_strs <- substring(hex_clean, seq(1L, by = 2L, length.out = n_bytes),
                                     seq(2L, by = 2L, length.out = n_bytes))
  png_bytes <- as.raw(strtoi(byte_strs, 16L))

  list(
    png_bytes    = png_bytes,
    width_twips  = width_twips,
    height_twips = height_twips
  )
}
