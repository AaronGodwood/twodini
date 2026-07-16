# parsing.R - RTF -> structured R data

# Whether the compiled C fast-path is available. Resolved lazily on first use
# and cached: the DLL registered by useDynLib() is not guaranteed to be loaded
# at the moment this file is sourced during namespace construction, so an eager
# check here can spuriously report FALSE and silently disable the C path.
.c_available <- local({
  cached <- NA
  function() {
    if (!is.na(cached)) return(cached)
    cached <<- tryCatch({
      getNativeSymbolInfo("C_find_matching_brace", PACKAGE = "houdini")
      TRUE
    }, error = function(e) FALSE)
    cached
  }
})

# 1. LOW-LEVEL RTF HELPERS

# Read raw bytes and convert to UTF-8 string, handling CP1252/Latin-1.
# readLines(encoding=) never errors on wrong encodings, it just mis-marks the
# string - so detect real UTF-8 with validUTF8() and convert otherwise.
rtf_read_raw <- function(path) {
  bytes <- readBin(path, what = "raw", n = file.size(path))
  bytes <- bytes[bytes != as.raw(0L)]
  text  <- rawToChar(bytes)
  if (validUTF8(text)) {
    Encoding(text) <- "UTF-8"
  } else {
    # SAS/Word on Windows write CP1252 (latin1 plus smart quotes, dashes, ...)
    text <- iconv(text, from = "CP1252", to = "UTF-8", sub = "?")
  }
  # Normalise any residual \r\n or \r
  text <- gsub("\r\n", "\n", text, fixed = TRUE)
  gsub("\r",   "\n", text, fixed = TRUE)
}

# Resolve RTF escape sequences in a text string:
#   \'xx  -> the character for hex xx (CP1252, converted to UTF-8)
#   \uN   -> Unicode code point N, followed by \ucN fallback chars we skip
#   \ucN  -> sets the fallback length for subsequent \uN escapes (default 1);
#            consumed and not emitted
# A fallback "character" may itself be a \'xx escape (counts as one).
rtf_unescape_r <- function(text) {
  if (!grepl("\\\\(u|')", text, perl = TRUE)) return(text)

  n   <- nchar(text)
  out <- character()
  pos <- 1L
  uc  <- 1L

  repeat {
    m <- regexpr("\\\\(uc[0-9]+|u-?[0-9]+|'[0-9a-fA-F]{2})",
                 substr(text, pos, n), perl = TRUE)
    if (m == -1L) {
      out <- c(out, substr(text, pos, n))
      break
    }
    start <- pos + m[1L] - 1L
    len   <- attr(m, "match.length")
    out   <- c(out, substr(text, pos, start - 1L))
    tok   <- substr(text, start, start + len - 1L)
    pos   <- start + len

    if (startsWith(tok, "\\uc")) {
      uc <- as.integer(substr(tok, 4L, len))
      if (substr(text, pos, pos) == " ") pos <- pos + 1L  # delimiter
    } else if (startsWith(tok, "\\u")) {
      cp <- as.integer(substr(tok, 3L, len))
      if (cp < 0L) cp <- cp + 65536L
      out <- c(out, intToUtf8(cp))
      if (substr(text, pos, pos) == " ") pos <- pos + 1L  # delimiter
      # Skip the uc fallback characters; stop early at structure we shouldn't eat
      k <- uc
      while (k > 0L && pos <= n) {
        ch <- substr(text, pos, pos)
        if (ch == "\\") {
          if (substr(text, pos + 1L, pos + 1L) == "'") pos <- pos + 4L else break
        } else if (ch == "{" || ch == "}") {
          break
        } else {
          pos <- pos + 1L
        }
        k <- k - 1L
      }
    } else {
      # \'xx hex escape
      hex <- substr(tok, 3L, 4L)
      ch  <- rawToChar(as.raw(strtoi(hex, 16L)))
      out <- c(out, iconv(ch, from = "CP1252", to = "UTF-8", sub = "?"))
    }
  }

  paste(out, collapse = "")
}

rtf_unescape <- function(text) {
  if (.c_available()) .Call(C_rtf_unescape, text) else rtf_unescape_r(text)
}

# Fast hex string to raw vector conversion.
# Processes in chunks of 4000 bytes (8000 hex chars) to limit intermediate
# string allocation compared to one-pair-at-a-time substring calls.
hex_to_raw <- function(hex) {
  n <- nchar(hex) %/% 2L
  chunk <- 4000L
  parts <- vector("list", ceiling(n / chunk))
  pi <- 1L
  pos <- 1L
  while (pos <= n * 2L) {
    end <- min(pos + chunk * 2L - 1L, n * 2L)
    piece <- substr(hex, pos, end)
    nb <- nchar(piece) %/% 2L
    starts <- seq(1L, by = 2L, length.out = nb)
    parts[[pi]] <- as.raw(strtoi(substring(piece, starts, starts + 1L), 16L))
    pi <- pi + 1L
    pos <- end + 1L
  }
  unlist(parts)
}

# Find the position of the closing brace matching the opening brace at `start`.
# Converts to raw bytes for O(1) indexing instead of per-character substr calls.
find_matching_brace_r <- function(text, start) {
  bytes <- charToRaw(text)
  n <- length(bytes)
  open  <- charToRaw("{")
  close <- charToRaw("}")
  depth <- 0L
  pos <- start
  while (pos <= n) {
    b <- bytes[pos]
    if (b == open) depth <- depth + 1L
    else if (b == close) {
      depth <- depth - 1L
      if (depth == 0L) return(pos)
    }
    pos <- pos + 1L
  }
  NA_integer_
}

find_matching_brace <- function(text, start) {
  if (.c_available()) .Call(C_find_matching_brace, text, as.integer(start))
  else find_matching_brace_r(text, start)
}

# Remove all groups tagged with any of the given tags in a single pass.
# Avoids repeated full-text scans when multiple tags need stripping.
remove_groups <- function(text, tags) {
  for (tag in tags) {
    text <- remove_group(text, tag)
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
  m <- regexpr(paste0("\\", tag, "(?![a-zA-Z])"), text, perl = TRUE)
  if (m == -1L) return(NA_character_)

  start <- m[1]
  brace_pos <- NA_integer_
  for (i in seq(start - 1L, max(1L, start - 5L), by = -1L)) {
    if (substr(text, i, i) == "{") { brace_pos <- i; break }
  }
  if (is.na(brace_pos)) return(NA_character_)

  end_pos <- find_matching_brace(text, brace_pos)
  if (is.na(end_pos)) return(NA_character_)

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

    end_pos <- find_matching_brace(text, brace_pos)
    if (is.na(end_pos)) break

    text <- paste0(substr(text, 1L, brace_pos - 1L),
                   substr(text, end_pos + 1L, nchar(text)))
  }
  text
}

# 4. TABLE PARSING

# findInterval re-validates `vec` on every call (anyNA + is.unsorted + an
# as.double copy) - O(length(vec)) work that turns per-row lookups over large
# sections quadratic. Positions here are always sorted doubles, so skip the
# checks where this R version allows it (R >= 4.3).
fint <- if (getRversion() >= "4.3.0") {
  function(x, vec) findInterval(x, vec, checkSorted = FALSE, checkNA = FALSE)
} else {
  findInterval
}

# Start positions of all matches of `pattern`, ascending; numeric(0) if none.
# Doubles, not integers, so fint() avoids a per-call as.double copy.
token_positions <- function(text, pattern) {
  p <- gregexpr(pattern, text, perl = TRUE)[[1]]
  if (p[1L] == -1L) numeric() else as.double(p)
}

# Elements of the sorted position vector `pos` that fall within [lo, hi]
pos_within <- function(pos, lo, hi) {
  if (length(pos) == 0L) return(numeric())
  i1 <- fint(lo - 1L, pos) + 1L
  i2 <- fint(hi, pos)
  if (i1 > i2) numeric() else pos[i1:i2]
}

# Does any element of the sorted position vector `pos` fall within [lo, hi]?
any_within <- function(pos, lo, hi) {
  length(pos) != 0L && fint(hi, pos) > fint(lo - 1L, pos)
}

# Given RTF text for one section (header or body), parse all rows.
# Returns a list of row objects:
#   list(is_header, cells = list(list(text, width_twips, colspan)))
#
# Each token type is located with a single regex scan over the whole section;
# per-row and per-cell tests then become sorted-position lookups instead of
# fresh regex passes over row substrings.
parse_rtf_table <- function(section_text) {
  trowd_pos <- token_positions(section_text, "\\\\trowd(?![a-zA-Z])")
  if (length(trowd_pos) == 0L) return(list())

  row_pos <- token_positions(section_text, "\\\\row(?![a-zA-Z])")
  if (length(row_pos) == 0L) return(list())

  trhdr_pos    <- token_positions(section_text, "\\\\trhdr(?![a-zA-Z])")
  clmgf_pos    <- token_positions(section_text, "\\\\clmgf(?![a-zA-Z])")
  clmrg_pos    <- token_positions(section_text, "\\\\clmrg(?![a-zA-Z])")
  cell_pos     <- token_positions(section_text, "\\\\cell(?![a-zA-Z])")
  ql_pos       <- token_positions(section_text, "\\\\ql(?![a-zA-Z])")
  qr_pos       <- token_positions(section_text, "\\\\qr(?![a-zA-Z])")
  qc_pos       <- token_positions(section_text, "\\\\qc(?![a-zA-Z])")
  bold_on_pos  <- token_positions(section_text, "\\\\b(?![a-zA-Z0-9])")
  bold_off_pos <- token_positions(section_text, "\\\\b0(?![a-zA-Z])")

  cellx_m   <- gregexpr("\\\\cellx([0-9]+)", section_text, perl = TRUE)
  cellx_pos <- cellx_m[[1]]
  if (cellx_pos[1L] == -1L) {
    cellx_pos  <- numeric()
    cellx_vals <- integer()
  } else {
    cellx_pos  <- as.double(cellx_pos)
    cellx_vals <- as.integer(gsub("\\\\cellx", "",
                                  regmatches(section_text, cellx_m)[[1]]))
  }

  rows <- vector("list", length(trowd_pos))
  nr   <- 0L

  for (ti in seq_along(trowd_pos)) {
    row_start <- trowd_pos[ti]

    # First \row after this \trowd closes the row
    ri <- fint(row_start, row_pos) + 1L
    if (ri > length(row_pos)) next
    row_close <- row_pos[ri] + 3L  # last char of "\row"

    # --- is_header: \trhdr present? ---
    is_hdr <- any_within(trhdr_pos, row_start, row_close)

    # --- cell boundary positions (\cellxN) ---
    di1 <- fint(row_start - 1L, cellx_pos) + 1L
    di2 <- fint(row_close, cellx_pos)
    if (di1 <= di2) {
      cx_pos_row <- cellx_pos[di1:di2]
      cx_vals    <- cellx_vals[di1:di2]
    } else {
      cx_pos_row <- numeric()
      cx_vals    <- integer()
    }

    # Cell widths from successive \cellx differences
    n_cells_def <- length(cx_vals)
    widths <- if (n_cells_def > 0L) c(cx_vals[1L], diff(cx_vals)) else integer()

    # --- merge flags per cell position ---
    # \clmgf = first of merge, \clmrg = continuation; each belongs to the
    # next \cellx that follows it in the row definition
    merge_first <- logical(n_cells_def)
    merge_cont  <- logical(n_cells_def)
    for (p in pos_within(clmgf_pos, row_start, row_close)) {
      k <- fint(p, cx_pos_row) + 1L
      if (k <= n_cells_def) merge_first[k] <- TRUE
    }
    for (p in pos_within(clmrg_pos, row_start, row_close)) {
      k <- fint(p, cx_pos_row) + 1L
      if (k <= n_cells_def) merge_cont[k] <- TRUE
    }

    # --- extract cell contents (\cell boundaries) ---
    # Chunk i runs from just after the previous \cell (or the row start)
    # up to just before \cell i; text after the final \cell is discarded.
    # A row with no \cell yields one chunk spanning the whole row.
    cp <- pos_within(cell_pos, row_start, row_close)
    if (length(cp) > 0L) {
      chunk_starts <- c(row_start, cp[-length(cp)] + 5L)
      chunk_ends   <- cp - 1L
    } else {
      chunk_starts <- row_start
      chunk_ends   <- row_close
    }
    cell_chunks <- substring(section_text, chunk_starts, chunk_ends)

    n_cells_content <- length(cell_chunks)

    # Merge flags per content cell (FALSE beyond the defined cells)
    flag_first <- logical(n_cells_content)
    flag_cont  <- logical(n_cells_content)
    nd <- min(n_cells_content, n_cells_def)
    if (nd > 0L) {
      flag_first[seq_len(nd)] <- merge_first[seq_len(nd)]
      flag_cont[seq_len(nd)]  <- merge_cont[seq_len(nd)]
    }

    # --- bold fallback for header detection ---
    if (!is_hdr && n_cells_content > 0L) {
      # has any non-whitespace character (= nchar(trimws(x)) > 0, but one
      # vectorised regex call instead of trimws's two sub() calls per row)
      non_empty <- which(grepl("[^ \t\r\n]", cell_chunks, perl = TRUE))
      if (length(non_empty) > 0L) {
        all_bold <- TRUE
        for (ci in non_empty) {
          # Has \b (bold on) and no \b0 (bold off)
          if (!any_within(bold_on_pos, chunk_starts[ci], chunk_ends[ci]) ||
              any_within(bold_off_pos, chunk_starts[ci], chunk_ends[ci])) {
            all_bold <- FALSE
            break
          }
        }
        if (all_bold) is_hdr <- TRUE
      }
    }

    # --- clean cell text ---
    cells <- vector("list", n_cells_content)
    for (ci in seq_len(n_cells_content)) {
      cs <- chunk_starts[ci]
      ce <- chunk_ends[ci]

      # Alignment from RTF control words within the cell chunk
      cell_align <- if (any_within(ql_pos, cs, ce)) {
        "left"
      } else if (any_within(qr_pos, cs, ce)) {
        "right"
      } else if (any_within(qc_pos, cs, ce)) {
        "center"
      } else {
        NA_character_
      }

      cells[[ci]] <- list(
        text           = rtf_cell_to_text(cell_chunks[ci]),
        width_twips    = if (ci <= n_cells_def) widths[ci] else 0L,
        align          = cell_align,
        is_merge_first = flag_first[ci],
        is_merge_cont  = flag_cont[ci]
      )
    }

    if (any(flag_first) || any(flag_cont)) {
      # --- resolve merged cells: repeat first-cell text into continuations ---
      if (n_cells_content > 1L) {
        last_first_text <- ""
        for (ci in seq_len(n_cells_content)) {
          if (flag_first[ci]) {
            last_first_text <- cells[[ci]]$text
          } else if (flag_cont[ci]) {
            cells[[ci]]$text <- last_first_text
          }
        }
      }

      # --- compute colspan for each cell ---
      # A merged group: is_merge_first followed by N is_merge_cont cells -> colspan = N+1
      ci <- 1L
      while (ci <= n_cells_content) {
        if (flag_first[ci]) {
          span <- 1L
          j <- ci + 1L
          while (j <= n_cells_content && flag_cont[j]) {
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
    } else {
      # No merges in this row: every cell spans one column
      for (ci in seq_len(n_cells_content)) cells[[ci]]$colspan <- 1L
    }

    nr <- nr + 1L
    rows[[nr]] <- list(is_header = is_hdr, cells = cells)
  }

  rows[seq_len(nr)]
}

# Strip RTF markup from a cell's raw text, returning clean plain text
rtf_cell_to_text_r <- function(raw) {
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
    # (but not \uN / \ucN, which rtf_unescape decodes later)
    new_text <- gsub("\\{\\\\(?!u-?[0-9]|uc[0-9])[a-zA-Z]+[-]?[0-9]*[ ]?\\}", "", new_text, perl = TRUE)
    if (identical(new_text, text)) break
    text <- new_text
  }

  # Remove remaining RTF control words (\word or \word123), but preserve
  # \uN and \ucN - they are decoded (not stripped) by rtf_unescape below
  text <- gsub("\\\\(?!u-?[0-9]|uc[0-9])[a-zA-Z]+[-]?[0-9]*\\s?", "", text, perl = TRUE)
  # Remove remaining control symbols (\<symbol>), preserving \'xx hex escapes
  text <- gsub("\\\\(?!')[^a-zA-Z]", "", text, perl = TRUE)
  # Remove stray braces
  text <- gsub("[{}]", "", text, fixed = FALSE)

  # Apply unicode/hex unescaping on what remains
  text <- rtf_unescape(text)

  trimws(text)
}

rtf_cell_to_text <- function(raw) {
  if (.c_available()) .Call(C_rtf_cell_to_text, raw) else rtf_cell_to_text_r(raw)
}

# 5. PAGE PARSING

parse_page <- function(page_text) {
  # Extract header and footer groups, then body is what remains
  header_text <- extract_group(page_text, "\\header")
  footer_text <- extract_group(page_text, "\\footer")  # not used currently

  body_text <- remove_groups(page_text, c(
    "\\header", "\\footer", "\\headerf", "\\footerf",
    "\\headerl", "\\headerr", "\\footerl", "\\footerr"
  ))

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
  labels <- vector("list", length(pages))
  for (pi in seq_along(pages)) {
    page_labels <- character()
    for (row in pages[[pi]]$data_rows) {
      if (length(row$cells) == 0L) next
      col1 <- trimws(row$cells[[1]]$text)
      if (nchar(col1) == 0L) next
      rest_empty <- all(vapply(row$cells[-1], cell_effectively_empty, logical(1)))
      if (rest_empty) page_labels <- c(page_labels, col1)
    }
    labels[[pi]] <- page_labels
  }
  unique(unlist(labels, use.names = FALSE))
}

# A cell counts as empty for timeline-label detection if it has no text OR is
# a merge continuation (parse_rtf_table copies the first cell's text into
# continuations, so a label merged across the row would otherwise never match)
cell_effectively_empty <- function(cell) {
  isTRUE(cell$colspan == 0L) || nchar(trimws(cell$text)) == 0L
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
        all(vapply(row$cells[-1], cell_effectively_empty, logical(1)))

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

# 10. SHARED FILTER HELPERS (used by html.R and xml.R)

#' Resolve column specification to integer indices
#'
#' Handles NULL (all), integer vector, or character names matched against headers.
#' @param cols Column spec (NULL, integer vector, or character names)
#' @param n_cols_total Total number of columns
#' @param header_rows Header rows (for name matching)
#' @return Integer vector of valid 1-based column indices
resolve_cols <- function(cols, n_cols_total, header_rows) {
  if (is.null(cols) || length(cols) == 0L) return(seq_len(n_cols_total))

  cols_int <- suppressWarnings(as.integer(cols))
  if (anyNA(cols_int)) {
    ref_names <- if (length(header_rows) > 0L) {
      vapply(header_rows[[1]]$cells, function(c) c$text, character(1))
    } else character()
    cols_int <- match(as.character(cols), ref_names)
    cols_int <- cols_int[!is.na(cols_int)]
  }
  cols_int <- cols_int[cols_int >= 1L & cols_int <= n_cols_total]
  if (length(cols_int) == 0L) seq_len(n_cols_total) else cols_int
}

#' Slice data rows by start/end range
#'
#' @param data_rows List of row objects
#' @param row_start 1-based start (NULL = 1)
#' @param row_end 1-based end (NULL = last)
#' @return Sliced list of row objects
slice_rows <- function(data_rows, row_start = NULL, row_end = NULL) {
  n <- length(data_rows)
  if (n == 0L) return(list())
  rs <- if (is.null(row_start)) 1L else max(1L, as.integer(row_start))
  re <- if (is.null(row_end))   n   else min(n, as.integer(row_end))
  if (rs <= re) data_rows[seq(rs, re)] else list()
}

#' Prepare a combined table with filtering and exclusions applied
#'
#' Common pipeline used by html.R and xml.R public functions.
#' @param pages Pre-parsed pages (from parse_rtf) or NULL
#' @param path RTF file path (used only if pages is NULL)
#' @param excluded_cols Integer vector of column indices to exclude
#' @param excluded_rows Integer vector of data row indices to exclude
#' @param excluded_header_rows Integer vector of header row indices to exclude
#' @param parameters Parameter filter
#' @param timelines Timeline filter
#' @return list(combined, included_cols) where combined has rows filtered
prepare_table <- function(pages = NULL, path = NULL,
                          excluded_cols = NULL, excluded_rows = NULL,
                          excluded_header_rows = NULL,
                          parameters = NULL, timelines = NULL) {
  if (is.null(pages)) pages <- parse_rtf(path)
  pages    <- filter_pages(pages, parameters)
  pages    <- filter_timelines(pages, timelines)
  combined <- combine_pages(pages)

  n_cols <- length(combined$col_widths_twips)
  n_data <- length(combined$data_rows)
  n_hdr  <- length(combined$header_rows)

  inc_cols <- setdiff(seq_len(n_cols), excluded_cols        %||% integer())
  inc_rows <- setdiff(seq_len(n_data), excluded_rows        %||% integer())
  inc_hdrs <- setdiff(seq_len(n_hdr),  excluded_header_rows %||% integer())

  combined$data_rows   <- combined$data_rows[inc_rows]
  combined$header_rows <- combined$header_rows[inc_hdrs]

  list(combined = combined, included_cols = inc_cols)
}

# 11. IMAGE EXTRACTION

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
#' Locates the first \code{\\{\\pict ... \\pngblip ... <hex>\\}} group, decodes the
#' hex-encoded bytes to raw, and returns the image data plus its declared
#' dimensions in twips (\code{\\picwgoal} / \code{\\pichgoal}).
#'
#' @param path Path to the .rtf file
#' @return list(png_bytes = raw, width_twips = integer, height_twips = integer)
#'   or NULL if no PNG found
extract_png <- function(path) {
  text <- rtf_read_raw(path)
  n    <- nchar(text)

  # Find the first \pict group that contains \pngblip. Earlier \pict groups
  # may be WMF/EMF renditions of the same image, so keep scanning past them.
  pict_text   <- NULL
  search_from <- 1L
  repeat {
    m <- regexpr("\\\\pict(?![a-zA-Z])", substr(text, search_from, n), perl = TRUE)
    if (m == -1L) return(NULL)
    pict_pos <- search_from + m[1L] - 1L

    # Walk back to opening brace
    brace_pos <- NA_integer_
    for (i in seq(pict_pos - 1L, max(1L, pict_pos - 5L), by = -1L)) {
      if (substr(text, i, i) == "{") { brace_pos <- i; break }
    }
    if (is.na(brace_pos)) { search_from <- pict_pos + 5L; next }

    end_pos <- find_matching_brace(text, brace_pos)
    if (is.na(end_pos)) return(NULL)

    candidate <- substr(text, brace_pos, end_pos)
    if (grepl("\\\\pngblip(?![a-zA-Z])", candidate, perl = TRUE)) {
      pict_text <- candidate
      break
    }
    search_from <- end_pos + 1L
  }

  # Drop {\*...} destination groups (e.g. {\*\blipuid <32 hex chars>}) so
  # their hex payload can't be mistaken for image data below
  pict_text <- gsub("\\{\\\\\\*[^{}]*\\}", "", pict_text, perl = TRUE)

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
  if (nchar(hex_clean) %% 2L != 0L) return(NULL)
  png_bytes <- hex_to_raw(hex_clean)

  list(
    png_bytes    = png_bytes,
    width_twips  = width_twips,
    height_twips = height_twips
  )
}
