# tests/test_rtf_parsing.R
#
# Regression test: parse every RTF in the test_data folder and check
# that cell values contain only expected characters (no garbled unicode,
# no box chars, no stray control words).
#
# Usage:
#   source("tests/test_rtf_parsing.R")
#   run_rtf_tests()                          # uses ../test_data/rtf
#   run_rtf_tests("path/to/rtf/folder")      # custom folder
#
# On first run with --save flag, writes a snapshot CSV of all parsed cell
# values. On subsequent runs, compares against the snapshot and reports
# any differences.

pass_count <- 0L
fail_count <- 0L

check <- function(label, expr) {
  result <- tryCatch(isTRUE(expr), error = function(e) FALSE)
  if (result) {
    pass_count <<- pass_count + 1L
  } else {
    fail_count <<- fail_count + 1L
    message(sprintf("  [FAIL] %s", label))
  }
}

# Characters that should never appear in parsed cell text
has_artifacts <- function(text) {
  # Box replacement char (U+FFFD or U+25A1), stray backslashes from
  # unstripped control words, or raw RTF fragments
  grepl("[\ufffd\u25a1]", text, perl = TRUE) ||
    grepl("\\\\[a-zA-Z]{2,}", text, perl = TRUE)
}

#' Parse all RTFs in a folder and run checks
#'
#' @param rtf_folder Path to folder containing .rtf files
#' @param save_snapshot If TRUE, write/overwrite the snapshot CSV
#' @param snapshot_path Path to the snapshot CSV file
run_rtf_tests <- function(rtf_folder = "../test_data/rtf",
                          save_snapshot = FALSE,
                          snapshot_path = "tests/rtf_snapshot.csv") {
  files <- list.files(rtf_folder, pattern = "\\.rtf$",
                      full.names = TRUE, ignore.case = TRUE)
  if (length(files) == 0L) {
    stop("No RTF files found in: ", rtf_folder)
  }

  pass_count <<- 0L
  fail_count <<- 0L

  message(sprintf("Testing %d RTF files from %s\n", length(files), rtf_folder))

  all_rows <- list()

  for (f in files) {
    fname <- basename(f)

    # --- Image RTFs: just check detection + extraction ---
    if (is_image_rtf(f)) {
      check(
        sprintf("%s: is_image_rtf returns TRUE", fname),
        TRUE
      )
      img <- tryCatch(extract_png(f), error = function(e) NULL)
      check(
        sprintf("%s: extract_png returns non-NULL", fname),
        !is.null(img)
      )
      if (!is.null(img)) {
        check(
          sprintf("%s: PNG signature valid", fname),
          length(img$png_bytes) >= 4L &&
            identical(img$png_bytes[1:4],
                      as.raw(c(0x89, 0x50, 0x4e, 0x47)))
        )
        check(
          sprintf("%s: dimensions present", fname),
          !is.na(img$width_twips) && !is.na(img$height_twips)
        )
      }
      next
    }

    # --- Table RTFs: parse and validate cell content ---
    pages <- tryCatch(parse_rtf(f), error = function(e) {
      message(sprintf("  [FAIL] %s: parse error: %s", fname, conditionMessage(e)))
      fail_count <<- fail_count + 1L
      NULL
    })
    if (is.null(pages)) next

    check(sprintf("%s: at least 1 page", fname), length(pages) >= 1L)

    combined <- combine_pages(pages)

    check(
      sprintf("%s: has header rows", fname),
      length(combined$header_rows) >= 1L
    )
    check(
      sprintf("%s: has data rows", fname),
      length(combined$data_rows) >= 1L
    )
    check(
      sprintf("%s: has column widths", fname),
      length(combined$col_widths_twips) >= 1L
    )

    # Check every cell for artifacts
    artifact_found <- FALSE
    check_cells <- function(rows, row_type) {
      for (ri in seq_along(rows)) {
        for (ci in seq_along(rows[[ri]]$cells)) {
          cell_text <- rows[[ri]]$cells[[ci]]$text
          if (has_artifacts(cell_text)) {
            if (!artifact_found) {
              message(sprintf("  [FAIL] %s: artifact in %s row %d col %d: '%s'",
                              fname, row_type, ri, ci, cell_text))
              artifact_found <<- TRUE
            }
          }
        }
      }
    }
    check_cells(combined$header_rows, "header")
    check_cells(combined$data_rows, "data")
    check(sprintf("%s: no artifacts in cell text", fname), !artifact_found)

    # Collect snapshot data
    collect_row <- function(rows, row_type) {
      for (ri in seq_along(rows)) {
        vals <- vapply(rows[[ri]]$cells, function(c) c$text, character(1))
        all_rows[[length(all_rows) + 1L]] <<- data.frame(
          file     = fname,
          row_type = row_type,
          row_idx  = ri,
          cells    = paste(vals, collapse = " | "),
          stringsAsFactors = FALSE
        )
      }
    }
    collect_row(combined$header_rows, "header")
    collect_row(combined$data_rows, "data")

    # Validate HTML generation
    html <- tryCatch(get_table_html(f), error = function(e) NULL)
    check(sprintf("%s: HTML generates", fname), !is.null(html))
    if (!is.null(html)) {
      check(sprintf("%s: HTML has <table>", fname), grepl("<table", html, fixed = TRUE))
    }

    # Validate XML generation
    xml_str <- tryCatch(get_table_xml(f), error = function(e) NULL)
    check(sprintf("%s: XML generates", fname), !is.null(xml_str))
    if (!is.null(xml_str)) {
      check(sprintf("%s: XML has <w:tbl>", fname), grepl("<w:tbl", xml_str, fixed = TRUE))
    }

    # Validate filtering doesn't crash
    params <- get_parameters(pages)
    tlines <- get_timelines(pages)
    if (length(params) > 0L) {
      filtered <- tryCatch(
        get_table_html_output(f, parameters = params[1]),
        error = function(e) NULL
      )
      check(sprintf("%s: parameter filter works", fname), !is.null(filtered))
    }
    if (length(tlines) > 0L) {
      filtered <- tryCatch(
        get_table_html_output(f, timelines = tlines[1]),
        error = function(e) NULL
      )
      check(sprintf("%s: timeline filter works", fname), !is.null(filtered))
    }
  }

  # Snapshot handling
  snapshot_df <- do.call(rbind, all_rows)

  if (save_snapshot) {
    write.csv(snapshot_df, snapshot_path, row.names = FALSE)
    message(sprintf("\nSnapshot saved: %s (%d rows)", snapshot_path, nrow(snapshot_df)))
  } else if (file.exists(snapshot_path)) {
    prev <- read.csv(snapshot_path, stringsAsFactors = FALSE)
    if (identical(dim(prev), dim(snapshot_df)) &&
        all(prev$file == snapshot_df$file) &&
        all(prev$cells == snapshot_df$cells)) {
      message("\nSnapshot comparison: MATCH")
      pass_count <<- pass_count + 1L
    } else {
      message("\nSnapshot comparison: MISMATCH")
      fail_count <<- fail_count + 1L
      # Show first few differences
      if (nrow(prev) != nrow(snapshot_df)) {
        message(sprintf("  Row count changed: %d -> %d", nrow(prev), nrow(snapshot_df)))
      } else {
        diffs <- which(prev$cells != snapshot_df$cells)
        for (di in head(diffs, 5)) {
          message(sprintf("  %s %s row %d:",
                          snapshot_df$file[di], snapshot_df$row_type[di],
                          snapshot_df$row_idx[di]))
          message(sprintf("    was:  %s", prev$cells[di]))
          message(sprintf("    now:  %s", snapshot_df$cells[di]))
        }
        if (length(diffs) > 5L)
          message(sprintf("  ... and %d more differences", length(diffs) - 5L))
      }
    }
  } else {
    message(sprintf(
      "\nNo snapshot found at %s. Run with save_snapshot=TRUE to create one.",
      snapshot_path
    ))
  }

  message(sprintf(
    "\n============================\nResults: %d passed, %d failed\n============================",
    pass_count, fail_count
  ))

  invisible(list(passed = pass_count, failed = fail_count, snapshot = snapshot_df))
}
