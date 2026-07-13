# tests/test_pdf_parsing.R
#
# Smoke test for the Phase-1 PDF parser (R/pdf_parse.R).
#
# Generates synthetic fixtures via tests/make_test_pdfs.R, runs
# pdf_page_streams() on each, and verifies:
#   - page count matches expectation
#   - every page has at least one /F1 font with a non-empty CMap
#   - decoding the single <hex> Tj operand through the CMap reproduces
#     the expected text
#
# Usage:
#   source("tests/test_pdf_parsing.R")
#   run_pdf_tests()                              # default: tempdir()
#   run_pdf_tests("C:/path/to/write/pdfs")       # persist fixtures

# ---- SETUP ----

# Locate the package source files. This script can be source()d from either
# the package root (Twodini/) or from tests/, so try both.
.src_one <- function(rel) {
  for (base in c(".", "..", "Twodini", "Twodini/Twodini")) {
    p <- file.path(base, rel)
    if (file.exists(p)) { source(p, chdir = TRUE); return(invisible(TRUE)) }
  }
  stop("Could not locate ", rel)
}

.src_one("R/pdf_parse.R")
.src_one("tests/make_test_pdfs.R")

# ---- TEST RUNNER ----

pass_count <- 0L
fail_count <- 0L

check <- function(label, expr) {
  result <- tryCatch(isTRUE(expr), error = function(e) {
    message(sprintf("  [ERROR] %s: %s", label, conditionMessage(e)))
    FALSE
  })
  if (result) {
    pass_count <<- pass_count + 1L
  } else {
    fail_count <<- fail_count + 1L
    message(sprintf("  [FAIL] %s", label))
  }
}

# ---- MINIMAL CONTENT-STREAM DECODER ----
#
# Enough to extract the text shown by a single "BT ... <HEX> Tj ... ET" block.
# This is a test helper, not part of the parser — Phase 3 will replace it
# with the real glyph-placement pipeline.
#
# Assumes the fixture content streams use a single hex-string Tj operand,
# which is how the generators are written. Anything fancier falls through
# and the assertion fails, which is the behaviour we want.
decode_simple_text <- function(content_bytes, cmap) {
  txt <- rawToChar(content_bytes)
  # Find the hex operand of Tj, e.g. "<4142> Tj"
  m <- regexpr("<([0-9A-Fa-f]+)>\\s*Tj", txt, perl = TRUE)
  if (m == -1L) return(NA_character_)
  hex <- regmatches(txt, regexpr("<[0-9A-Fa-f]+>", txt, perl = TRUE))
  hex <- sub("^<(.*)>$", "\\1", hex)

  # Split into 2-char code units and map through the CMap
  codes <- substring(hex, seq(1L, nchar(hex) - 1L, by = 2L),
                          seq(2L, nchar(hex),      by = 2L))
  codes <- tolower(codes)

  cps <- cmap[codes]
  if (anyNA(cps)) {
    return(paste0("<unmapped:", paste(codes[is.na(cps)], collapse = ","), ">"))
  }
  intToUtf8(as.integer(cps), multiple = FALSE)
}

# ---- MAIN ----

#' Generate synthetic PDF fixtures and verify pdf_page_streams parses them.
#'
#' @param pdf_dir Directory to write fixtures into. Defaults to a tempdir.
#' @return invisibly, list(pass, fail)
run_pdf_tests <- function(pdf_dir = tempfile("pdf_fixtures_")) {
  if (!dir.exists(pdf_dir)) dir.create(pdf_dir, recursive = TRUE)

  pass_count <<- 0L
  fail_count <<- 0L

  message(sprintf("Generating fixtures in %s", pdf_dir))
  expectations <- make_test_pdfs(pdf_dir)
  message(sprintf("Testing %d PDF fixtures\n", length(expectations)))

  for (name in names(expectations)) {
    message(sprintf("-- %s", name))
    path <- file.path(pdf_dir, name)
    exp <- expectations[[name]]

    pages <- tryCatch(pdf_page_streams(path),
                      error = function(e) {
                        message(sprintf("  [ERROR] pdf_page_streams: %s",
                                        conditionMessage(e)))
                        NULL
                      })
    if (is.null(pages)) {
      fail_count <<- fail_count + 1L
      next
    }

    check(sprintf("%s: page count", name),
          length(pages) == exp$n_pages)

    for (pi in seq_along(pages)) {
      p <- pages[[pi]]
      check(sprintf("%s page %d: has content bytes", name, pi),
            is.raw(p$content) && length(p$content) > 0L)
      check(sprintf("%s page %d: has /F1 font", name, pi),
            "/F1" %in% names(p$fonts))
      f1 <- p$fonts[["/F1"]]
      check(sprintf("%s page %d: CMap non-empty", name, pi),
            !is.null(f1) && length(f1$cmap) > 0L)

      if (!is.null(f1) && length(f1$cmap) > 0L &&
          pi <= length(exp$texts)) {
        got <- decode_simple_text(p$content, f1$cmap)
        check(sprintf("%s page %d: decoded text '%s'", name, pi, exp$texts[[pi]]),
              identical(got, exp$texts[[pi]]))
      }
    }
  }

  message(sprintf("\n%d passed, %d failed", pass_count, fail_count))
  invisible(list(pass = pass_count, fail = fail_count))
}

# Allow `Rscript tests/test_pdf_parsing.R` to auto-run
if (!interactive() && identical(sys.nframe(), 0L)) {
  res <- run_pdf_tests()
  if (res$fail > 0L) quit(status = 1L)
}
