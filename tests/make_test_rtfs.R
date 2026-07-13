# tests/make_test_rtfs.R
#
# Generates synthetic RTF files matching the expected SAS output format.
# Run this script to produce test fixtures in a directory of your choice.
#
# Usage:
#   source("tests/make_test_rtfs.R")
#   make_test_rtfs("C:/path/to/output/folder")

# ============================================================
# LOW-LEVEL RTF BUILDING HELPERS
# ============================================================

# Convenience constructor so all cells have consistent named fields
cell <- function(text, width_twips, bold = FALSE,
                 merge_first = FALSE, merge_cont = FALSE) {
  list(text        = as.character(text),
       width_twips = as.integer(width_twips),  # strip any names
       bold        = bold,
       merge_first = merge_first,
       merge_cont  = merge_cont)
}

# One RTF table row.
# cells: list of cell() objects
# trhdr: TRUE if this is a header row (\trhdr)
rtf_row <- function(cells, trhdr = FALSE) {
  trhdr_tag <- if (trhdr) "\\trhdr" else ""

  # Cell definitions: merge flags then \cellx boundary
  boundary  <- 0L
  cell_defs <- paste(vapply(cells, function(c) {
    boundary  <<- boundary + c$width_twips
    merge_tag <- if (isTRUE(c$merge_first)) "\\clmgf"
            else if (isTRUE(c$merge_cont))  "\\clmrg"
            else                             ""
    sprintf("%s\\cellx%d", merge_tag, boundary)
  }, character(1)), collapse = "")

  # Cell contents
  cell_contents <- paste(vapply(cells, function(c) {
    bold_on  <- if (isTRUE(c$bold)) "\\b "  else ""
    bold_off <- if (isTRUE(c$bold)) "\\b0 " else ""
    sprintf("{%s%s%s\\cell}", bold_on, c$text, bold_off)
  }, character(1)), collapse = "")

  sprintf("\\trowd%s%s%s\\row\n", trhdr_tag, cell_defs, cell_contents)
}

# RTF header section containing a parameter line in its bottom row
rtf_header_section <- function(parameter = NULL, page_width_twips = 12240L,
                               margin_twips = 1800L) {
  text_width <- page_width_twips - 2L * margin_twips  # 8640 twips

  # Top row: study/table title (purely decorative, tests robustness)
  title_row <- rtf_row(list(
    list(text = "Study XYZ - Summary Table", width_twips = text_width, bold = TRUE)
  ), trhdr = FALSE)

  # Bottom row: parameter line (or blank if no parameter)
  param_text <- if (!is.null(parameter)) sprintf("Parameter: %s", parameter) else ""
  param_row  <- rtf_row(list(
    list(text = param_text, width_twips = text_width, bold = FALSE)
  ), trhdr = FALSE)

  sprintf("{\\header %s%s}", title_row, param_row)
}

# RTF footer section
rtf_footer_section <- function(page_num_text = "Page 1") {
  sprintf("{\\footer {\\pard %s\\par}}", page_num_text)
}

# Section definition block (page setup for A4 portrait with standard margins)
# pgwsxn=12240 (8.5in), pghsxn=15840 (11in), margl/r=1800 (1.25in), margt/b=1440 (1in)
rtf_sectd <- function() {
  "\\sectd\\pgwsxn12240\\pghsxn15840\\marglsxn1800\\margrsxn1800\\margtsxn1440\\margbsxn1440"
}

# Wrap pages into a complete RTF document
rtf_document <- function(pages) {
  page_blocks <- paste(vapply(seq_along(pages), function(i) {
    sprintf("%s\n%s", pages[[i]], if (i < length(pages)) "\\sect" else "")
  }, character(1)), collapse = "\n")
  sprintf("{\\rtf1\\ansi\\deff0\n{\\fonttbl{\\f0 Times New Roman;}}\n%s\n}", page_blocks)
}

# ============================================================
# FIXTURE 1: Single-parameter, single-timeline table
# Simple 4-column table, 1 page, no filtering needed.
# ============================================================

make_rtf_simple <- function(path) {
  wl <- 2160L; wd <- 1440L  # label / data column widths in twips

  hdr <- rtf_row(list(
    cell("Treatment Group", wl, bold = TRUE),
    cell("N",               wd, bold = TRUE),
    cell("Mean",            wd, bold = TRUE),
    cell("SD",              wd, bold = TRUE)
  ), trhdr = TRUE)

  rows <- paste(
    rtf_row(list(cell("Placebo",     wl), cell("52",    wd), cell("185.3", wd), cell("24.1", wd))),
    rtf_row(list(cell("Treatment A", wl), cell("49",    wd), cell("178.6", wd), cell("22.7", wd))),
    rtf_row(list(cell("Treatment B", wl), cell("51",    wd), cell("172.4", wd), cell("21.3", wd)))
  )

  page <- paste(
    rtf_sectd(),
    rtf_header_section(parameter = "Cholesterol"),
    rtf_footer_section("Page 1"),
    hdr, rows,
    sep = "\n"
  )

  writeLines(rtf_document(list(page)), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# FIXTURE 2: Multi-parameter table (3 pages, one per parameter)
# Tests parameter filtering.
# ============================================================

make_rtf_multi_param <- function(path) {
  wl <- 2160L; wd <- 1440L

  hdr <- rtf_row(list(
    cell("Treatment Group", wl, bold = TRUE),
    cell("N",               wd, bold = TRUE),
    cell("Mean",            wd, bold = TRUE),
    cell("SD",              wd, bold = TRUE)
  ), trhdr = TRUE)

  make_page <- function(param, means, page_n) {
    rows <- paste(
      rtf_row(list(cell("Placebo",     wl), cell("52", wd), cell(means[1], wd), cell("24.1", wd))),
      rtf_row(list(cell("Treatment A", wl), cell("49", wd), cell(means[2], wd), cell("22.7", wd)))
    )
    paste(
      rtf_sectd(),
      rtf_header_section(parameter = param),
      rtf_footer_section(sprintf("Page %d", page_n)),
      hdr, rows,
      sep = "\n"
    )
  }

  pages <- list(
    make_page("Cholesterol",   c("185.3", "178.6"), 1L),
    make_page("Triglycerides", c("142.1", "138.4"), 2L),
    make_page("HDL",           c("52.3",  "54.1"),  3L)
  )

  writeLines(rtf_document(pages), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# FIXTURE 3: Multi-timeline table (timelines interleaved across 2 pages)
# Tests timeline filtering.
# ============================================================

make_rtf_multi_timeline <- function(path) {
  wl <- 2160L; wd <- 1440L

  hdr <- rtf_row(list(
    cell("Treatment Group", wl, bold = TRUE),
    cell("N",               wd, bold = TRUE),
    cell("Mean",            wd, bold = TRUE),
    cell("SD",              wd, bold = TRUE)
  ), trhdr = TRUE)

  # Timeline label row: only col 1 filled, rest empty
  timeline_row <- function(label) {
    rtf_row(list(cell(label, wl), cell("", wd), cell("", wd), cell("", wd)))
  }

  data_rows <- function(placebo_mean, trt_mean) {
    paste(
      rtf_row(list(cell("Placebo",     wl), cell("52", wd), cell(placebo_mean, wd), cell("24.1", wd))),
      rtf_row(list(cell("Treatment A", wl), cell("49", wd), cell(trt_mean,     wd), cell("22.7", wd)))
    )
  }

  page1 <- paste(
    rtf_sectd(),
    rtf_header_section(parameter = "Cholesterol"),
    rtf_footer_section("Page 1"),
    hdr,
    timeline_row("Week 1"),  data_rows("185.3", "178.6"),
    timeline_row("Week 4"),  data_rows("183.1", "175.2"),
    sep = "\n"
  )

  page2 <- paste(
    rtf_sectd(),
    rtf_header_section(parameter = "Cholesterol"),
    rtf_footer_section("Page 2"),
    hdr,
    timeline_row("Week 8"),  data_rows("180.5", "171.8"),
    timeline_row("Week 12"), data_rows("177.2", "168.4"),
    sep = "\n"
  )

  writeLines(rtf_document(list(page1, page2)), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# FIXTURE 4: Multi-parameter AND multi-timeline
# Tests both filters together.
# ============================================================

make_rtf_combined_filters <- function(path) {
  wl <- 2160L; wd <- 1440L

  hdr <- rtf_row(list(
    cell("Treatment Group", wl, bold = TRUE),
    cell("N",               wd, bold = TRUE),
    cell("Mean",            wd, bold = TRUE),
    cell("SD",              wd, bold = TRUE)
  ), trhdr = TRUE)

  timeline_row <- function(label) {
    rtf_row(list(cell(label, wl), cell("", wd), cell("", wd), cell("", wd)))
  }

  data_row <- function(trt, n, mean, sd) {
    rtf_row(list(cell(trt, wl), cell(n, wd), cell(mean, wd), cell(sd, wd)))
  }

  make_param_page <- function(param, page_n) {
    paste(
      rtf_sectd(),
      rtf_header_section(parameter = param),
      rtf_footer_section(sprintf("Page %d", page_n)),
      hdr,
      timeline_row("Week 1"),
      data_row("Placebo",     "52", "185.3", "24.1"),
      data_row("Treatment A", "49", "178.6", "22.7"),
      timeline_row("Week 4"),
      data_row("Placebo",     "51", "183.1", "23.8"),
      data_row("Treatment A", "48", "175.2", "21.9"),
      sep = "\n"
    )
  }

  pages <- list(
    make_param_page("Cholesterol",   1L),
    make_param_page("Triglycerides", 2L)
  )

  writeLines(rtf_document(pages), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# FIXTURE 5: Merged header cells (colspan spanning)
# Tests \clmgf / \clmrg merge detection.
# ============================================================

make_rtf_merged_headers <- function(path) {
  # Layout: Label | <-- Dose A (spans 2) --> | <-- Dose B (spans 2) -->
  #         Col1  |   N   |  Mean  |   N   |  Mean
  w_label <- 2160L
  w_sub   <- 1080L  # each sub-column

  # Top header row: merged cells spanning pairs
  top_hdr <- rtf_row(list(
    cell("",       w_label, bold = TRUE),
    cell("Dose A", w_sub,   bold = TRUE, merge_first = TRUE),
    cell("Dose A", w_sub,   bold = TRUE, merge_cont  = TRUE),
    cell("Dose B", w_sub,   bold = TRUE, merge_first = TRUE),
    cell("Dose B", w_sub,   bold = TRUE, merge_cont  = TRUE)
  ), trhdr = TRUE)

  # Second header row: sub-column labels
  sub_hdr <- rtf_row(list(
    cell("Treatment Group", w_label, bold = TRUE),
    cell("N",               w_sub,   bold = TRUE),
    cell("Mean",            w_sub,   bold = TRUE),
    cell("N",               w_sub,   bold = TRUE),
    cell("Mean",            w_sub,   bold = TRUE)
  ), trhdr = TRUE)

  rows <- paste(
    rtf_row(list(cell("Placebo",     w_label), cell("52", w_sub), cell("185.3", w_sub),
                 cell("48",          w_sub),   cell("172.1", w_sub))),
    rtf_row(list(cell("Treatment A", w_label), cell("49", w_sub), cell("178.6", w_sub),
                 cell("45",          w_sub),   cell("165.3", w_sub)))
  )

  page <- paste(
    rtf_sectd(),
    rtf_header_section(parameter = NULL),  # no parameter — page always kept
    rtf_footer_section("Page 1"),
    top_hdr, sub_hdr, rows,
    sep = "\n"
  )

  writeLines(rtf_document(list(page)), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# FIXTURE 6: Image RTF (PNG embedded via \pngblip)
# Minimal valid PNG (1x1 white pixel), declared at 4x3 inches.
# ============================================================

make_rtf_image <- function(path) {
  # Smallest valid PNG: 1x1 white pixel, 67 bytes
  png_hex <- paste0(
    "89504e470d0a1a0a",                          # signature
    "0000000d49484452",                          # IHDR length + type
    "00000001",                                  # width = 1
    "00000001",                                  # height = 1
    "08020000009001",                            # bit depth, colour, compression, filter, interlace + CRC (partial)
    "2e00",                                      # CRC remainder
    "0000000c49444154",                          # IDAT length + type
    "789c6260f8cf00000002",                      # zlib-compressed pixel data
    "00019e21bc33",                              # IDAT CRC
    "00000000",                                  # IEND length
    "49454e44ae426082"                           # IEND type + CRC
  )

  w_twips <- 5760L  # 4 inches
  h_twips <- 4320L  # 3 inches

  pict <- sprintf(
    "{\\pict\\pngblip\\picwgoal%d\\pichgoal%d\n%s\n}",
    w_twips, h_twips, png_hex
  )

  page <- paste(
    rtf_sectd(),
    rtf_header_section(parameter = NULL),
    rtf_footer_section("Page 1"),
    pict,
    sep = "\n"
  )

  writeLines(rtf_document(list(page)), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# FIXTURE 7: Escape sequences (\'xx hex, \uN unicode, \ucN fallbacks)
# Locks in CP1252 + Unicode decoding; all escapes are plain ASCII on disk.
# ============================================================

make_rtf_escapes <- function(path) {
  wl <- 2160L; wd <- 1440L

  hdr <- rtf_row(list(
    cell("Statistic",              wl, bold = TRUE),
    cell("Mean \\'b1 SD",          wd, bold = TRUE),   # plus-minus
    cell("\\'93Range\\'94",        wd, bold = TRUE)    # CP1252 smart quotes
  ), trhdr = TRUE)

  rows <- paste(
    rtf_row(list(cell("\\uc1\\u8804\\'3f 5 years", wl),   # <= with hex fallback
                 cell("185.3 \\'b1 24.1", wd),
                 cell("120\\u8211?190", wd))),             # en dash, ? fallback
    rtf_row(list(cell("\\uc0\\u916 from baseline", wl),   # Delta, no fallback
                 cell("-6.7 \\'b1 3.2", wd),
                 cell("-12\\u8211?-1", wd)))
  )

  page <- paste(
    rtf_sectd(),
    rtf_header_section(parameter = NULL),
    rtf_footer_section("Page 1"),
    hdr, rows,
    sep = "\n"
  )

  writeLines(rtf_document(list(page)), path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# MAIN: generate all fixtures
# ============================================================

#' Generate all test RTF fixtures into a directory
#'
#' @param output_dir Directory to write the RTF files into (created if needed)
make_test_rtfs <- function(output_dir = "tests/testthat/fixtures") {
  dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

  files <- list(
    "01_simple.rtf"            = make_rtf_simple,
    "02_multi_param.rtf"       = make_rtf_multi_param,
    "03_multi_timeline.rtf"    = make_rtf_multi_timeline,
    "04_combined_filters.rtf"  = make_rtf_combined_filters,
    "05_merged_headers.rtf"    = make_rtf_merged_headers,
    "06_image.rtf"             = make_rtf_image,
    "07_escapes.rtf"           = make_rtf_escapes
  )

  for (fname in names(files)) {
    fpath <- file.path(output_dir, fname)
    files[[fname]](fpath)
    message(sprintf("  Written: %s", fpath))
  }

  message(sprintf("\n%d RTF fixtures written to: %s", length(files), output_dir))
  invisible(output_dir)
}
