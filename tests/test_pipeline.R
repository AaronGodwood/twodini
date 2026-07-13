# tests/test_pipeline.R
#
# Manual integration test for everything except RTF parsing.
# Builds synthetic RTF fixtures (table + image), runs them through
# html/xml generation and docx injection, and reports pass/fail.
#
# Usage (from R console, with package loaded):
#   source("tests/test_pipeline.R")
#   run_tests(docx_path = "path/to/your/bookmarked.docx",
#             table_bookmark = "MyTableBookmark",
#             image_bookmark = "MyImageBookmark")
#
# The docx must contain at least the two named bookmarks.
# Output is written to a temp file and opened if the shell supports it.

# ============================================================
# RTF FIXTURE BUILDERS
# ============================================================

# Build a minimal hand-crafted RTF with a two-page table.
# Page 1: one header row + two data rows (with a timeline label on row 1).
# Page 2: same header, two more data rows under a different timeline.
# Header section contains "Parameter: Cholesterol" on page 1,
# "Parameter: Triglycerides" on page 2.
make_table_rtf <- function(path) {
  # Helper: build an RTF table row
  # cells: list of list(text, width_twips, bold, trhdr)
  rtf_row <- function(cells, trhdr = FALSE) {
    trhdr_tag <- if (trhdr) "\\trhdr" else ""
    # Build cell definitions
    boundary <- 0L
    cell_defs <- paste(vapply(cells, function(c) {
      boundary <<- boundary + c$width_twips
      sprintf("\\cellx%d", boundary)
    }, character(1)), collapse = "")

    # Build cell contents
    cell_contents <- paste(vapply(cells, function(c) {
      bold_on  <- if (isTRUE(c$bold)) "\\b " else ""
      bold_off <- if (isTRUE(c$bold)) "\\b0 " else ""
      sprintf("{%s%s%s\\cell}", bold_on, c$text, bold_off)
    }, character(1)), collapse = "")

    sprintf("\\trowd%s%s%s\\row\n", trhdr_tag, cell_defs, cell_contents)
  }

  # Column widths in twips (label col 1440, three data cols 1080 each)
  w_label <- 1440L
  w_data  <- 1080L

  header_row <- rtf_row(
    list(
      list(text = "Treatment",  width_twips = w_label, bold = TRUE),
      list(text = "N",          width_twips = w_data,  bold = TRUE),
      list(text = "Mean",       width_twips = w_data,  bold = TRUE),
      list(text = "SD",         width_twips = w_data,  bold = TRUE)
    ),
    trhdr = TRUE
  )

  timeline1_row <- rtf_row(list(
    list(text = "Week 1", width_twips = w_label, bold = FALSE),
    list(text = "",       width_twips = w_data,  bold = FALSE),
    list(text = "",       width_twips = w_data,  bold = FALSE),
    list(text = "",       width_twips = w_data,  bold = FALSE)
  ))

  data_rows_p1 <- paste(
    rtf_row(list(
      list(text = "Placebo",  width_twips = w_label, bold = FALSE),
      list(text = "50",       width_twips = w_data,  bold = FALSE),
      list(text = "5.2",      width_twips = w_data,  bold = FALSE),
      list(text = "1.1",      width_twips = w_data,  bold = FALSE)
    )),
    rtf_row(list(
      list(text = "Treatment A", width_twips = w_label, bold = FALSE),
      list(text = "48",          width_twips = w_data,  bold = FALSE),
      list(text = "4.8",         width_twips = w_data,  bold = FALSE),
      list(text = "0.9",         width_twips = w_data,  bold = FALSE)
    ))
  )

  timeline2_row <- rtf_row(list(
    list(text = "Week 4", width_twips = w_label, bold = FALSE),
    list(text = "",       width_twips = w_data,  bold = FALSE),
    list(text = "",       width_twips = w_data,  bold = FALSE),
    list(text = "",       width_twips = w_data,  bold = FALSE)
  ))

  data_rows_p2 <- paste(
    rtf_row(list(
      list(text = "Placebo",  width_twips = w_label, bold = FALSE),
      list(text = "49",       width_twips = w_data,  bold = FALSE),
      list(text = "5.5",      width_twips = w_data,  bold = FALSE),
      list(text = "1.3",      width_twips = w_data,  bold = FALSE)
    )),
    rtf_row(list(
      list(text = "Treatment A", width_twips = w_label, bold = FALSE),
      list(text = "47",          width_twips = w_data,  bold = FALSE),
      list(text = "4.5",         width_twips = w_data,  bold = FALSE),
      list(text = "0.8",         width_twips = w_data,  bold = FALSE)
    ))
  )

  rtf_header <- function(param) {
    param_row <- rtf_row(list(
      list(text = sprintf("Parameter: %s", param), width_twips = w_label + 3L * w_data,
           bold = FALSE)
    ))
    sprintf("{\\header %s}", param_row)
  }

  page1 <- sprintf("\\sectd\n%s\n%s%s%s",
                   rtf_header("Cholesterol"),
                   header_row, timeline1_row, data_rows_p1)
  page2 <- sprintf("\\sectd\n%s\n%s%s%s",
                   rtf_header("Triglycerides"),
                   header_row, timeline2_row, data_rows_p2)

  rtf <- sprintf("{\\rtf1\\ansi\n%s\\sect\n%s\\sect\n}", page1, page2)
  writeLines(rtf, path, useBytes = FALSE)
  invisible(path)
}

# Build a minimal RTF containing a hex-encoded 1x1 white PNG embedded via \pngblip.
# The PNG is the smallest valid PNG file (67 bytes).
make_image_rtf <- function(path) {
  # Minimal 1x1 white PNG (67 bytes), hardcoded as hex
  png_hex <- paste0(
    "89504e470d0a1a0a",  # PNG signature
    "0000000d49484452", "000000010000000108020000009001", "2e00",  # IHDR chunk
    "0000000c49444154", "789c6260f8cf0000000200019e21bc33",        # IDAT chunk
    "00000000" , "49454e44ae426082"                                 # IEND chunk
  )
  # Realistic declared dimensions: 4 inches x 3 inches in twips
  # (1 inch = 1440 twips)
  w_twips <- 4L * 1440L  # 5760
  h_twips <- 3L * 1440L  # 4320

  pict <- sprintf(
    "{\\pict\\pngblip\\picwgoal%d\\pichgoal%d\n%s\n}",
    w_twips, h_twips, png_hex
  )
  rtf <- sprintf("{\\rtf1\\ansi\n\\sectd\n%s\\sect\n}", pict)
  writeLines(rtf, path, useBytes = FALSE)
  invisible(path)
}

# ============================================================
# ASSERTION HELPER
# ============================================================

pass_count <- 0L
fail_count <- 0L

check <- function(label, expr) {
  result <- tryCatch(
    isTRUE(expr),
    error = function(e) { message("  ERROR: ", conditionMessage(e)); FALSE }
  )
  if (result) {
    pass_count <<- pass_count + 1L
    message(sprintf("  [PASS] %s", label))
  } else {
    fail_count <<- fail_count + 1L
    message(sprintf("  [FAIL] %s", label))
  }
}

# ============================================================
# TEST SUITES
# ============================================================

test_html <- function(table_rtf_path) {
  message("\n--- HTML generation ---")

  html <- get_table_html(table_rtf_path)
  check("Returns a character string",         is.character(html) && length(html) == 1L)
  check("Contains <table>",                   grepl("<table", html, fixed = TRUE))
  check("Contains <thead>",                   grepl("<thead>", html, fixed = TRUE))
  check("Contains <tbody>",                   grepl("<tbody>", html, fixed = TRUE))
  check("Contains <colgroup>",                grepl("<colgroup>", html, fixed = TRUE))
  check("Contains column header 'Treatment'", grepl("Treatment", html, fixed = TRUE))
  check("Contains data value '5.2'",          grepl("5.2", html, fixed = TRUE))
  check("Font style declared on table",       grepl("Times New Roman", html, fixed = TRUE))
  check("Has bottom border on last header row",
        grepl("border-bottom", html, fixed = TRUE))

  # Filtered: only Week 1, exclude cols 3-4
  html_f <- get_table_html_output(table_rtf_path,
                                  excluded_cols = 3:4,
                                  timelines     = "Week 1")
  check("Filtered HTML excludes Week 4 data",   !grepl("Week 4", html_f, fixed = TRUE))
  check("Filtered HTML excludes col 3 value",   !grepl("5.2", html_f, fixed = TRUE))

  # Filtered by parameter
  html_p <- get_table_html_output(table_rtf_path, parameters = "Cholesterol")
  check("Parameter filter keeps Cholesterol page", grepl("Week 1", html_p, fixed = TRUE))
  check("Parameter filter drops Triglycerides page", !grepl("Week 4", html_p, fixed = TRUE))
}

test_xml <- function(table_rtf_path) {
  message("\n--- Word XML generation ---")

  xml_str <- get_table_xml(table_rtf_path)
  check("Returns a character string",      is.character(xml_str) && length(xml_str) == 1L)
  check("Starts with <w:tbl>",             grepl("^<w:tbl", xml_str))
  check("Contains <w:tblGrid>",            grepl("<w:tblGrid>", xml_str, fixed = TRUE))
  check("Contains <w:tblHeader/>",         grepl("<w:tblHeader/>", xml_str, fixed = TRUE))
  check("Contains Times New Roman",        grepl("Times New Roman", xml_str, fixed = TRUE))
  check("Contains font size 20 half-pts",  grepl('w:val="20"', xml_str, fixed = TRUE))
  check("Contains bold tag",               grepl("<w:b/>", xml_str, fixed = TRUE))
  check("Contains left alignment",         grepl('w:val="left"',   xml_str, fixed = TRUE))
  check("Contains center alignment",       grepl('w:val="center"', xml_str, fixed = TRUE))
  check("Has bottom border on last row",
        grepl('w:val="single"', xml_str, fixed = TRUE))

  # Parses as valid XML
  parsed <- tryCatch(xml2::read_xml(xml_str), error = function(e) NULL)
  check("Valid XML", !is.null(parsed))
}

test_table_info <- function(table_rtf_path) {
  message("\n--- get_table_info ---")

  info <- get_table_info(table_rtf_path)
  check("n_cols is 4",                    info$n_cols == 4L)
  check("n_rows is 6 (2 pages x (1 timeline label + 2 data rows))",
        info$n_rows == 6L)
  check("col_names has 4 entries",        length(info$col_names) == 4L)
  check("col_names[1] is 'Treatment'",    info$col_names[1] == "Treatment")
  check("parameters detected",            length(info$parameters) > 0L)
  check("'Cholesterol' in parameters",    "Cholesterol" %in% info$parameters)
  check("'Triglycerides' in parameters",  "Triglycerides" %in% info$parameters)
  check("timelines detected",             length(info$timelines) > 0L)
  check("'Week 1' in timelines",          "Week 1" %in% info$timelines)
  check("'Week 4' in timelines",          "Week 4" %in% info$timelines)
}

test_image_extraction <- function(image_rtf_path) {
  message("\n--- Image extraction ---")

  check("is_image_rtf() TRUE for image RTF",  is_image_rtf(image_rtf_path))
  check("is_image_rtf() FALSE for table RTF",
        !is_image_rtf(file.path(dirname(image_rtf_path), "test_table.rtf")))

  img <- extract_png(image_rtf_path)
  check("extract_png() returns a list",          is.list(img))
  check("png_bytes is raw",                      is.raw(img$png_bytes))
  check("png_bytes is non-empty",                length(img$png_bytes) > 0L)
  check("width_twips is correct (5760)",         img$width_twips == 5760L)
  check("height_twips is correct (4320)",        img$height_twips == 4320L)
  check("PNG signature correct (bytes 1-4)",
        identical(img$png_bytes[1:4], as.raw(c(0x89, 0x50, 0x4e, 0x47))))
}

test_bookmark_extraction <- function(docx_path) {
  message("\n--- Bookmark extraction ---")

  bm <- extract_bookmarks(docx_path)
  check("Returns named character vector", is.character(bm) && !is.null(names(bm)))
  check("At least one bookmark found",    length(bm) > 0L)
  check("No internal bookmarks (no '_' prefix)",
        !any(startsWith(names(bm), "_")))
  message(sprintf("  INFO: Found bookmarks: %s", paste(names(bm), collapse = ", ")))
}

test_docx_session <- function(docx_path, table_rtf_path, image_rtf_path,
                              table_bookmark, image_bookmark) {
  message("\n--- Docx session (open/inject/close) ---")

  out_path <- tempfile(fileext = ".docx")
  on.exit(unlink(out_path), add = TRUE)

  session <- open_docx(docx_path)
  check("Session has jump_table",       is.list(session$jump_table))
  check("Table bookmark in jump table", !is.null(session$jump_table[[table_bookmark]]))
  check("Image bookmark in jump table", !is.null(session$jump_table[[image_bookmark]]))
  check("text_width_emu is positive",   session$text_width_emu > 0L)
  check("next_img_id starts at 1",      session$next_img_id >= 1L)

  # Inject table
  xml_str  <- get_table_xml(table_rtf_path)
  t_result <- inject_table(session, table_bookmark, xml_str)
  check("inject_table() returns TRUE", isTRUE(t_result))

  # Inject image
  img      <- extract_png(image_rtf_path)
  i_result <- inject_image(session, image_bookmark,
                            img$png_bytes, img$width_twips, img$height_twips)
  check("inject_image() returns TRUE", isTRUE(i_result))

  # Close and verify output file exists and is non-empty
  close_docx(session, out_path)
  check("Output .docx exists",            file.exists(out_path))
  check("Output .docx is non-empty",      file.info(out_path)$size > 0L)

  # Unzip and verify structure
  verify_dir <- tempfile()
  dir.create(verify_dir)
  on.exit(unlink(verify_dir, recursive = TRUE), add = TRUE)
  unzip(out_path, exdir = verify_dir)

  doc_xml <- file.path(verify_dir, "word", "document.xml")
  check("document.xml present in output", file.exists(doc_xml))

  content <- paste(readLines(doc_xml, warn = FALSE), collapse = "")
  check("Output contains <w:tbl>",     grepl("<w:tbl>", content, fixed = TRUE))
  check("Output contains <w:drawing>", grepl("w:drawing", content, fixed = TRUE))

  media_files <- list.files(file.path(verify_dir, "word", "media"))
  check("word/media/ contains PNG file",
        any(grepl("\\.png$", media_files, ignore.case = TRUE)))

  rels_xml <- file.path(verify_dir, "word", "_rels", "document.xml.rels")
  rels_content <- paste(readLines(rels_xml, warn = FALSE), collapse = "")
  check("Rels file references image relationship",
        grepl("relationships/image", rels_content, fixed = TRUE))
}

test_process_document <- function(docx_path, table_rtf_path, image_rtf_path,
                                  table_bookmark, image_bookmark) {
  message("\n--- process_document() end-to-end ---")

  out_path <- tempfile(fileext = ".docx")
  on.exit(unlink(out_path), add = TRUE)

  config <- data.frame(
    Bookmark = c(table_bookmark, image_bookmark),
    Table    = c("test_table",   "test_image"),
    stringsAsFactors = FALSE
  )
  paths <- list(
    test_table = table_rtf_path,
    test_image = image_rtf_path
  )
  # No filters — include everything (keyed by row index as character)
  selections <- list(
    "1" = list(excluded_cols = NULL, excluded_rows = NULL,
               excluded_header_rows = NULL,
               parameters = NULL, timelines = NULL),
    "2" = list()
  )

  err <- tryCatch({
    process_document(docx_path, config, paths, selections, out_path)
    NULL
  }, error = function(e) conditionMessage(e))

  check("process_document() completes without error", is.null(err))
  check("Output file exists",      file.exists(out_path))
  check("Output file non-empty",   file.info(out_path)$size > 0L)
}

# ============================================================
# MAIN ENTRY POINT
# ============================================================

#' Run all pipeline tests
#'
#' @param docx_path Path to a real .docx file containing at least two bookmarks
#' @param table_bookmark Name of the bookmark where the table should be injected
#' @param image_bookmark Name of the bookmark where the image should be injected
run_tests <- function(docx_path, table_bookmark, image_bookmark) {
  if (!file.exists(docx_path))
    stop("docx_path does not exist: ", docx_path)

  tmp_dir <- tempfile()
  dir.create(tmp_dir)
  on.exit(unlink(tmp_dir, recursive = TRUE), add = TRUE)

  table_rtf_path <- file.path(tmp_dir, "test_table.rtf")
  image_rtf_path <- file.path(tmp_dir, "test_image.rtf")

  message("Building RTF fixtures...")
  make_table_rtf(table_rtf_path)
  make_image_rtf(image_rtf_path)
  message("  test_table.rtf: ", table_rtf_path)
  message("  test_image.rtf: ", image_rtf_path)

  # Reset counters
  pass_count <<- 0L
  fail_count <<- 0L

  test_table_info(table_rtf_path)
  test_html(table_rtf_path)
  test_xml(table_rtf_path)
  test_image_extraction(image_rtf_path)
  test_bookmark_extraction(docx_path)
  test_docx_session(docx_path, table_rtf_path, image_rtf_path,
                    table_bookmark, image_bookmark)
  test_process_document(docx_path, table_rtf_path, image_rtf_path,
                        table_bookmark, image_bookmark)

  message(sprintf(
    "\n============================\nResults: %d passed, %d failed\n============================",
    pass_count, fail_count
  ))

  invisible(list(passed = pass_count, failed = fail_count))
}
