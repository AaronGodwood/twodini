# OOXML and HTML rendering of parsed tables.

simple_fixture <- function() test_path("fixtures", "01_simple.rtf")

test_that("get_table_xml produces valid OOXML with the expected shape", {
  xml_str <- get_table_xml(simple_fixture())

  # Fragment declares its own namespace, so it must parse standalone
  parsed <- xml2::read_xml(xml_str)
  expect_s3_class(parsed, "xml_document")

  expect_match(xml_str, "<w:tbl ", fixed = TRUE)
  expect_identical(
    lengths(regmatches(xml_str, gregexpr("<w:gridCol", xml_str))), 4L
  )
  # 1 header + 3 data rows
  expect_identical(
    lengths(regmatches(xml_str, gregexpr("<w:tr>", xml_str))), 4L
  )
})

test_that("column widths are scaled to the requested text width", {
  xml_str <- get_table_xml(simple_fixture(), text_width_twips = 9000L)
  widths <- as.integer(
    regmatches(xml_str, gregexpr('(?<=<w:gridCol w:w=")[0-9]+', xml_str,
                                 perl = TRUE))[[1]]
  )
  expect_identical(sum(widths), 9000L)
  # Ratios preserved: first column is 1.5x the others
  expect_true(abs(widths[1] / widths[2] - 2160 / 1440) < 0.01)
})

test_that("excluded columns are dropped from the XML", {
  xml_str <- get_table_xml(simple_fixture(), excluded_cols = 2L)
  expect_identical(
    lengths(regmatches(xml_str, gregexpr("<w:gridCol", xml_str))), 3L
  )
  expect_false(grepl("<w:t [^>]*>N</w:t>", xml_str))
})

test_that("merged header cells produce gridSpan and omit continuations", {
  xml_str <- get_table_xml(test_path("fixtures", "05_merged_headers.rtf"))
  expect_identical(
    lengths(regmatches(xml_str, gregexpr('<w:gridSpan w:val="2"/>', xml_str))),
    2L
  )
  # Top header row: 5 defined cells -> 3 rendered (1 plain + 2 merged)
  first_tr <- regmatches(xml_str, regexpr("<w:tr>.*?</w:tr>", xml_str))
  expect_identical(
    lengths(regmatches(first_tr, gregexpr("<w:tc>", first_tr))), 3L
  )
})

test_that("merged data-row cells produce gridSpan in XML and colspan in HTML", {
  rtf <- paste0(
    "{\\rtf1\\ansi\n\\sectd\n",
    "\\trowd\\trhdr\\cellx2000\\cellx4000\\cellx6000",
    "{\\b A\\b0\\cell}{\\b B\\b0\\cell}{\\b C\\b0\\cell}\\row\n",
    "\\trowd\\clmgf\\cellx2000\\clmrg\\cellx4000\\clmrg\\cellx6000",
    "{Spanning label\\cell}{\\cell}{\\cell}\\row\n",
    "\\sect}"
  )
  path <- tempfile(fileext = ".rtf")
  on.exit(unlink(path), add = TRUE)
  writeLines(rtf, path)

  xml_str <- get_table_xml(path)
  expect_match(xml_str, '<w:gridSpan w:val="3"/>', fixed = TRUE)

  html_out <- get_table_html_output(path)
  expect_match(html_out, 'colspan="3"', fixed = TRUE)

  # Selection pane: merged cell carries all spanned column indices
  html_sel <- get_table_html_selection(path)
  expect_match(html_sel, "data-col='[1,2,3]'", fixed = TRUE)
})

test_that("cell text is XML-escaped", {
  expect_identical(xml_escape("a<b & \"c\">"),
                   "a&lt;b &amp; &quot;c&quot;&gt;")
})

test_that("HTML output has table structure and escapes content", {
  html <- get_table_html(simple_fixture())
  expect_match(html, "<table", fixed = TRUE)
  expect_match(html, "<thead>", fixed = TRUE)
  expect_match(html, "Treatment Group", fixed = TRUE)
  expect_identical(htmlEscape("a<b>&c"), "a&lt;b&gt;&amp;c")
})

test_that("long tables are truncated in the HTML preview", {
  mk_block <- function(texts, is_header) {
    n <- length(texts)
    list(
      text      = matrix(texts, n, 1L),
      align     = matrix(NA_character_, n, 1L),
      colspan   = matrix(1L, n, 1L),
      width     = matrix(1000L, n, 1L),
      present   = matrix(TRUE, n, 1L),
      is_header = rep(is_header, n),
      row_id    = rep(NA_integer_, n)
    )
  }
  combined <- list(
    header = mk_block("H", TRUE),
    data   = mk_block(paste0("r", seq_len(250L)), FALSE),
    col_widths_twips = 1000
  )
  html <- build_html(combined)
  expect_match(html, "50 rows not shown", fixed = TRUE)
  expect_false(grepl(">r201<", html, fixed = TRUE))
})
