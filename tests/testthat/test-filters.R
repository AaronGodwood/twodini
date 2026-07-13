# Parameter/timeline filtering, page combining, and the shared column/row
# resolution helpers used by the HTML and XML renderers.

test_that("filter_pages keeps matching and NA-parameter pages", {
  pages <- parse_rtf(test_path("fixtures", "02_multi_param.rtf"))

  expect_length(filter_pages(pages, "Cholesterol"), 1L)
  expect_length(filter_pages(pages, c("Cholesterol", "HDL")), 2L)
  # NULL / empty filter means keep everything
  expect_length(filter_pages(pages, NULL), 3L)
  expect_length(filter_pages(pages, character()), 3L)

  # Pages without a parameter are always kept
  no_param <- parse_rtf(test_path("fixtures", "05_merged_headers.rtf"))
  expect_length(filter_pages(no_param, "Anything"), 1L)
})

test_that("filter_timelines keeps only the selected blocks", {
  pages <- parse_rtf(test_path("fixtures", "03_multi_timeline.rtf"))

  kept <- filter_timelines(pages, "Week 4")
  # Page 1 holds Week 1 and Week 4; only the Week 4 block (label + 2 rows) stays
  expect_length(kept[[1]]$data_rows, 3L)
  expect_identical(kept[[1]]$data_rows[[1]]$cells[[1]]$text, "Week 4")
  # Page 2 (Week 8 / Week 12) loses everything
  expect_length(kept[[2]]$data_rows, 0L)

  # No filter -> unchanged
  expect_identical(filter_timelines(pages, NULL), pages)
})

test_that("merged timeline label rows are detected and filterable", {
  rtf <- paste0(
    "{\\rtf1\\ansi\n\\sectd\n",
    "\\trowd\\trhdr\\cellx2000\\cellx4000\\cellx6000",
    "{\\b Treat\\b0\\cell}{\\b N\\b0\\cell}{\\b Mean\\b0\\cell}\\row\n",
    "\\trowd\\clmgf\\cellx2000\\clmrg\\cellx4000\\clmrg\\cellx6000",
    "{Week 1\\cell}{\\cell}{\\cell}\\row\n",
    "\\trowd\\cellx2000\\cellx4000\\cellx6000{Placebo\\cell}{50\\cell}{5.2\\cell}\\row\n",
    "\\trowd\\clmgf\\cellx2000\\clmrg\\cellx4000\\clmrg\\cellx6000",
    "{Week 2\\cell}{\\cell}{\\cell}\\row\n",
    "\\trowd\\cellx2000\\cellx4000\\cellx6000{Placebo\\cell}{48\\cell}{4.9\\cell}\\row\n",
    "\\sect}"
  )
  path <- tempfile(fileext = ".rtf")
  on.exit(unlink(path), add = TRUE)
  writeLines(rtf, path)

  pages <- parse_rtf(path)
  expect_identical(sort(get_timelines(pages)), c("Week 1", "Week 2"))

  kept <- filter_timelines(pages, "Week 2")
  expect_length(kept[[1]]$data_rows, 2L)
  expect_identical(kept[[1]]$data_rows[[1]]$cells[[1]]$text, "Week 2")
})

test_that("combine_pages concatenates data and keeps first-page headers", {
  pages <- parse_rtf(test_path("fixtures", "02_multi_param.rtf"))
  combined <- combine_pages(pages)

  expect_length(combined$header_rows, 1L)
  expect_length(combined$data_rows, 6L)  # 3 pages x 2 rows
  expect_identical(combined$col_widths_twips, c(2160, 1440, 1440, 1440))

  empty <- combine_pages(list())
  expect_length(empty$header_rows, 0L)
  expect_length(empty$data_rows, 0L)
})

test_that("resolve_cols handles indices, names and out-of-range input", {
  pages    <- parse_rtf(test_path("fixtures", "01_simple.rtf"))
  combined <- combine_pages(pages)
  hdrs     <- combined$header_rows

  expect_identical(resolve_cols(NULL, 4L, hdrs), 1:4)
  expect_identical(resolve_cols(c(1L, 3L), 4L, hdrs), c(1L, 3L))
  expect_identical(resolve_cols(c("N", "SD"), 4L, hdrs), c(2L, 4L))
  # Out-of-range indices are dropped
  expect_identical(resolve_cols(c(2L, 99L), 4L, hdrs), 2L)
  # Nothing valid -> all columns
  expect_identical(resolve_cols(99L, 4L, hdrs), 1:4)
})

test_that("slice_rows respects bounds", {
  rows <- as.list(1:5)
  expect_identical(slice_rows(rows, 2, 4), as.list(2:4))
  expect_identical(slice_rows(rows, NULL, NULL), rows)
  expect_identical(slice_rows(rows, 4, 2), list())
  expect_identical(slice_rows(list(), 1, 10), list())
})

test_that("prepare_table applies exclusions", {
  prep <- prepare_table(
    path = test_path("fixtures", "01_simple.rtf"),
    excluded_cols = 2L,
    excluded_rows = c(1L, 3L)
  )
  expect_identical(prep$included_cols, c(1L, 3L, 4L))
  expect_length(prep$combined$data_rows, 1L)
  expect_identical(prep$combined$data_rows[[1]]$cells[[1]]$text, "Treatment A")
})
