# Word document handling: bookmark extraction, table/image injection,
# and the end-to-end process_document() pipeline.
# All tests build a minimal .docx from scratch (see helper-docx.R).

test_that("extract_bookmarks lists names and skips internal bookmarks", {
  docx <- make_min_docx(bookmarks = c("TableBM", "ImageBM", "_GoBack"))
  on.exit(unlink(docx), add = TRUE)

  bm <- extract_bookmarks(docx)
  expect_identical(sort(names(bm)), c("ImageBM", "TableBM"))
})

test_that("open_docx reads text width and builds the bookmark jump table", {
  docx <- make_min_docx()
  on.exit(unlink(docx), add = TRUE)

  session <- open_docx(docx)
  on.exit(unlink(session$tmp_dir, recursive = TRUE), add = TRUE)

  # 12240 - 1800 - 1800 = 8640 twips of text width
  expect_identical(session$text_width_emu, 8640L * 635L)
  expect_identical(sort(names(session$jump_table)), c("ImageBM", "TableBM"))
  expect_identical(session$next_img_id, 1L)
})

test_that("next image id skips past non-contiguous media names", {
  docx <- make_min_docx(media = list(
    "image1.png" = as.raw(1:4),
    "image3.png" = as.raw(1:4)
  ))
  on.exit(unlink(docx), add = TRUE)

  session <- open_docx(docx)
  on.exit(unlink(session$tmp_dir, recursive = TRUE), add = TRUE)
  expect_identical(session$next_img_id, 4L)
})

test_that("inject_table places a table after the bookmark paragraph", {
  docx <- make_min_docx()
  out  <- tempfile(fileext = ".docx")
  on.exit(unlink(c(docx, out)), add = TRUE)

  session <- open_docx(docx)
  xml_str <- get_table_xml(test_path("fixtures", "01_simple.rtf"),
                           text_width_twips = 8640L)
  inject_table(session, "TableBM", xml_str)
  close_docx(session, out)

  doc <- read_docx_document(out)
  expect_match(doc, "<w:tbl", fixed = TRUE)
  expect_match(doc, "Treatment Group", fixed = TRUE)
  # Injected table sits after the bookmark paragraph
  expect_true(regexpr("TableBM", doc, fixed = TRUE) <
                regexpr("<w:tbl", doc, fixed = TRUE))
})

test_that("inject_table errors with a houdini error for unknown bookmarks", {
  docx <- make_min_docx()
  on.exit(unlink(docx), add = TRUE)
  session <- open_docx(docx)
  on.exit(unlink(session$tmp_dir, recursive = TRUE), add = TRUE)

  expect_error(inject_table(session, "Nope", "<w:tbl></w:tbl>"),
               class = "houdini_error_bookmark_missing")
})

test_that("inject_image writes media, relationship and content type", {
  docx <- make_min_docx()
  out  <- tempfile(fileext = ".docx")
  on.exit(unlink(c(docx, out)), add = TRUE)

  img <- extract_png(test_path("fixtures", "06_image.rtf"))
  session <- open_docx(docx)
  inject_image(session, "ImageBM", img$png_bytes,
               img$width_twips, img$height_twips)
  close_docx(session, out)

  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  unzip(out, exdir = tmp)

  expect_true(file.exists(file.path(tmp, "word", "media", "image1.png")))
  rels <- paste(readLines(file.path(tmp, "word", "_rels", "document.xml.rels"),
                          warn = FALSE), collapse = "")
  expect_match(rels, 'Id="rIdImg1"', fixed = TRUE)
  ct <- paste(readLines(file.path(tmp, "[Content_Types].xml"), warn = FALSE),
              collapse = "")
  expect_match(ct, 'Extension="png"', fixed = TRUE)

  doc <- paste(readLines(file.path(tmp, "word", "document.xml"), warn = FALSE),
               collapse = "\n")
  expect_match(doc, "<w:drawing>", fixed = TRUE)
  # Scaled to text width preserving 2:1 aspect (5760x4320 -> h = w * 0.75)
  w_emu <- 8640L * 635L
  expect_match(doc, sprintf('cx="%d"', w_emu), fixed = TRUE)
  expect_match(doc, sprintf('cy="%d"', round(w_emu * 4320 / 5760)), fixed = TRUE)
})

test_that("process_document injects everything and reports per-row status", {
  docx <- make_min_docx(bookmarks = c("TableBM", "ImageBM"))
  out  <- tempfile(fileext = ".docx")
  on.exit(unlink(c(docx, out)), add = TRUE)

  config <- data.frame(
    Bookmark = c("TableBM", "ImageBM"),
    Table    = c("tbl", "img"),
    stringsAsFactors = FALSE
  )
  rtf_paths <- list(
    tbl = test_path("fixtures", "01_simple.rtf"),
    img = test_path("fixtures", "06_image.rtf")
  )
  selections <- list("1" = list(excluded_cols = 2L))

  status <- process_document(docx, config, rtf_paths, selections, out)

  expect_true(all(vapply(status, is.null, logical(1))))
  expect_true(file.exists(out))

  doc <- read_docx_document(out)
  expect_match(doc, "<w:tbl", fixed = TRUE)
  expect_match(doc, "<w:drawing>", fixed = TRUE)
  # Excluded column 2 ("N") must not appear in the injected table
  expect_false(grepl(">N</w:t>", doc, fixed = TRUE))
})

test_that("process_document records errors instead of aborting", {
  docx <- make_min_docx(bookmarks = "TableBM")
  out  <- tempfile(fileext = ".docx")
  on.exit(unlink(c(docx, out)), add = TRUE)

  config <- data.frame(
    Bookmark = c("TableBM", "MissingBM"),
    Table    = c("nope", "tbl"),
    stringsAsFactors = FALSE
  )
  rtf_paths <- list(tbl = test_path("fixtures", "01_simple.rtf"))

  status <- process_document(docx, config, rtf_paths, list(), out)

  expect_s3_class(status[["1"]], "houdini_error_rtf_unreadable")
  expect_s3_class(status[["2"]], "houdini_error_xml_inject_failed")
  expect_true(file.exists(out))
})
