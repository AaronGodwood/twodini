# PNG detection and extraction from image RTFs.

image_fixture <- function() test_path("fixtures", "06_image.rtf")

png_signature <- as.raw(c(0x89, 0x50, 0x4e, 0x47))

test_that("image RTFs are detected", {
  expect_true(is_image_rtf(image_fixture()))
  expect_false(is_image_rtf(test_path("fixtures", "01_simple.rtf")))
})

test_that("extract_png returns bytes and declared dimensions", {
  img <- extract_png(image_fixture())
  expect_type(img, "list")
  expect_identical(img$png_bytes[1:4], png_signature)
  expect_identical(img$width_twips, 5760L)
  expect_identical(img$height_twips, 4320L)
})

test_that("extract_png returns NULL when no PNG is present", {
  expect_null(extract_png(test_path("fixtures", "01_simple.rtf")))
})

test_that("blipuid groups and non-PNG picts don't corrupt extraction", {
  png_hex <- paste0(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489",
    "0000000d4944415478da63fcffff3f030005fe02fea72d994e",
    "0000000049454e44ae426082"
  )
  rtf <- sprintf(paste0(
    "{\\rtf1\\ansi\n",
    # decoy WMF rendition before the PNG
    "{\\pict\\wmetafile8\\picw100\\pich100 0102030405}\n",
    "{\\pict{\\*\\picprop}\\pngblip\\picwgoal2880\\pichgoal1440\n",
    "{\\*\\blipuid deadbeefdeadbeefdeadbeefdeadbeef}\n%s\n}}"
  ), png_hex)
  path <- tempfile(fileext = ".rtf")
  on.exit(unlink(path), add = TRUE)
  writeLines(rtf, path)

  img <- extract_png(path)
  expect_identical(img$png_bytes[1:4], png_signature)
  # blipuid hex must not be prepended to the image bytes
  expect_identical(length(img$png_bytes), nchar(png_hex) %/% 2L)
  expect_identical(img$width_twips, 2880L)
})

test_that("hex_to_raw round-trips", {
  bytes <- as.raw(c(0x00, 0x0f, 0xf0, 0xff, 0xab))
  expect_identical(hex_to_raw(paste(sprintf("%02x", as.integer(bytes)),
                                    collapse = "")),
                   bytes)
})
