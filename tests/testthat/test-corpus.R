# Optional corpus regression: parses every RTF in the repo-level test_data
# folder (outside the package, so this only runs on dev machines that have it;
# CI and R CMD check skip it automatically).

test_that("full local RTF corpus parses without artifacts", {
  corpus <- file.path("..", "..", "..", "test_data", "rtf")
  skip_if_not(dir.exists(corpus), "local test_data corpus not present")

  files <- list.files(corpus, pattern = "\\.rtf$", full.names = TRUE,
                      ignore.case = TRUE)
  expect_gt(length(files), 0L)

  for (f in files) {
    if (is_image_rtf(f)) {
      img <- extract_png(f)
      expect_false(is.null(img), label = paste("PNG extracted from", basename(f)))
      next
    }

    pages <- parse_rtf(f)
    expect_gte(length(pages), 1L)

    combined <- combine_pages(pages)
    for (b in list(combined$header, combined$data)) {
      for (txt in b$text[b$present]) {
        expect_false(
          grepl("[�□]", txt, perl = TRUE) ||
            grepl("\\\\[a-zA-Z]{2,}", txt, perl = TRUE),
          label = sprintf("artifact-free cell in %s ('%s')",
                          basename(f), txt)
        )
      }
    }
  }
})
