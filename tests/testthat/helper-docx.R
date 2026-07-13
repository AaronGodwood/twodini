# Build a minimal but valid .docx containing named bookmarks, so docx tests
# don't depend on a binary template checked into the repo.
#
# bookmarks: character vector of bookmark names, one paragraph each.
# media: optional named list of file name -> raw bytes to place in word/media/.
# Returns the path to the created .docx.
make_min_docx <- function(path = tempfile(fileext = ".docx"),
                          bookmarks = c("TableBM", "ImageBM"),
                          media = NULL) {
  tmp <- tempfile()
  dir.create(file.path(tmp, "_rels"), recursive = TRUE)
  dir.create(file.path(tmp, "word", "_rels"), recursive = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    "</Types>"
  ), file.path(tmp, "[Content_Types].xml"))

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
    "</Relationships>"
  ), file.path(tmp, "_rels", ".rels"))

  paras <- vapply(seq_along(bookmarks), function(i) {
    sprintf(paste0(
      '<w:p><w:bookmarkStart w:id="%d" w:name="%s"/>',
      '<w:bookmarkEnd w:id="%d"/>',
      "<w:r><w:t>%s placeholder</w:t></w:r></w:p>"
    ), i - 1L, bookmarks[i], i - 1L, bookmarks[i])
  }, character(1))

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
    "<w:body>",
    paras,
    # 8.5x11in page, 1.25in side margins -> 8640 twips of text width
    '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>',
    '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>',
    "</w:body>",
    "</w:document>"
  ), file.path(tmp, "word", "document.xml"))

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
  ), file.path(tmp, "word", "_rels", "document.xml.rels"))

  for (nm in names(media)) {
    media_dir <- file.path(tmp, "word", "media")
    dir.create(media_dir, showWarnings = FALSE, recursive = TRUE)
    writeBin(media[[nm]], file.path(media_dir, nm))
  }

  all_files <- list.files(tmp, recursive = TRUE, full.names = FALSE,
                          all.files = TRUE)
  zip::zip(zipfile = path, files = all_files, root = tmp,
           mode = "mirror", include_directories = FALSE)
  unlink(tmp, recursive = TRUE)
  path
}

# Unzip a .docx and read word/document.xml as text (for content assertions)
read_docx_document <- function(docx_path) {
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  unzip(docx_path, exdir = tmp)
  paste(readLines(file.path(tmp, "word", "document.xml"), warn = FALSE),
        collapse = "\n")
}
