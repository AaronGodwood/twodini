# docx.R - .docx bookmark extraction and Word XML table injection
# A .docx is a ZIP archive; we unzip, edit word/document.xml, rezip.

# HELPERS

# Namespace URIs
W_NS   <- "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS <- "http://schemas.openxmlformats.org/package/2006/relationships"
IMG_REL_TYPE <- "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

# EMU (English Metric Units) per twip: 1 twip = 635 EMU
TWIPS_TO_EMU <- 635L
# EMU per inch
EMU_PER_INCH <- 914400L

docx_rezip <- function(tmp_dir, output_path) {
  # List all files recursively (including hidden like _rels/.rels).
  # Use mirror mode so subdirectory paths are preserved inside the zip.
  # include_directories = FALSE avoids empty dir entries that confuse Word.
  all_files <- list.files(tmp_dir, recursive = TRUE, full.names = FALSE,
                           all.files = TRUE)
  zip::zip(zipfile = output_path, files = all_files, root = tmp_dir,
           mode = "mirror", include_directories = FALSE)
}

# Walk up the xml2 node tree to find the nearest ancestor (or self) named `tag`
find_ancestor <- function(node, tag) {
  current <- node
  for (i in seq_len(20L)) {
    if (xml_name(current) == tag) return(current)
    parent <- xml_parent(current)
    if (inherits(parent, "xml_document")) return(NULL)
    current <- parent
  }
  NULL
}

# BOOKMARK EXTRACTION

#' Extract bookmark names from a Word document
#'
#' @param docx_path Path to the .docx file
#' @return Named character vector: bookmark name -> bookmark id
extract_bookmarks <- function(docx_path) {
  tmp <- tempfile()
  dir.create(tmp)
  on.exit(unlink(tmp, recursive = TRUE), add = TRUE)
  unzip(docx_path, exdir = tmp)

  doc <- read_xml(file.path(tmp, "word", "document.xml"))
  ns  <- c(w = W_NS)

  # xml_attr() does not accept an ns argument for attribute lookup -
  # pull name and id via XPath string() instead, which resolves prefixes correctly
  nodes    <- xml_find_all(doc, ".//w:bookmarkStart", ns = ns)
  if (length(nodes) == 0L) return(character())

  bm_names <- xml_attr(nodes, "name")   # attribute is un-prefixed on the element
  bm_ids   <- xml_attr(nodes, "id")

  keep <- !is.na(bm_names) & nchar(bm_names) > 0L & !startsWith(bm_names, "_")
  setNames(bm_ids[keep], bm_names[keep])
}

# TEXT WIDTH DETECTION

# Read the text width of the document in EMU from word/document.xml sectPr.
# Falls back to 6 inches (A4 with standard margins) if not found.
docx_text_width_emu <- function(doc) {
  # sectPr holds pgSz (page size) and pgMar (margins)
  pg_sz  <- xml_find_first(doc, ".//w:sectPr/w:pgSz",  ns = c(w = W_NS))
  pg_mar <- xml_find_first(doc, ".//w:sectPr/w:pgMar", ns = c(w = W_NS))

  if (!is.na(pg_sz) && !is.na(pg_mar)) {
    page_w  <- as.integer(xml_attr(pg_sz,  "w:w",     ns = c(w = W_NS)))
    mar_l   <- as.integer(xml_attr(pg_mar, "w:left",  ns = c(w = W_NS)))
    mar_r   <- as.integer(xml_attr(pg_mar, "w:right", ns = c(w = W_NS)))

    if (!any(is.na(c(page_w, mar_l, mar_r)))) {
      # All values are in twips
      text_width_twips <- page_w - mar_l - mar_r
      return(text_width_twips * TWIPS_TO_EMU)
    }
  }

  # Fallback: 6 inches
  6L * EMU_PER_INCH
}

# IMAGE RELATIONSHIP MANAGEMENT

# Add a PNG file to word/media/ and register it in word/_rels/document.xml.rels.
# Returns the relationship ID (e.g. "rId10") for use in the drawing XML.
# Mutates session$next_img_id (counter) and session$rels_doc (live xml2 doc).
add_image_relationship <- function(session, png_bytes) {
  # Write PNG bytes to word/media/
  img_id  <- session$next_img_id
  img_name <- sprintf("image%d.png", img_id)
  img_path <- file.path(session$tmp_dir, "word", "media", img_name)
  dir.create(dirname(img_path), showWarnings = FALSE, recursive = TRUE)
  writeBin(png_bytes, img_path)

  session$next_img_id <- img_id + 1L

  # Choose a relationship ID that doesn't clash with existing ones
  rel_id <- sprintf("rIdImg%d", img_id)

  # Add <Relationship> to word/_rels/document.xml.rels
  new_rel <- sprintf(
    '<Relationship xmlns="%s" Id="%s" Type="%s" Target="media/%s"/>',
    REL_NS, rel_id, IMG_REL_TYPE, img_name
  )
  xml_add_child(xml_root(session$rels_doc), read_xml(new_rel))

  # Ensure [Content_Types].xml has a Default entry for png
  CT_NS <- "http://schemas.openxmlformats.org/package/2006/content-types"
  existing_png <- xml_find_first(
    session$ct_doc,
    ".//ct:Default[@Extension='png']",
    ns = c(ct = CT_NS)
  )
  if (inherits(existing_png, "xml_missing")) {
    new_ct <- sprintf(
      '<Default xmlns="%s" Extension="png" ContentType="image/png"/>',
      CT_NS
    )
    xml_add_child(xml_root(session$ct_doc), read_xml(new_ct))
  }

  rel_id
}

# DOCUMENT SESSION - open / inject / close

#' Open a Word document for editing
#'
#' Unzips the .docx, parses word/document.xml and its relationships file into
#' memory, detects text width, and builds a bookmark jump table.
#' Call inject_table() / inject_image() any number of times, then close_docx().
#'
#' @param docx_path Path to the source .docx file
#' @return A docx session environment
open_docx <- function(docx_path) {
  tmp <- tempfile()
  dir.create(tmp)
  unzip(docx_path, exdir = tmp)

  xml_path   <- file.path(tmp, "word", "document.xml")
  rels_path  <- file.path(tmp, "word", "_rels", "document.xml.rels")
  ct_path    <- file.path(tmp, "[Content_Types].xml")
  doc        <- read_xml(xml_path)
  rels_doc   <- read_xml(rels_path)
  ct_doc     <- read_xml(ct_path)

  # Build bookmark jump table once: name -> <w:p> node reference
  bm_nodes   <- xml_find_all(doc, ".//w:bookmarkStart", ns = c(w = W_NS))
  jump_table <- list()
  for (bm in bm_nodes) {
    name <- xml_attr(bm, "name")
    if (is.na(name) || startsWith(name, "_")) next
    para <- find_ancestor(bm, "p")
    if (!is.null(para)) jump_table[[name]] <- para
  }

  # Detect next available image index from existing media files
  media_dir   <- file.path(tmp, "word", "media")
  existing    <- if (dir.exists(media_dir)) list.files(media_dir, pattern = "\\.png$") else character()
  next_img_id <- length(existing) + 1L

  # Text width in EMU (read once, used by every inject_image call)
  text_width_emu <- docx_text_width_emu(doc)

  session <- new.env(parent = emptyenv())
  session$tmp_dir        <- tmp
  session$xml_path       <- xml_path
  session$rels_path      <- rels_path
  session$ct_path        <- ct_path
  session$doc            <- doc
  session$rels_doc       <- rels_doc
  session$ct_doc         <- ct_doc
  session$jump_table     <- jump_table
  session$next_img_id    <- next_img_id
  session$text_width_emu <- text_width_emu
  session
}

#' Inject a Word XML table after the paragraph containing a bookmark
#'
#' @param session A docx session returned by open_docx()
#' @param bookmark_name Name of the bookmark
#' @param xml_string A <w:tbl> XML string (from get_table_xml())
#' @return Invisibly TRUE on success, FALSE if bookmark not found
inject_table <- function(session, bookmark_name, xml_string) {
  para_node <- session$jump_table[[bookmark_name]]

  if (is.null(para_node)) {
    warning(sprintf("Bookmark '%s' not found in document.", bookmark_name))
    return(invisible(FALSE))
  }

  xml_add_sibling(para_node, read_xml(xml_string), .where = "after")
  invisible(TRUE)
}

#' Inject a PNG image after the paragraph containing a bookmark
#'
#' Writes the PNG into word/media/, registers the relationship, then inserts
#' a minimal <w:drawing> block. The image is scaled to the document text width
#' maintaining the original aspect ratio from the RTF \picwgoal / \pichgoal.
#'
#' @param session A docx session returned by open_docx()
#' @param bookmark_name Name of the bookmark
#' @param png_bytes raw vector of PNG bytes (from extract_png())
#' @param width_twips Original image width in twips (from extract_png())
#' @param height_twips Original image height in twips (from extract_png())
#' @return Invisibly TRUE on success, FALSE if bookmark not found
inject_image <- function(session, bookmark_name, png_bytes, width_twips, height_twips) {
  para_node <- session$jump_table[[bookmark_name]]

  if (is.null(para_node)) {
    warning(sprintf("Bookmark '%s' not found in document.", bookmark_name))
    return(invisible(FALSE))
  }

  # Scale to text width maintaining aspect ratio
  target_w_emu <- session$text_width_emu
  if (!is.na(width_twips) && !is.na(height_twips) && width_twips > 0L) {
    aspect       <- height_twips / width_twips
    target_h_emu <- round(target_w_emu * aspect)
  } else {
    # No dimension info - fall back to a square
    target_h_emu <- target_w_emu
  }

  # Capture id before add_image_relationship increments the counter
  img_id <- session$next_img_id
  rel_id <- add_image_relationship(session, png_bytes)

  # Minimal <w:drawing> / <wp:inline> XML
  # Namespaces declared inline so the fragment is self-contained when parsed
  drawing_xml <- sprintf(
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
       <w:r>
         <w:drawing>
           <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
             <wp:extent cx="%d" cy="%d"/>
             <wp:docPr id="%d" name="Image%d"/>
             <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
               <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                 <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                   <pic:nvPicPr>
                     <pic:cNvPr id="%d" name="Image%d"/>
                     <pic:cNvPicPr/>
                   </pic:nvPicPr>
                   <pic:blipFill>
                     <a:blip r:embed="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                     <a:stretch><a:fillRect/></a:stretch>
                   </pic:blipFill>
                   <pic:spPr>
                     <a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>
                     <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                   </pic:spPr>
                 </pic:pic>
               </a:graphicData>
             </a:graphic>
           </wp:inline>
         </w:drawing>
       </w:r>
     </w:p>',
    target_w_emu, target_h_emu,  # wp:extent
    img_id, img_id,              # wp:docPr id + name
    img_id, img_id,              # pic:cNvPr id + name
    rel_id,                      # r:embed
    target_w_emu, target_h_emu   # a:ext
  )

  xml_add_sibling(para_node, read_xml(drawing_xml), .where = "after")
  invisible(TRUE)
}

#' Finalise and save a Word document session
#'
#' Serialises document.xml and the updated relationships file, rezips the
#' .docx, and cleans up the temp directory.
#'
#' @param session A docx session returned by open_docx()
#' @param output_path Path to write the finished .docx
close_docx <- function(session, output_path) {
  write_xml(session$doc,      session$xml_path)
  write_xml(session$rels_doc, session$rels_path)
  write_xml(session$ct_doc,   session$ct_path)
  docx_rezip(session$tmp_dir, output_path)
  unlink(session$tmp_dir, recursive = TRUE)
  invisible(output_path)
}

# DOCUMENT PROCESSING (calls open / inject / close)

#' Process a Word document: inject all configured RTF content at bookmarks
#'
#' Auto-detects whether each RTF file contains an image (\\pngblip) or a
#' table and calls the appropriate inject function.
#'
#' @param word_path Path to the source .docx file
#' @param config data.frame with columns: bookmark, tablename
#' @param rtf_paths Named list: table_name -> rtf file path
#' @param selections Named list: row_index (as character) -> list(cols, row_start,
#'   row_end, parameters, timelines)
#' @param output_path Path to write the final .docx
process_document <- function(word_path, config, rtf_paths, selections, output_path) {
  session <- open_docx(word_path)
  on.exit(unlink(session$tmp_dir, recursive = TRUE), add = TRUE)

  # row index (as character) -> "ok" | error message string
  status <- vector("list", nrow(config))
  names(status) <- as.character(seq_len(nrow(config)))

  for (i in seq_len(nrow(config))) {
    bm_name  <- config$Bookmark[i]
    tbl_name <- config$Table[i]

    if (!tbl_name %in% names(rtf_paths)) {
      status[[i]] <- "RTF file not found in folder"
      next
    }

    rtf_path <- rtf_paths[[tbl_name]]

    if (is_image_rtf(rtf_path)) {
      img <- extract_png(rtf_path)
      if (is.null(img)) {
        status[[i]] <- "Could not extract PNG from RTF"
        next
      }
      tryCatch(
        inject_image(session, bm_name, img$png_bytes, img$width_twips, img$height_twips),
        error = function(e) status[[i]] <<- conditionMessage(e)
      )

    } else {
      sel <- selections[[as.character(i)]]
      tw_twips <- round(session$text_width_emu / TWIPS_TO_EMU)
      tryCatch({
        xml_str <- get_table_xml(
          rtf_path,
          excluded_cols        = sel$excluded_cols,
          excluded_rows        = sel$excluded_rows,
          excluded_header_rows = sel$excluded_header_rows,
          parameters           = sel$parameters,
          timelines            = sel$timelines,
          text_width_twips     = tw_twips
        )
        inject_table(session, bm_name, xml_str)
      }, error = function(e) status[[i]] <<- conditionMessage(e))
    }
  }

  close_docx(session, output_path)
  invisible(status)
}
