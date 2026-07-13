# tests/make_test_pdfs.R
#
# Generates synthetic PDF files that exercise the Phase-1 parser
# (R/pdf_parse.R). Hand-built from raw bytes — no external PDF library —
# so test failures reflect parser bugs, not library quirks.
#
# Usage:
#   source("tests/make_test_pdfs.R")
#   make_test_pdfs("C:/path/to/output/folder")
#
# What gets generated (one file each):
#   - simple_one_page.pdf    : one page, traditional xref, bfchar CMap
#   - two_pages.pdf          : two pages, shared font, tests page tree
#   - bfrange_cmap.pdf       : one page, CMap uses bfrange (single-target + array form)
#   - flate_content.pdf      : one page, content stream FlateDecode-compressed
#
# Each generator returns the expected (page -> decoded text) mapping so the
# smoke test can diff against what pdf_page_streams + a naive decoder give.

# ---- 1. LOW-LEVEL PDF BUILDING HELPERS ----

# Accumulate bytes into a growable buffer with offset tracking.
# Each `pdf_add()` returns the 1-based byte offset where the chunk started —
# needed because xref entries must point at the start of each indirect object.
pdf_buf_new <- function() {
  list(bytes = raw(0), pos = 0L)
}

pdf_buf_add <- function(buf, chunk) {
  if (is.character(chunk)) chunk <- charToRaw(chunk)
  start <- buf$pos + 1L
  buf$bytes <- c(buf$bytes, chunk)
  buf$pos <- buf$pos + length(chunk)
  list(buf = buf, start = start)
}

# Write an indirect object "N G obj\n<body>\nendobj\n".
# Returns updated buffer and the offset of "N" (xref target).
pdf_write_object <- function(buf, obj_num, body) {
  r <- pdf_buf_add(buf, sprintf("%d 0 obj\n", obj_num))
  r <- pdf_buf_add(r$buf, body)
  if (is.raw(body) || !endsWith(if (is.raw(body)) "" else body, "\n")) {
    r <- pdf_buf_add(r$buf, "\n")
  }
  r <- pdf_buf_add(r$buf, "endobj\n")
  list(buf = r$buf, offset = buf$pos + 1L)
}

# Write a stream object. `dict_extras` goes inside the << >> alongside /Length.
# `stream_bytes` must be raw. Applies no filter unless caller puts /Filter in
# dict_extras (and pre-compresses stream_bytes).
pdf_write_stream_object <- function(buf, obj_num, dict_extras, stream_bytes) {
  stopifnot(is.raw(stream_bytes))
  body_header <- sprintf("<<%s /Length %d>>\nstream\n",
                         dict_extras, length(stream_bytes))
  r <- pdf_buf_add(buf, sprintf("%d 0 obj\n", obj_num))
  r <- pdf_buf_add(r$buf, body_header)
  r <- pdf_buf_add(r$buf, stream_bytes)
  r <- pdf_buf_add(r$buf, "\nendstream\nendobj\n")
  list(buf = r$buf, offset = buf$pos + 1L)
}

# Compose the trailing xref table + trailer + startxref + %%EOF.
# `offsets` is a named integer-ish vector: names = obj numbers (as character),
# values = byte offset of "N 0 obj" for that object.
pdf_finalise <- function(buf, offsets, root_obj, info_obj = NULL) {
  xref_offset <- buf$pos + 1L

  obj_nums <- sort(as.integer(names(offsets)))
  size <- max(obj_nums) + 1L  # xref size includes object 0

  # Build xref table. Object 0 is always "0000000000 65535 f"
  lines <- c(
    sprintf("xref\n0 %d\n", size),
    sprintf("%010d %05d f \n", 0L, 65535L)
  )
  for (n in seq_len(size - 1L)) {
    off <- offsets[[as.character(n)]]
    if (is.null(off)) {
      # Free object
      lines <- c(lines, sprintf("%010d %05d f \n", 0L, 0L))
    } else {
      lines <- c(lines, sprintf("%010d %05d n \n", off - 1L, 0L))
    }
  }
  # PDF offsets are 0-based, our buffer is 1-based — hence off - 1L

  # Trailer
  info_str <- if (!is.null(info_obj)) sprintf(" /Info %d 0 R", info_obj) else ""
  trailer <- sprintf("trailer\n<< /Size %d /Root %d 0 R%s >>\nstartxref\n%d\n%%%%EOF\n",
                     size, root_obj, info_str, xref_offset - 1L)

  r <- pdf_buf_add(buf, paste(lines, collapse = ""))
  r <- pdf_buf_add(r$buf, trailer)
  r$buf$bytes
}

# FlateDecode (zlib) compression via memCompress.
pdf_flate <- function(bytes) {
  if (is.character(bytes)) bytes <- charToRaw(bytes)
  memCompress(bytes, type = "gzip")
}

# ---- 2. MINIMAL TOUNICODE CMAP BUILDERS ----

# A ToUnicode CMap stream is itself a tiny PostScript program. We emit the
# minimum skeleton plus the per-glyph mapping table the parser needs.
cmap_header <- paste(
  "/CIDInit /ProcSet findresource begin",
  "12 dict begin",
  "begincmap",
  "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def",
  "/CMapName /Adobe-Identity-UCS def",
  "/CMapType 2 def",
  "1 begincodespacerange <00> <FF> endcodespacerange",
  sep = "\n"
)

cmap_footer <- paste(
  "endcmap",
  "CMapName currentdict /CMap defineresource pop",
  "end end",
  sep = "\n"
)

# Build a bfchar-based CMap mapping the given (glyph_code -> unicode) pairs.
# codes: integer vector of glyph codes (0-255)
# unis : integer vector of unicode code points
make_bfchar_cmap <- function(codes, unis) {
  stopifnot(length(codes) == length(unis))
  entries <- sprintf("<%02X> <%04X>", codes, unis)
  body <- sprintf("%d beginbfchar\n%s\nendbfchar",
                  length(codes), paste(entries, collapse = "\n"))
  paste(cmap_header, body, cmap_footer, sep = "\n")
}

# Build a bfrange-based CMap. Uses two ranges:
#   1. Contiguous single-target form:   <start> <end> <base_unicode>
#   2. Array form for a short explicit list: <start> <end> [<u1> <u2> ...]
make_bfrange_cmap <- function(range1_start, range1_end, range1_base,
                              range2_start, range2_end, range2_unis) {
  e1 <- sprintf("<%02X> <%02X> <%04X>", range1_start, range1_end, range1_base)
  arr <- paste(sprintf("<%04X>", range2_unis), collapse = " ")
  e2 <- sprintf("<%02X> <%02X> [%s]", range2_start, range2_end, arr)
  body <- sprintf("2 beginbfrange\n%s\n%s\nendbfrange", e1, e2)
  paste(cmap_header, body, cmap_footer, sep = "\n")
}

# ---- 3. CONTENT STREAM HELPERS ----

# Single-line text at (x,y) using font /F1 at given size.
# `text` is a hex string of glyph codes (e.g. "4142" = codes 0x41, 0x42).
content_show_hex <- function(x, y, size, hex) {
  sprintf("BT\n/F1 %d Tf\n%d %d Td\n<%s> Tj\nET\n", size, x, y, hex)
}

# ---- 4. PAGE + FONT OBJECT BUILDERS ----

# Produce a Type1 font dict body that references ToUnicode object `tu_obj`.
font_dict_body <- function(tu_obj) {
  sprintf("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /ToUnicode %d 0 R >>",
          tu_obj)
}

# Produce a page dict body.
page_dict_body <- function(parent_obj, contents_obj, font_obj,
                           media_box = "[0 0 612 792]") {
  sprintf(paste0(
    "<< /Type /Page /Parent %d 0 R /MediaBox %s ",
    "/Resources << /Font << /F1 %d 0 R >> >> ",
    "/Contents %d 0 R >>"
  ), parent_obj, media_box, font_obj, contents_obj)
}

# ---- 5. PDF GENERATORS ----

# Helper: finish building, return byte vector.
finalise_simple <- function(buf, offsets, root_obj) {
  pdf_finalise(buf, offsets, root_obj)
}

# Generator 1: one page, traditional xref, bfchar CMap, uncompressed content.
make_simple_one_page <- function() {
  # Object layout:
  #   1 = Catalog
  #   2 = Pages (root)
  #   3 = Page
  #   4 = Font
  #   5 = ToUnicode CMap
  #   6 = Content stream
  #
  # Glyph mapping: 0x41 -> 'H', 0x42 -> 'i'
  # Content: BT /F1 12 Tf 72 720 Td <4142> Tj ET

  buf <- pdf_buf_new()
  r <- pdf_buf_add(buf, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n")
  buf <- r$buf

  offsets <- list()

  # Object 1: Catalog
  r <- pdf_write_object(buf, 1L, "<< /Type /Catalog /Pages 2 0 R >>")
  buf <- r$buf; offsets[["1"]] <- r$offset

  # Object 2: Pages
  r <- pdf_write_object(buf, 2L, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
  buf <- r$buf; offsets[["2"]] <- r$offset

  # Object 3: Page
  r <- pdf_write_object(buf, 3L, page_dict_body(2L, 6L, 4L))
  buf <- r$buf; offsets[["3"]] <- r$offset

  # Object 4: Font
  r <- pdf_write_object(buf, 4L, font_dict_body(5L))
  buf <- r$buf; offsets[["4"]] <- r$offset

  # Object 5: ToUnicode CMap stream
  cmap_text <- make_bfchar_cmap(codes = c(0x41L, 0x42L),
                                unis  = c(utf8ToInt("H"), utf8ToInt("i")))
  r <- pdf_write_stream_object(buf, 5L, "", charToRaw(cmap_text))
  buf <- r$buf; offsets[["5"]] <- r$offset

  # Object 6: content stream
  content <- content_show_hex(72L, 720L, 12L, "4142")
  r <- pdf_write_stream_object(buf, 6L, "", charToRaw(content))
  buf <- r$buf; offsets[["6"]] <- r$offset

  list(
    bytes = finalise_simple(buf, offsets, root_obj = 1L),
    expect = list(n_pages = 1L, texts = list("Hi"))
  )
}

# Generator 2: two pages sharing the same font, tests page tree walking.
make_two_pages <- function() {
  # Object layout:
  #   1 Catalog, 2 Pages, 3 Page-A, 4 Page-B, 5 Font, 6 ToUnicode,
  #   7 Content-A ("Hi"), 8 Content-B ("iH")
  buf <- pdf_buf_new()
  r <- pdf_buf_add(buf, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"); buf <- r$buf
  offsets <- list()

  r <- pdf_write_object(buf, 1L, "<< /Type /Catalog /Pages 2 0 R >>")
  buf <- r$buf; offsets[["1"]] <- r$offset

  r <- pdf_write_object(buf, 2L, "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>")
  buf <- r$buf; offsets[["2"]] <- r$offset

  r <- pdf_write_object(buf, 3L, page_dict_body(2L, 7L, 5L))
  buf <- r$buf; offsets[["3"]] <- r$offset

  r <- pdf_write_object(buf, 4L, page_dict_body(2L, 8L, 5L))
  buf <- r$buf; offsets[["4"]] <- r$offset

  r <- pdf_write_object(buf, 5L, font_dict_body(6L))
  buf <- r$buf; offsets[["5"]] <- r$offset

  cmap_text <- make_bfchar_cmap(codes = c(0x41L, 0x42L),
                                unis  = c(utf8ToInt("H"), utf8ToInt("i")))
  r <- pdf_write_stream_object(buf, 6L, "", charToRaw(cmap_text))
  buf <- r$buf; offsets[["6"]] <- r$offset

  r <- pdf_write_stream_object(buf, 7L, "",
                               charToRaw(content_show_hex(72L, 720L, 12L, "4142")))
  buf <- r$buf; offsets[["7"]] <- r$offset

  r <- pdf_write_stream_object(buf, 8L, "",
                               charToRaw(content_show_hex(72L, 720L, 12L, "4241")))
  buf <- r$buf; offsets[["8"]] <- r$offset

  list(
    bytes = finalise_simple(buf, offsets, root_obj = 1L),
    expect = list(n_pages = 2L, texts = list("Hi", "iH"))
  )
}

# Generator 3: bfrange CMap. Exercises both single-target and array variants.
make_bfrange_cmap_pdf <- function() {
  # Mapping plan:
  #   Range 1 (contiguous): codes 0x41..0x43 -> unicode 'A'..'C'   (single-target)
  #   Range 2 (array):      codes 0x61..0x62 -> ['z', 'y']         (array form)
  #
  # Content uses hex string <414243 6162> expected text "ABCzy"
  buf <- pdf_buf_new()
  r <- pdf_buf_add(buf, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"); buf <- r$buf
  offsets <- list()

  r <- pdf_write_object(buf, 1L, "<< /Type /Catalog /Pages 2 0 R >>")
  buf <- r$buf; offsets[["1"]] <- r$offset

  r <- pdf_write_object(buf, 2L, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
  buf <- r$buf; offsets[["2"]] <- r$offset

  r <- pdf_write_object(buf, 3L, page_dict_body(2L, 6L, 4L))
  buf <- r$buf; offsets[["3"]] <- r$offset

  r <- pdf_write_object(buf, 4L, font_dict_body(5L))
  buf <- r$buf; offsets[["4"]] <- r$offset

  cmap_text <- make_bfrange_cmap(
    range1_start = 0x41L, range1_end = 0x43L, range1_base = utf8ToInt("A"),
    range2_start = 0x61L, range2_end = 0x62L,
    range2_unis  = c(utf8ToInt("z"), utf8ToInt("y"))
  )
  r <- pdf_write_stream_object(buf, 5L, "", charToRaw(cmap_text))
  buf <- r$buf; offsets[["5"]] <- r$offset

  content <- content_show_hex(72L, 720L, 12L, "4142436162")
  r <- pdf_write_stream_object(buf, 6L, "", charToRaw(content))
  buf <- r$buf; offsets[["6"]] <- r$offset

  list(
    bytes = finalise_simple(buf, offsets, root_obj = 1L),
    expect = list(n_pages = 1L, texts = list("ABCzy"))
  )
}

# Big-endian encode an integer to `width` bytes.
pdf_int_to_be <- function(x, width) {
  if (width == 0L) return(raw(0))
  out <- raw(width)
  for (i in seq.int(width, 1L)) {
    out[i] <- as.raw(bitwAnd(x, 0xFFL))
    x <- bitwShiftR(x, 8L)
  }
  out
}

# Generator 4: content stream compressed with FlateDecode.
make_flate_content <- function() {
  buf <- pdf_buf_new()
  r <- pdf_buf_add(buf, "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"); buf <- r$buf
  offsets <- list()

  r <- pdf_write_object(buf, 1L, "<< /Type /Catalog /Pages 2 0 R >>")
  buf <- r$buf; offsets[["1"]] <- r$offset

  r <- pdf_write_object(buf, 2L, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>")
  buf <- r$buf; offsets[["2"]] <- r$offset

  r <- pdf_write_object(buf, 3L, page_dict_body(2L, 6L, 4L))
  buf <- r$buf; offsets[["3"]] <- r$offset

  r <- pdf_write_object(buf, 4L, font_dict_body(5L))
  buf <- r$buf; offsets[["4"]] <- r$offset

  cmap_text <- make_bfchar_cmap(codes = c(0x41L, 0x42L),
                                unis  = c(utf8ToInt("H"), utf8ToInt("i")))
  r <- pdf_write_stream_object(buf, 5L, "", charToRaw(cmap_text))
  buf <- r$buf; offsets[["5"]] <- r$offset

  content <- content_show_hex(72L, 720L, 12L, "4142")
  compressed <- pdf_flate(content)
  r <- pdf_write_stream_object(buf, 6L, " /Filter /FlateDecode", compressed)
  buf <- r$buf; offsets[["6"]] <- r$offset

  list(
    bytes = finalise_simple(buf, offsets, root_obj = 1L),
    expect = list(n_pages = 1L, texts = list("Hi"))
  )
}

# Generator 5: a subset of dict objects packed into an /ObjStm compressed
# object stream. Exercises pdf_resolve_compressed_object().
#
# Layout:
#   1 = Catalog (uncompressed; must be reachable from the trailer without
#                needing an already-resolved /ObjStm)
#   2 = Pages   (goes inside ObjStm)
#   3 = Page    (goes inside ObjStm)
#   4 = Font    (uncompressed — has a /ToUnicode stream, can't pack streams)
#   5 = ToUnicode (stream, uncompressed)
#   6 = Content (stream, uncompressed)
#   7 = ObjStm containing objects 2 and 3
#
# Because we emit a traditional xref table (not a cross-reference stream),
# this file mixes `/Type /ObjStm` with a traditional xref — which is actually
# invalid per spec (compressed objects require an xref *stream*). So this
# generator is paired with Generator 6, which uses an xref stream. Generator 5
# is kept minimal and skipped in the core smoke test; it's here for manual
# exploration only.
#
# To keep the suite clean, we implement the compressed-object path inside
# Generator 6 below (XRefStm + ObjStm together, spec-compliant).

# Generator 6: cross-reference stream with compressed objects (/ObjStm).
# Exercises BOTH pdf_parse_xref_stream() and pdf_resolve_compressed_object().
make_xref_stream <- function() {
  # We emit:
  #   1 = Catalog         (in ObjStm)
  #   2 = Pages           (in ObjStm)
  #   3 = Page            (in ObjStm)
  #   4 = Font            (uncompressed — points to stream)
  #   5 = ToUnicode       (uncompressed stream)
  #   6 = Content         (uncompressed stream)
  #   7 = ObjStm          (uncompressed stream; contains objs 1-3)
  #   8 = XRefStm         (the xref itself, as a stream)
  #
  # PDF text content mirrors simple_one_page: "Hi".

  buf <- pdf_buf_new()
  r <- pdf_buf_add(buf, "%PDF-1.5\n%\xE2\xE3\xCF\xD3\n"); buf <- r$buf

  # Track offsets of uncompressed objects (for xref stream type-1 entries).
  # Compressed objects get type-2 entries pointing to the ObjStm.
  direct_offsets <- list()
  compressed_idx <- list()  # obj_num -> index within the ObjStm

  # Build ObjStm content. Object-stream format:
  #   header:  "obj1_num off1 obj2_num off2 ..."
  #   body:    concatenated object bodies (just the body — no "N G obj" wrapper)
  #
  # `/First` in the ObjStm dict is the byte offset where the body begins.
  packed_objs <- list(
    list(num = 1L, body = "<< /Type /Catalog /Pages 2 0 R >>"),
    list(num = 2L, body = "<< /Type /Pages /Kids [3 0 R] /Count 1 >>"),
    list(num = 3L, body = page_dict_body(2L, 6L, 4L))
  )

  # Compute offsets within the body
  body_offsets <- integer(length(packed_objs))
  cumulative <- 0L
  for (i in seq_along(packed_objs)) {
    body_offsets[i] <- cumulative
    cumulative <- cumulative + nchar(packed_objs[[i]]$body) + 1L  # +1 for newline
  }

  # Build header string
  header <- paste(vapply(seq_along(packed_objs), function(i) {
    sprintf("%d %d", packed_objs[[i]]$num, body_offsets[i])
  }, character(1)), collapse = " ")
  header_bytes <- charToRaw(paste0(header, "\n"))
  first_offset <- length(header_bytes)

  # Build body
  body_text <- paste(vapply(packed_objs, function(o) o$body, character(1)),
                     collapse = "\n")
  body_bytes <- charToRaw(paste0(body_text, "\n"))

  objstm_payload <- c(header_bytes, body_bytes)
  objstm_compressed <- memCompress(objstm_payload, type = "gzip")

  # Record compressed-object indices for xref stream entries
  for (i in seq_along(packed_objs)) {
    compressed_idx[[as.character(packed_objs[[i]]$num)]] <- i - 1L
  }

  # Emit object 4 (Font)
  r <- pdf_write_object(buf, 4L, font_dict_body(5L))
  buf <- r$buf; direct_offsets[["4"]] <- r$offset

  # Emit object 5 (ToUnicode)
  cmap_text <- make_bfchar_cmap(codes = c(0x41L, 0x42L),
                                unis  = c(utf8ToInt("H"), utf8ToInt("i")))
  r <- pdf_write_stream_object(buf, 5L, "", charToRaw(cmap_text))
  buf <- r$buf; direct_offsets[["5"]] <- r$offset

  # Emit object 6 (content stream)
  r <- pdf_write_stream_object(buf, 6L, "",
                               charToRaw(content_show_hex(72L, 720L, 12L, "4142")))
  buf <- r$buf; direct_offsets[["6"]] <- r$offset

  # Emit object 7 (ObjStm)
  objstm_dict <- sprintf(" /Type /ObjStm /N %d /First %d /Filter /FlateDecode",
                         length(packed_objs), first_offset)
  r <- pdf_write_stream_object(buf, 7L, objstm_dict, objstm_compressed)
  buf <- r$buf; direct_offsets[["7"]] <- r$offset

  # Build xref stream entries (object 0 + objects 1..8).
  # Entry widths: W = [1, 4, 2]  (type, field2, field3)
  W <- c(1L, 4L, 2L)
  size <- 9L  # includes object 0
  entries <- raw(0)

  # Object 0: type 0 (free), next-free = 0, gen = 65535
  entries <- c(entries,
               pdf_int_to_be(0L, W[1]),
               pdf_int_to_be(0L, W[2]),
               pdf_int_to_be(65535L, W[3]))

  for (n in seq_len(size - 1L)) {
    key <- as.character(n)
    if (!is.null(compressed_idx[[key]])) {
      # Type 2: compressed — field2 = stream_obj, field3 = index
      entries <- c(entries,
                   pdf_int_to_be(2L, W[1]),
                   pdf_int_to_be(7L, W[2]),
                   pdf_int_to_be(compressed_idx[[key]], W[3]))
    } else if (!is.null(direct_offsets[[key]])) {
      entries <- c(entries,
                   pdf_int_to_be(1L, W[1]),
                   pdf_int_to_be(direct_offsets[[key]] - 1L, W[2]),
                   pdf_int_to_be(0L, W[3]))
    } else if (n == 8L) {
      # Object 8 is the xref stream itself — fill in after we know its offset.
      # Placeholder for now; patched below.
      entries <- c(entries,
                   pdf_int_to_be(1L, W[1]),
                   pdf_int_to_be(0L, W[2]),         # patched
                   pdf_int_to_be(0L, W[3]))
    } else {
      entries <- c(entries,
                   pdf_int_to_be(0L, W[1]),
                   pdf_int_to_be(0L, W[2]),
                   pdf_int_to_be(0L, W[3]))
    }
  }

  xref_stream_offset <- buf$pos + 1L

  # Patch object 8's offset into the entry block.
  entry_size <- sum(W)
  obj8_entry_start <- 8L * entry_size + 1L  # 1-based
  patch <- pdf_int_to_be(xref_stream_offset - 1L, W[2])
  entries[(obj8_entry_start + W[1]):(obj8_entry_start + W[1] + W[2] - 1L)] <- patch

  xref_stream_body <- memCompress(entries, type = "gzip")
  xref_dict <- sprintf(
    " /Type /XRef /Size %d /W [%d %d %d] /Root 1 0 R /Filter /FlateDecode",
    size, W[1], W[2], W[3]
  )
  r <- pdf_write_stream_object(buf, 8L, xref_dict, xref_stream_body)
  buf <- r$buf

  # startxref + %%EOF — no trailer dict needed (xref stream dict serves as trailer)
  r <- pdf_buf_add(buf, sprintf("startxref\n%d\n%%%%EOF\n", xref_stream_offset - 1L))
  buf <- r$buf

  list(
    bytes = buf$bytes,
    expect = list(n_pages = 1L, texts = list("Hi"))
  )
}

# ---- 6. TOP-LEVEL ----

#' Generate all synthetic test PDFs into a directory.
#'
#' @param dir Output directory (created if missing).
#' @return Named list of expectations, one per generated file:
#'   list(<filename> = list(n_pages, texts))
#' @export
make_test_pdfs <- function(dir) {
  if (!dir.exists(dir)) dir.create(dir, recursive = TRUE)

  gens <- list(
    "simple_one_page.pdf" = make_simple_one_page,
    "two_pages.pdf"       = make_two_pages,
    "bfrange_cmap.pdf"    = make_bfrange_cmap_pdf,
    "flate_content.pdf"   = make_flate_content,
    "xref_stream.pdf"     = make_xref_stream
  )

  expectations <- list()
  for (name in names(gens)) {
    out <- gens[[name]]()
    writeBin(out$bytes, file.path(dir, name))
    expectations[[name]] <- out$expect
  }

  expectations
}
