# errors.R — Structured error taxonomy for houdini
#
# Every public failure path should stop() with one of these constructors.
# All errors share the class hierarchy:
#   c("houdini_error_<kind>", "houdini_error", "error", "condition")
#
# Fields common to all:
#   message  — human-readable summary (used by conditionMessage())
#   hint     — actionable advice shown in the UI
#   call     — always NULL (we don't expose internal call stacks to users)
#
# Additional fields are documented per constructor.

# 
# Internal constructor
# 

houdini_error <- function(kind, message, hint, fields = list()) {
  structure(
    c(
      list(message = message, hint = hint, call = NULL),
      fields
    ),
    class = c(
      paste0("houdini_error_", kind),
      "houdini_error",
      "error",
      "condition"
    )
  )
}

# 
# Bookmark errors
# 

#' Bookmark not found in the Word document
#' @param name Bookmark name
err_bookmark_missing <- function(name) {
  houdini_error(
    "bookmark_missing",
    sprintf("Bookmark '%s' not found in the Word document.", name),
    "Check that the bookmark name in the config table matches exactly (case-sensitive) the bookmark in the Word file.",
    list(bookmark = name)
  )
}

#' Bookmark is inside a table cell, header, footer, or text box
#' @param name Bookmark name
#' @param context One of: "table", "header", "footer", "textbox"
err_bookmark_bad_context <- function(name, context) {
  location <- switch(context,
    table   = "inside a table cell",
    header  = "in a page header",
    footer  = "in a page footer",
    textbox = "inside a text box",
    sprintf("in an unsupported location (%s)", context)
  )
  houdini_error(
    "bookmark_bad_context",
    sprintf("Bookmark '%s' is %s. Injecting a table or image here will produce invalid output.", name, location),
    "Move the bookmark to the document body (not inside any table, header, footer, or text box).",
    list(bookmark = name, context = context)
  )
}

#' Duplicate bookmark mapping across config rows
#' @param name Bookmark name
#' @param rows Integer vector of 1-based row indices that reference this bookmark
err_bookmark_duplicate <- function(name, rows) {
  houdini_error(
    "bookmark_duplicate",
    sprintf("Bookmark '%s' is mapped in multiple rows: %s.", name, paste(rows, collapse = ", ")),
    "Each bookmark should appear in at most one config row. Remove or rename the duplicate entries.",
    list(bookmark = name, rows = rows)
  )
}

# 
# RTF / file errors
# 

#' RTF file cannot be read
#' @param path File path
#' @param cause Underlying condition or message string
err_rtf_unreadable <- function(path, cause = NULL) {
  cause_msg <- if (!is.null(cause)) {
    if (inherits(cause, "condition")) conditionMessage(cause) else as.character(cause)
  } else ""
  msg <- sprintf("Cannot read RTF file '%s'.", basename(path))
  if (nzchar(cause_msg)) msg <- paste0(msg, " ", cause_msg)
  houdini_error(
    "rtf_unreadable",
    msg,
    "Check that the file exists, is not open in another application, and is a valid RTF file.",
    list(path = path)
  )
}

#' RTF parsing failed at a specific stage
#' @param path File path
#' @param stage One of: "split_pages", "parse_page", "parse_row", "extract_image"
#' @param cause Underlying condition or message string
err_rtf_parse_failed <- function(path, stage, cause = NULL) {
  cause_msg <- if (!is.null(cause)) {
    if (inherits(cause, "condition")) conditionMessage(cause) else as.character(cause)
  } else ""
  msg <- sprintf("Failed to parse RTF file '%s' (stage: %s).", basename(path), stage)
  if (nzchar(cause_msg)) msg <- paste0(msg, " ", cause_msg)
  houdini_error(
    "rtf_parse_failed",
    msg,
    "The RTF file may be malformed or use unsupported SAS ODS features. Try opening it in a text editor to inspect the structure.",
    list(path = path, stage = stage)
  )
}

#' PNG image could not be extracted from an RTF file
#' @param path File path
#' @param cause Underlying condition or message string
err_image_extract_failed <- function(path, cause = NULL) {
  cause_msg <- if (!is.null(cause)) {
    if (inherits(cause, "condition")) conditionMessage(cause) else as.character(cause)
  } else ""
  msg <- sprintf("Could not extract image from RTF file '%s'.", basename(path))
  if (nzchar(cause_msg)) msg <- paste0(msg, " ", cause_msg)
  houdini_error(
    "image_extract_failed",
    msg,
    "Check that the RTF contains a valid embedded PNG image (\\pngblip). WMF and EMF formats are not supported.",
    list(path = path)
  )
}

# 
# Injection errors
# 

#' XML table injection failed
#' @param bookmark Bookmark name
#' @param cause Underlying condition or message string
err_xml_inject_failed <- function(bookmark, cause = NULL) {
  cause_msg <- if (!is.null(cause)) {
    if (inherits(cause, "condition")) conditionMessage(cause) else as.character(cause)
  } else ""
  msg <- sprintf("Failed to inject table at bookmark '%s'.", bookmark)
  if (nzchar(cause_msg)) msg <- paste0(msg, " ", cause_msg)
  houdini_error(
    "xml_inject_failed",
    msg,
    "The generated Word XML may be malformed. Check that the RTF table structure is valid and re-run validation.",
    list(bookmark = bookmark)
  )
}

#' Image injection failed
#' @param bookmark Bookmark name
#' @param cause Underlying condition or message string
err_image_inject_failed <- function(bookmark, cause = NULL) {
  cause_msg <- if (!is.null(cause)) {
    if (inherits(cause, "condition")) conditionMessage(cause) else as.character(cause)
  } else ""
  msg <- sprintf("Failed to inject image at bookmark '%s'.", bookmark)
  if (nzchar(cause_msg)) msg <- paste0(msg, " ", cause_msg)
  houdini_error(
    "image_inject_failed",
    msg,
    "Check that the Word document is not password-protected and that the bookmark context is valid.",
    list(bookmark = bookmark)
  )
}

# 
# Config / selection errors
# 

#' Excluded column index is out of range
#' @param bookmark Bookmark name (for context)
#' @param indices Integer vector of out-of-range indices
#' @param max_col Maximum valid column index
err_exclusion_out_of_range <- function(bookmark, indices, max_col) {
  houdini_error(
    "exclusion_out_of_range",
    sprintf(
      "Excluded column index %s is out of range for bookmark '%s' (table has %d column%s).",
      paste(indices, collapse = ", "), bookmark, max_col,
      if (max_col == 1L) "" else "s"
    ),
    "Re-open the table preview and re-apply column exclusions. The table structure may have changed since the config was last saved.",
    list(bookmark = bookmark, indices = indices, max_col = max_col)
  )
}

# 
# Utility: extract hint from any condition
# 

#' Return the hint field if the condition is a houdini_error, otherwise NULL
#' @param e A condition object
err_hint <- function(e) {
  if (inherits(e, "houdini_error")) e$hint else NULL
}

#' Format an error for display in the UI (message + hint on a new line if present)
#' @param e A condition object
err_format <- function(e) {
  msg  <- conditionMessage(e)
  hint <- err_hint(e)
  if (!is.null(hint)) paste0(msg, "\nHint: ", hint) else msg
}
