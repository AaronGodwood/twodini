#' @useDynLib houdini, .registration = TRUE
#' @importFrom stats setNames
#' @importFrom utils unzip
#' @importFrom rstudioapi selectDirectory isAvailable
#' @importFrom xml2 read_xml write_xml xml_find_all xml_find_first xml_attr xml_name xml_parent xml_add_child xml_add_sibling xml_root
"_PACKAGE"

# Null-coalescing helper shared across the package.
`%||%` <- function(a, b) if (!is.null(a)) a else b
