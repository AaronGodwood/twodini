#' Launch the Houdini Shiny application
#'
#' Opens the interactive Word document generator in your default browser.
#'
#' @param ... Arguments passed to \code{\link[shiny]{runApp}}, e.g.
#'   \code{port}, \code{launch.browser}.
#' @return Called for its side effect of launching the app.
#' @export
run_app <- function(...) {
  shiny::runApp(houdini_app(), ...)
}

