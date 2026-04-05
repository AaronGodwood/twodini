`%||%` <- function(a, b) if (!is.null(a)) a else b

#' Return the Houdini Shiny application object
#'
#' Called internally by \code{\link{run_app}}. You can also pass the result
#' directly to \code{shiny::runApp()} if you need more control.
#'
#' @return A Shiny app object.
#' @import shiny
#' @import rhandsontable
#' @export
houdini_app <- function() {

ui <- fluidPage(

  tags$head(
    tags$style(HTML("
      .well { background-color: #f8f9fa; }
      .btn-primary { margin-top: 10px; }
      .section-header {
        margin-top: 20px;
        margin-bottom: 15px;
        padding-bottom: 5px;
        border-bottom: 2px solid #dee2e6;
      }
      .action-buttons { margin-top: 15px; }
      .status-box {
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
      }
      .status-success { background-color: #d4edda; border: 1px solid #c3e6cb; }
      .status-warning { background-color: #fff3cd; border: 1px solid #ffeeba; }
      .status-info    { background-color: #d1ecf1; border: 1px solid #bee5eb; }
      .preview-panel {
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 15px;
        min-height: 400px;
        max-height: 600px;
        overflow-y: auto;
        background-color: #fff;
      }
      .preview-placeholder {
        color: #999;
        text-align: center;
        padding-top: 50px;
      }
      .filter-panel {
        border: 1px solid #eee;
        border-radius: 4px;
        padding: 10px;
        margin-bottom: 10px;
        background: #fafafa;
      }
      .sel-pane th[data-col] { cursor: pointer; }
      .sel-pane tr[data-row] { cursor: pointer; }
      .preview-centred table { margin-left: auto; margin-right: auto; }
    "))
  ),

  titlePanel("Houdini V2"),

  fluidRow(
    # Left panel: inputs 
    column(3,
           wellPanel(
             h4("1. Word Document", class = "section-header"),
             fileInput("word_file", NULL,
                       accept = c(".docx", ".doc"),
                       placeholder = "Select Word file..."),
             uiOutput("bookmark_status"),

             h4("2. RTF Source", class = "section-header"),
             actionButton("pick_folder", "Choose RTF Folder...", class = "btn-primary"),
             uiOutput("rtf_folder_display"),
             uiOutput("rtf_status"),

             h4("3. Generate", class = "section-header"),
             downloadButton("download_result", "Download Result",
                            class = "btn-success btn-lg btn-block"),
             downloadButton("download_log", "Download Log",
                            class = "btn-outline-secondary btn-block",
                            style = "margin-top:6px;"),

             uiOutput("row_warning_display")
           )
    ),

    
    # Middle panel: config table 
    column(4,
           h4("Table Configuration", class = "section-header"),
           p("Map bookmarks to RTF tables. Click a row to configure filters and preview."),

           div(style = "height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 5px;",
               rHandsontableOutput("config_table")
           ),

           div(style = "display:flex;justify-content:space-between;align-items:flex-start;",
               div(class = "action-buttons",
                   actionButton("add_row",    "Add Row",    class = "btn-outline-secondary"),
                   actionButton("remove_row", "Remove Row", class = "btn-outline-secondary"),
                   actionButton("clear_all",  "Clear All",  class = "btn-outline-danger"),
                   fileInput("import_excel", NULL, accept = ".xlsx",
                             placeholder = "Import Excel...",
                             buttonLabel = "Import Excel",
                             width = "160px")
               ),
               downloadButton("export_excel", "Export Excel",
                              class = "btn-outline-secondary",
                              style = "margin-top:15px;")
           ),

           uiOutput("filter_panel")
    ),

    # Right panel: two-pane interactive preview 
    
    column(5,
           h5("Selection", class = "section-header",
              style = "margin-bottom:4px;font-size:0.95em;color:#555;"),
           p(style = "font-size:0.78em;color:#888;margin:0 0 6px;",
             "Click column headers to exclude columns. Click rows to exclude rows."),
           div(class = "sel-pane preview-centred",
               style = "border:1px solid #ddd;border-radius:5px;padding:8px;
                        max-height:320px;overflow:auto;background:#fff;",
               uiOutput("selection_preview")
           ),
           h5("Output", class = "section-header",
              style = "margin-top:14px;margin-bottom:4px;font-size:0.95em;color:#555;"),
           div(class = "preview-centred",
               style = "border:1px solid #ddd;border-radius:5px;padding:8px;
                        max-height:320px;overflow:auto;background:#fff;",
               uiOutput("output_preview")
           )
    )
  )
)

server <- function(input, output, session) {

 
  # REACTIVE VALUES

  

  available_bookmarks <- reactiveVal(character())
  available_tables    <- reactiveVal(character())
  rtf_paths           <- reactiveVal(list())
  rtf_folder_path     <- reactiveVal(NULL)

  
  config_data <- reactiveVal(
    data.frame(
      Bookmark  = character(),
      Table = character(),
      stringsAsFactors = FALSE
    )
  )

  current_table_name <- reactiveVal(NULL)
  current_row_index  <- reactiveVal(NULL)

  # table_info cache: table_name -> get_table_info() result
  table_info_cache <- reactiveVal(list())

  # parse_rtf cache: table_name -> parse_rtf() result (list of pages)
  parse_cache <- reactiveVal(list())

  # Get cached parsed pages for a table, parsing on first access
  get_cached_pages <- function(tbl_name) {
    cache <- parse_cache()
    if (!is.null(cache[[tbl_name]])) return(cache[[tbl_name]])
    paths <- rtf_paths()
    if (!tbl_name %in% names(paths)) return(NULL)
    pages <- parse_rtf(paths[[tbl_name]])
    cache[[tbl_name]] <- pages
    parse_cache(cache)
    pages
  }

  # selections keyed by row index (character):
  #   list(excluded_cols, excluded_rows, excluded_header_rows, parameters, timelines)
  table_selections    <- reactiveVal(list())
  last_gen_status     <- reactiveVal(NULL)

  #
  # WORD FILE

  

  observeEvent(input$word_file, {
    req(input$word_file)
    bookmarks <- tryCatch(
      extract_bookmarks(input$word_file$datapath),
      error = function(e) {
        showNotification(paste("Bookmark error:", conditionMessage(e)), type = "error")
        character()
      }
    )
    
    available_bookmarks(bookmarks)
  })

  
  output$bookmark_status <- renderUI({
    bm <- available_bookmarks()
    if (length(bm) > 0) {
      div(class = "status-box status-success",
          icon("check-circle"), paste(length(bm), "bookmarks found"))
    } else if (!is.null(input$word_file)) {
      div(class = "status-box status-warning",
          icon("exclamation-triangle"), "No bookmarks found")
    }
  })

  # RTF FOLDER


  observeEvent(input$pick_folder, {
    folder <- tryCatch(
      rstudioapi::selectDirectory(caption = "Select RTF folder"),
      error = function(e) NULL
    )
    if (is.null(folder) || !nzchar(folder)) return()

    rtf_folder_path(folder)

    rtf_files <- list.files(folder, pattern = "\\.rtf$", ignore.case = TRUE)
    if (length(rtf_files) == 0L) {
      available_tables(character())
      showNotification("No RTF files found in folder", type = "warning")
      return()
    }
    tbl_names  <- tools::file_path_sans_ext(rtf_files)
    full_paths <- file.path(folder, rtf_files)
    available_tables(tbl_names)
    rtf_paths(setNames(as.list(full_paths), tbl_names))
    table_info_cache(list())
    parse_cache(list())
    showNotification(paste("Found", length(rtf_files), "RTF files"), type = "message")
  })

  output$rtf_folder_display <- renderUI({
    folder <- rtf_folder_path()
    if (!is.null(folder)) {
      p(style = "font-size:0.8em;color:#555;margin:4px 0 0;word-break:break-all;", folder)
    }
  })

  output$rtf_status <- renderUI({
    tbls <- available_tables()
    if (length(tbls) > 0) {
      div(class = "status-box status-success",
          icon("check-circle"), paste(length(tbls), "RTF files found"))
    }
  })


  
  # CONFIG TABLE

  

  output$config_table <- renderRHandsontable({
    df <- config_data()
    if (nrow(df) == 0) {
      df <- data.frame(Bookmark = rep("", 5), Table = rep("", 5),
                       stringsAsFactors = FALSE)
    }

    bm  <- available_bookmarks()
    tbl <- available_tables()

    hot <- rhandsontable(df, rowHeaders = TRUE, selectCallback = TRUE,
                         overflow = "visible") |>
      hot_cols(colWidths = c(160, 160))

    hot <- if (length(bm) > 0) {
      hot |> hot_col("Bookmark", type = "dropdown", source = c("", names(bm)), strict = FALSE)
    } else {
      hot |> hot_col("Bookmark", type = "text")
    }

    hot <- if (length(tbl) > 0) {
      hot |> hot_col("Table", type = "dropdown", source = c("", tbl), strict = FALSE)
    } else {
      hot |> hot_col("Table", type = "text")
    }

    hot
  })

  observeEvent(input$config_table, {
    if (!is.null(input$config_table)) config_data(hot_to_r(input$config_table))
  })

  observeEvent(input$add_row, {
    df <- config_data()
    config_data(rbind(df, data.frame(Bookmark = "", Table = "",
                                     stringsAsFactors = FALSE)))
  })

  observeEvent(input$remove_row, {
    df <- config_data()
    if (nrow(df) > 1) config_data(df[-nrow(df), , drop = FALSE])
  })

  observeEvent(input$clear_all, {
    config_data(data.frame(Bookmark = rep("", 5), Table = rep("", 5),
                           stringsAsFactors = FALSE))
    table_selections(list())
  })


  # EXCEL IMPORT

  observeEvent(input$import_excel, {
    req(input$import_excel)
    path <- input$import_excel$datapath

    xl <- tryCatch(
      readxl::read_excel(path, col_types = "text"),
      error = function(e) {
        showNotification(paste("Could not read Excel file:", conditionMessage(e)), type = "error")
        NULL
      }
    )
    if (is.null(xl)) return()

    # Must have Bookmark and Table columns (case-insensitive)
    col_lower <- tolower(names(xl))
    bm_col  <- which(col_lower == "bookmark")[1L]
    tbl_col <- which(col_lower == "table")[1L]

    if (is.na(bm_col) || is.na(tbl_col)) {
      showNotification(
        "Excel file must contain 'Bookmark' and 'Table' columns", type = "error"
      )
      return()
    }

    # Build config_data frame - strip any .rtf extension from Tables
    
    new_config <- data.frame(
      Bookmark  = as.character(xl[[bm_col]]),
      Table = tools::file_path_sans_ext(as.character(xl[[tbl_col]])),
      stringsAsFactors = FALSE
    )
    # Replace NA with empty string
    new_config$Bookmark[is.na(new_config$Bookmark)]   <- ""
    new_config$Table[is.na(new_config$Table)] <- ""

    config_data(new_config)

    # Parse optional filter columns into table_selections
    # Recognised column names (case-insensitive): parameters, timelines
    param_col  <- which(col_lower == "parameters")[1L]
    tline_col  <- which(col_lower == "timelines")[1L]

    split_semi <- function(x) {
      if (is.na(x) || !nzchar(trimws(x))) return(character())
      trimws(strsplit(x, ";", fixed = TRUE)[[1L]])
    }

    # Fresh selections keyed by row index - discard any previous state
    sels <- list()

    for (i in seq_len(nrow(new_config))) {
      tname <- new_config$Table[i]
      if (!nzchar(tname)) next

      params <- if (!is.na(param_col)) split_semi(xl[[param_col]][i]) else character()
      tlines <- if (!is.na(tline_col)) split_semi(xl[[tline_col]][i]) else character()

      sels[[as.character(i)]] <- list(
        excluded_cols        = NULL,
        excluded_rows        = NULL,
        excluded_header_rows = NULL,
        parameters           = if (length(params) > 0L) params else NULL,
        timelines            = if (length(tlines) > 0L) tlines else NULL
      )
    }

    table_selections(sels)
    showNotification(
      paste("Imported", nrow(new_config), "rows from Excel"), type = "message"
    )
  })

  # PER-ROW VALIDATION

  # Returns a list of length nrow(config_data()).
  # Each element: list(type = "error"|"warning"|"", msg = "plain text message")
  row_warnings <- reactive({
    df    <- config_data()
    bm    <- available_bookmarks()
    tbls  <- available_tables()
    paths <- rtf_paths()
    cache <- table_info_cache()
    sels  <- table_selections()

    lapply(seq_len(nrow(df)), function(i) {
      bm_val  <- trimws(df$Bookmark[i])
      tbl_val <- trimws(df$Table[i])

      if (!nzchar(bm_val) && !nzchar(tbl_val)) return(list(type = "", msg = ""))

      errors   <- character()
      warnings <- character()

      if (nzchar(bm_val) && length(bm) > 0L && !bm_val %in% names(bm))
        errors <- c(errors, paste0("Bookmark '", bm_val, "' not found in Word document"))

      if (nzchar(tbl_val) && length(tbls) > 0L && !tbl_val %in% tbls)
        errors <- c(errors, paste0("Table '", tbl_val, "' not found in RTF folder"))

      if (nzchar(tbl_val) && length(paths) > 0L && tbl_val %in% names(paths)) {
        info <- cache[[tbl_val]]
        sel  <- sels[[as.character(i)]]

        if (!is.null(info) && !isTRUE(info$is_image)) {
          # Only validate against a non-empty known list - if the table has no
          # parameters/timelines detected, we can't meaningfully validate imports
          if (!is.null(sel$parameters) && length(sel$parameters) > 0L &&
              length(info$parameters) > 0L) {
            bad_p <- setdiff(sel$parameters, info$parameters)
            if (length(bad_p) > 0L)
              warnings <- c(warnings,
                paste0("Unknown parameter(s): ", paste(bad_p, collapse = ", ")))
          }
          if (!is.null(sel$timelines) && length(sel$timelines) > 0L &&
              length(info$timelines) > 0L) {
            bad_t <- setdiff(sel$timelines, info$timelines)
            if (length(bad_t) > 0L)
              warnings <- c(warnings,
                paste0("Unknown timeline(s): ", paste(bad_t, collapse = ", ")))
          }
          if (!is.null(sel$excluded_cols) && length(sel$excluded_cols) > 0L &&
              info$n_cols > 0L) {
            bad_c <- sel$excluded_cols[
              sel$excluded_cols < 1L | sel$excluded_cols > info$n_cols
            ]
            if (length(bad_c) > 0L)
              warnings <- c(warnings,
                paste0("Excluded column index(es) out of range: ",
                       paste(bad_c, collapse = ", ")))
          }
        }
      }

      if (length(errors) > 0L)
        return(list(type = "error",   msg = paste(errors,   collapse = "; ")))
      if (length(warnings) > 0L)
        return(list(type = "warning", msg = paste(warnings, collapse = "; ")))

      list(type = "", msg = "")
    })
  })

  # Scrollable warnings panel - all rows with issues, shown below the preview
  output$row_warning_display <- renderUI({
    warns <- row_warnings()
    df    <- config_data()

    items <- lapply(seq_along(warns), function(i) {
      w <- warns[[i]]
      if (!nzchar(w$type)) return(NULL)

      is_error <- w$type == "error"
      bg  <- if (is_error) "#f8d7da" else "#fff3cd"
      bdr <- if (is_error) "#f5c2c7" else "#ffecb5"
      ico <- if (is_error) "\u26a0" else "\u26a0"
      lbl <- if (is_error) "Error" else "Warning"

      bm_val  <- trimws(df$Bookmark[i])
      tbl_val <- trimws(df$Table[i])
      row_lbl <- paste0("Row ", i,
                        if (nzchar(bm_val))  paste0(" \u2013 ", bm_val)  else "",
                        if (nzchar(tbl_val)) paste0(" / ", tbl_val) else "")

      div(style = sprintf(
            "padding:5px 8px;margin-bottom:4px;border-radius:4px;
             background:%s;border:1px solid %s;font-size:0.82em;", bg, bdr),
          strong(paste0(ico, " ", lbl, ": ")), row_lbl,
          tags$br(),
          span(style = "color:#555;", w$msg))
    })

    items <- Filter(Negate(is.null), items)
    if (length(items) == 0L) return(NULL)

    div(style = "margin-top:12px;max-height:350px;overflow-y:auto;",
        h5(style = "margin:0 0 6px;font-size:0.9em;color:#666;", "Validation warnings"),
        items)
  })

  # ROW SELECTION -> load table info

  observeEvent(input$config_table_select, {
    sel <- input$config_table_select
    if (is.null(sel)) return()

    df  <- config_data()
    row <- sel$select$r

    if (row > 0 && row <= nrow(df)) {
      tbl_name <- df$Table[row]
      paths    <- rtf_paths()

      if (!is.null(tbl_name) && tbl_name != "" && tbl_name %in% names(paths)) {
        current_table_name(tbl_name)
        current_row_index(row)

        # Load table info if not cached (skip for image RTFs - no table to inspect)
        cache <- table_info_cache()
        if (is.null(cache[[tbl_name]])) {
          if (isTRUE(is_image_rtf(paths[[tbl_name]]))) {
            # Sentinel so the cache entry exists and current_info() is non-NULL
            cache[[tbl_name]] <- list(n_cols = 0L, n_rows = 0L,
                                      col_names = character(),
                                      parameters = character(),
                                      timelines = character(),
                                      is_image = TRUE)
            table_info_cache(cache)
          } else {
            info <- tryCatch({
              pages <- get_cached_pages(tbl_name)
              combined <- combine_pages(pages)
              ref_row <- if (length(combined$header_rows) > 0L) {
                combined$header_rows[[1]]
              } else if (length(combined$data_rows) > 0L) {
                combined$data_rows[[1]]
              } else NULL
              n_cols <- if (!is.null(ref_row)) length(ref_row$cells) else 0L
              col_names <- if (!is.null(ref_row)) {
                vapply(ref_row$cells, function(cell) cell$text, character(1))
              } else character()
              list(
                n_cols     = n_cols,
                n_rows     = length(combined$data_rows),
                col_names  = col_names,
                parameters = get_parameters(pages),
                timelines  = get_timelines(pages)
              )
            }, error = function(e) {
              showNotification(paste("Error reading RTF:", conditionMessage(e)), type = "error")
              NULL
            })
            if (!is.null(info)) {
              cache[[tbl_name]] <- info
              table_info_cache(cache)
            }
          }
        }

      } else {
        current_table_name(NULL)
        current_row_index(NULL)
      }
    }
  })

  # Convenience reactive: current table info
  current_info <- reactive({
    tbl_name <- current_table_name()
    if (is.null(tbl_name)) return(NULL)
    table_info_cache()[[tbl_name]]
  })

  # FILTER PANEL

  output$filter_panel <- renderUI({
    info <- current_info()
    if (is.null(info) || isTRUE(info$is_image)) return(NULL)
    if (length(info$parameters) == 0L && length(info$timelines) == 0L) return(NULL)

    prev        <- table_selections()[[as.character(current_row_index())]]
    prev_params <- prev$parameters %||% info$parameters
    prev_tlines <- prev$timelines  %||% info$timelines

    div(class = "filter-panel",
        h5("Filters"),

        if (length(info$parameters) > 0L) {
          selectizeInput(
            "sel_parameters", "Parameters:",
            choices  = info$parameters,
            selected = prev_params,
            multiple = TRUE,
            options  = list(plugins = list("remove_button"),
                            placeholder = "All parameters")
          )
        },

        if (length(info$timelines) > 0L) {
          selectizeInput(
            "sel_timelines", "Timelines:",
            choices  = info$timelines,
            selected = prev_tlines,
            multiple = TRUE,
            options  = list(plugins = list("remove_button"),
                            placeholder = "All timelines")
          )
        }
    )
  })

  # SAVE SELECTIONS WHEN INPUTS CHANGE

  # Save parameters/timelines when dropdowns change
  observeEvent(
    list(input$sel_parameters, input$sel_timelines),
    {
      row  <- isolate(current_row_index())
      info <- isolate(current_info())
      if (is.null(row) || is.null(info) || isTRUE(info$is_image)) return()

      sels <- isolate(table_selections())
      prev <- sels[[as.character(row)]] %||% list()
      sels[[as.character(row)]] <- list(
        excluded_cols        = prev$excluded_cols,
        excluded_rows        = prev$excluded_rows,
        excluded_header_rows = prev$excluded_header_rows,
        parameters           = input$sel_parameters %||% info$parameters,
        timelines            = input$sel_timelines  %||% info$timelines
      )
      table_selections(sels)
    },
    ignoreNULL = FALSE
  )

  # Save exclusions when JS sends updated sets from the selection pane
  observeEvent(
    list(input$preview_excluded_cols,
         input$preview_excluded_rows,
         input$preview_excluded_header_rows),
    {
      row <- isolate(current_row_index())
      if (is.null(row)) return()

      sels <- isolate(table_selections())
      prev <- sels[[as.character(row)]] %||% list()
      sels[[as.character(row)]] <- list(
        excluded_cols        = as.integer(input$preview_excluded_cols        %||% integer()),
        excluded_rows        = as.integer(input$preview_excluded_rows        %||% integer()),
        excluded_header_rows = as.integer(input$preview_excluded_header_rows %||% integer()),
        parameters           = prev$parameters,
        timelines            = prev$timelines
      )
      table_selections(sels)
    },
    ignoreNULL = FALSE
  )

  # INTERACTIVE PREVIEW - selection pane + output pane

  # JS for the selection pane (injected after each render)
  selection_js <- '
(function() {
  var excCols = [], excRows = [], excHdrs = [];

  function toggle(arr, val) {
    var i = arr.indexOf(val);
    if (i === -1) arr.push(val); else arr.splice(i, 1);
  }

  function applyStyles(tbl) {
    tbl.querySelectorAll("tr[data-row]").forEach(function(tr) {
      var ri   = parseInt(tr.dataset.row, 10);
      var type = tr.dataset.rowtype;
      var excl = type === "header" ? excHdrs.indexOf(ri) !== -1
                                   : excRows.indexOf(ri) !== -1;
      tr.querySelectorAll("td,th").forEach(function(cell) {
        var raw = cell.dataset.col;
        var cols;
        try { cols = JSON.parse(raw); if (!Array.isArray(cols)) cols = [cols]; }
        catch(e) { cols = [parseInt(raw, 10)]; }
        var cExcl = cols.some(function(c) { return excCols.indexOf(c) !== -1; });
        cell.style.opacity = (excl || cExcl) ? "0.3" : "";
      });
      tr.style.opacity = excl ? "0.3" : "";
    });
  }

  var timer;
  function sendUpdate(tbl) {
    clearTimeout(timer);
    timer = setTimeout(function() {
      Shiny.setInputValue("preview_excluded_cols",
        excCols.slice().sort(function(a,b){return a-b;}), {priority:"event"});
      Shiny.setInputValue("preview_excluded_rows",
        excRows.slice().sort(function(a,b){return a-b;}), {priority:"event"});
      Shiny.setInputValue("preview_excluded_header_rows",
        excHdrs.slice().sort(function(a,b){return a-b;}), {priority:"event"});
    }, 300);
  }

  var container = document.getElementById("sel-pane-container");
  if (!container) return;
  try { excCols = JSON.parse(container.dataset.excCols || "[]"); } catch(e) {}
  try { excRows = JSON.parse(container.dataset.excRows || "[]"); } catch(e) {}
  try { excHdrs = JSON.parse(container.dataset.excHdrs || "[]"); } catch(e) {}

  var tbl = container.querySelector("table");
  if (!tbl) return;
  applyStyles(tbl);

  tbl.addEventListener("click", function(e) {
    var th = e.target.closest("th[data-col]");
    var tr = e.target.closest("tr[data-row]");
    if (th) {
      var raw = th.dataset.col;
      var cols;
      try { cols = JSON.parse(raw); if (!Array.isArray(cols)) cols = [cols]; }
      catch(err) { cols = [parseInt(raw, 10)]; }
      cols.forEach(function(c) { toggle(excCols, c); });
      applyStyles(tbl);
      sendUpdate(tbl);
      e.stopPropagation();
    } else if (tr) {
      var ri   = parseInt(tr.dataset.row, 10);
      var type = tr.dataset.rowtype;
      if (type === "header") toggle(excHdrs, ri); else toggle(excRows, ri);
      applyStyles(tbl);
      sendUpdate(tbl);
    }
  });
})();
'

  # Reactive that captures the "identity" of the current table+filters
  # (drives selection pane re-render - not per-click)
  selection_pane_key <- reactive({
    row <- current_row_index()
    sel <- if (!is.null(row)) table_selections()[[as.character(row)]] else NULL
    list(
      table  = current_table_name(),
      params = sel$parameters,
      tlines = sel$timelines
    )
  })

  output$selection_preview <- renderUI({
    key      <- selection_pane_key()
    tbl_name <- key$table
    paths    <- rtf_paths()
    info     <- current_info()

    if (is.null(tbl_name) || !tbl_name %in% names(paths)) {
      return(p(style = "color:#999;padding:20px;text-align:center;",
               "Select a row to preview"))
    }

    if (isTRUE(info$is_image)) {
      img <- tryCatch(extract_png(paths[[tbl_name]]), error = function(e) NULL)
      if (is.null(img))
        return(HTML("<p style='color:red'>Could not extract image.</p>"))
      b64 <- paste0("data:image/png;base64,",
                    base64enc::base64encode(img$png_bytes))
      w_px <- if (!is.na(img$width_twips))  round(img$width_twips  * 96 / 1440) else NULL
      h_px <- if (!is.na(img$height_twips)) round(img$height_twips * 96 / 1440) else NULL
      img_style <- paste0(
        "max-width:100%;height:auto;",
        if (!is.null(w_px)) paste0("width:", w_px, "px;") else "",
        if (!is.null(h_px)) paste0("height:", h_px, "px;") else ""
      )
      return(tags$div(style = "text-align:center;",
                      tags$img(src = b64, style = img_style)))
    }

    row <- current_row_index()
    sel <- table_selections()[[as.character(row)]] %||% list()
    ec  <- sel$excluded_cols        %||% integer()
    er  <- sel$excluded_rows        %||% integer()
    eh  <- sel$excluded_header_rows %||% integer()

    int_to_json <- function(x) {
      if (length(x) == 0L) return("[]")
      paste0("[", paste(x, collapse = ","), "]")
    }

    cached_pages <- get_cached_pages(tbl_name)

    html_content <- tryCatch(
      get_table_html_selection(
        paths[[tbl_name]],
        excluded_cols        = ec,
        excluded_rows        = er,
        excluded_header_rows = eh,
        parameters           = sel$parameters,
        timelines            = sel$timelines,
        pages                = cached_pages
      ),
      error = function(e) sprintf(
        "<p style='color:red'>Selection error: %s</p>",
        htmlEscape(conditionMessage(e))
      )
    )

    tagList(
      div(
        id = "sel-pane-container",
        `data-exc-cols` = int_to_json(ec),
        `data-exc-rows` = int_to_json(er),
        `data-exc-hdrs` = int_to_json(eh),
        HTML(html_content)
      ),
      tags$script(HTML(selection_js))
    )
  })

  # Debounced exclusions for the output pane
  raw_exclusions <- reactive({
    list(
      cols  = as.integer(input$preview_excluded_cols         %||% integer()),
      rows  = as.integer(input$preview_excluded_rows         %||% integer()),
      hdrs  = as.integer(input$preview_excluded_header_rows  %||% integer())
    )
  })
  debounced_exclusions <- debounce(raw_exclusions, 300)

  output$output_preview <- renderUI({
    tbl_name <- current_table_name()
    paths    <- rtf_paths()
    info     <- current_info()

    if (is.null(tbl_name) || !tbl_name %in% names(paths)) {
      return(p(style = "color:#999;padding:20px;text-align:center;",
               "No table selected"))
    }

    if (isTRUE(info$is_image)) {
      img <- tryCatch(extract_png(paths[[tbl_name]]), error = function(e) NULL)
      if (is.null(img))
        return(HTML("<p style='color:red'>Could not extract image.</p>"))
      b64 <- paste0("data:image/png;base64,",
                    base64enc::base64encode(img$png_bytes))
      w_px <- if (!is.na(img$width_twips))  round(img$width_twips  * 96 / 1440) else NULL
      h_px <- if (!is.na(img$height_twips)) round(img$height_twips * 96 / 1440) else NULL
      img_style <- paste0(
        "max-width:100%;height:auto;",
        if (!is.null(w_px)) paste0("width:", w_px, "px;") else "",
        if (!is.null(h_px)) paste0("height:", h_px, "px;") else ""
      )
      return(tags$div(style = "text-align:center;",
                      tags$img(src = b64, style = img_style)))
    }

    excl <- debounced_exclusions()
    row  <- current_row_index()
    sel  <- table_selections()[[as.character(row)]] %||% list()

    cached_pages <- get_cached_pages(tbl_name)

    html_content <- tryCatch(
      get_table_html_output(
        paths[[tbl_name]],
        excluded_cols        = excl$cols,
        excluded_rows        = excl$rows,
        excluded_header_rows = excl$hdrs,
        parameters           = sel$parameters,
        timelines            = sel$timelines,
        pages                = cached_pages
      ),
      error = function(e) sprintf(
        "<p style='color:red'>Output error: %s</p>",
        htmlEscape(conditionMessage(e))
      )
    )

    HTML(html_content)
  })

  # LOG DOWNLOAD

  output$download_log <- downloadHandler(
    filename = function() {
      paste0("houdini_log_", format(Sys.Date(), "%Y%m%d"), ".txt")
    },
    content = function(file) {
      df    <- config_data()
      sels  <- table_selections()
      paths <- rtf_paths()

      lines <- character()

      lines <- c(lines,
        "  Houdini Document Generation Log",
        paste0("Generated : ", format(Sys.time(), "%Y-%m-%d %H:%M:%S")),
        paste0("User      : ", Sys.info()[["user"]]),
        paste0("Word file : ", if (!is.null(input$word_file)) input$word_file$name else "(not set)"),
        paste0("RTF folder: ", rtf_folder_path() %||% "(not set)"),
        ""
      )

      valid_rows <- which(nzchar(trimws(df$Bookmark)) & nzchar(trimws(df$Table)))

      if (length(valid_rows) == 0L) {
        lines <- c(lines, "(no table mappings defined)")
      } else {
        gen_status <- last_gen_status()

      fmt_vec <- function(x, none = "(all)") {
        if (is.null(x) || length(x) == 0L) none else paste(x, collapse = "; ")
      }

      for (i in valid_rows) {
          bm_val  <- trimws(df$Bookmark[i])
          tbl_val <- trimws(df$Table[i])
          sel     <- sels[[as.character(i)]] %||% list()
          err     <- gen_status[[as.character(i)]]

          if (!is.null(err)) {
            lines <- c(lines,
              paste0("Row       : ", i),
              paste0("Bookmark  : ", bm_val),
              paste0("Table     : ", tbl_val, ".rtf"),
              paste0("Status    : ERROR - ", err),
              ""
            )
          } else {
            lines <- c(lines,
              paste0("Row       : ", i),
              paste0("Bookmark  : ", bm_val),
              paste0("Table     : ", tbl_val, ".rtf"),
              paste0("Parameters: ", fmt_vec(sel$parameters)),
              paste0("Timelines : ", fmt_vec(sel$timelines)),
              ""
            )
          }
        }
      }

      writeLines(lines, file)
    }
  )

  # EXCEL EXPORT

  output$export_excel <- downloadHandler(
    filename = function() {
      paste0("houdini_config_", format(Sys.Date(), "%Y%m%d"), ".xlsx")
    },
    content = function(file) {
      df   <- config_data()
      sels <- table_selections()

      semi_join <- function(x) if (length(x) == 0L || is.null(x)) "" else paste(x, collapse = "; ")

      out <- data.frame(
        Bookmark   = df$Bookmark,
        Table      = ifelse(nzchar(df$Table), paste0(df$Table, ".rtf"), df$Table),
        Parameters = vapply(seq_len(nrow(df)), function(i) {
          semi_join(sels[[as.character(i)]]$parameters)
        }, character(1)),
        Timelines  = vapply(seq_len(nrow(df)), function(i) {
          semi_join(sels[[as.character(i)]]$timelines)
        }, character(1)),
        stringsAsFactors = FALSE
      )

      writexl::write_xlsx(out, file)
    }
  )

  # DOCUMENT GENERATION

  output$download_result <- downloadHandler(
    filename = function() {
      if (!is.null(input$word_file)) paste0("modified_", input$word_file$name)
      else "output.docx"
    },
    content = function(file) {
      req(input$word_file)

      config <- config_data()
      config <- config[config$Bookmark != "" & config$Table != "", ]

      if (nrow(config) == 0L) {
        showNotification("No table mappings defined", type = "error"); return()
      }

      status <- tryCatch(
        process_document(
          word_path   = input$word_file$datapath,
          config      = config,
          rtf_paths   = rtf_paths(),
          selections  = table_selections(),
          output_path = file
        ),
        error = function(e) {
          showNotification(paste("Error generating document:", conditionMessage(e)), type = "error")
          NULL
        }
      )
      last_gen_status(status)
    }
  )
}

shiny::shinyApp(ui, server)
} # end houdini_app()
