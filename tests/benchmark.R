source("R/parsing.R")
source("R/html.R")
source("R/xml.R")

files <- list.files("../test_data/rtf", full.names = TRUE, pattern = "\\.rtf$")
table_files <- files[!vapply(files, is_image_rtf, logical(1))]
cat(sprintf("Benchmarking %d table RTFs\n\n", length(table_files)))

# Read all files into memory first so we're not measuring disk I/O
texts <- lapply(table_files, rtf_read_raw)
names(texts) <- basename(table_files)

# Find the largest file
sizes <- vapply(texts, nchar, numeric(1))
cat(sprintf("File sizes: min=%.0f, median=%.0f, max=%.0f chars\n",
            min(sizes), median(sizes), max(sizes)))
cat(sprintf("Largest: %s (%.0f chars)\n\n", names(which.max(sizes)), max(sizes)))

# --- Benchmark find_matching_brace ---
# Find a brace in the largest file to benchmark
big <- texts[[which.max(sizes)]]
brace_pos <- regexpr("\\{", big)

cat("=== find_matching_brace ===\n")
n_iter <- 100L
t0 <- proc.time()
for (i in seq_len(n_iter)) {
  find_matching_brace(big, brace_pos)
}
elapsed <- (proc.time() - t0)["elapsed"]
cat(sprintf("  %d iterations on %.0f-char string: %.3fs (%.1fms/call)\n",
            n_iter, nchar(big), elapsed, elapsed / n_iter * 1000))

# --- Benchmark rtf_unescape ---
# Find a file with actual escapes, or use a synthetic one
cat("\n=== rtf_unescape ===\n")
# Synthetic: 500 \'xx escapes in a 10k char string
synth <- paste0(
  paste(rep("hello ", 500), collapse = ""),
  paste(sprintf("\\'%02x", sample(0x20:0x7e, 500, replace = TRUE)), collapse = " ")
)
cat(sprintf("  Synthetic string: %d chars, ~500 escapes\n", nchar(synth)))
n_iter <- 200L
t0 <- proc.time()
for (i in seq_len(n_iter)) {
  rtf_unescape(synth)
}
elapsed <- (proc.time() - t0)["elapsed"]
cat(sprintf("  %d iterations: %.3fs (%.1fms/call)\n",
            n_iter, elapsed, elapsed / n_iter * 1000))

# --- Benchmark rtf_cell_to_text ---
cat("\n=== rtf_cell_to_text ===\n")
# Grab a raw cell chunk from the largest file
page_texts <- rtf_split_pages(big)
if (length(page_texts) > 0) {
  body <- remove_groups(page_texts[1], c(
    "\\header", "\\footer", "\\headerf", "\\footerf",
    "\\headerl", "\\headerr", "\\footerl", "\\footerr"
  ))
  chunks <- strsplit(body, "\\\\cell(?![a-zA-Z])", perl = TRUE)[[1]]
  chunks <- chunks[nchar(chunks) > 10]
  if (length(chunks) > 0) {
    test_cell <- chunks[1]
    cat(sprintf("  Cell chunk: %.0f chars\n", nchar(test_cell)))
    n_iter <- 1000L
    t0 <- proc.time()
    for (i in seq_len(n_iter)) {
      rtf_cell_to_text(test_cell)
    }
    elapsed <- (proc.time() - t0)["elapsed"]
    cat(sprintf("  %d iterations: %.3fs (%.3fms/call)\n",
                n_iter, elapsed, elapsed / n_iter * 1000))
  }
}

# --- Benchmark full parse_rtf pipeline ---
cat("\n=== Full parse_rtf (all table files) ===\n")
n_iter <- 3L
t0 <- proc.time()
for (i in seq_len(n_iter)) {
  for (f in table_files) {
    parse_rtf(f)
  }
}
elapsed <- (proc.time() - t0)["elapsed"]
cat(sprintf("  %d passes over %d files: %.3fs (%.1fms/file avg)\n",
            n_iter, length(table_files), elapsed,
            elapsed / n_iter / length(table_files) * 1000))

# --- Breakdown: where does time go in parse_rtf? ---
cat("\n=== Parse breakdown (largest file) ===\n")
big_path <- table_files[which.max(sizes)]

t0 <- proc.time()
for (i in seq_len(10)) text <- rtf_read_raw(big_path)
cat(sprintf("  rtf_read_raw:    %.1fms\n", (proc.time() - t0)["elapsed"] / 10 * 1000))

t0 <- proc.time()
for (i in seq_len(10)) pages <- rtf_split_pages(text)
cat(sprintf("  rtf_split_pages: %.1fms\n", (proc.time() - t0)["elapsed"] / 10 * 1000))

t0 <- proc.time()
for (i in seq_len(10)) lapply(pages, parse_page)
cat(sprintf("  parse_page:      %.1fms\n", (proc.time() - t0)["elapsed"] / 10 * 1000))

# Drill into parse_page
page1 <- pages[1]
t0 <- proc.time()
for (i in seq_len(10)) extract_group(page1, "\\header")
cat(sprintf("    extract_group:   %.1fms\n", (proc.time() - t0)["elapsed"] / 10 * 1000))

t0 <- proc.time()
for (i in seq_len(10)) remove_groups(page1, c(
  "\\header", "\\footer", "\\headerf", "\\footerf",
  "\\headerl", "\\headerr", "\\footerl", "\\footerr"
))
cat(sprintf("    remove_groups:   %.1fms\n", (proc.time() - t0)["elapsed"] / 10 * 1000))

header_text <- extract_group(page1, "\\header")
body_text <- remove_groups(page1, c(
  "\\header", "\\footer", "\\headerf", "\\footerf",
  "\\headerl", "\\headerr", "\\footerl", "\\footerr"
))

t0 <- proc.time()
for (i in seq_len(10)) parse_rtf_table(body_text)
cat(sprintf("    parse_rtf_table: %.1fms\n", (proc.time() - t0)["elapsed"] / 10 * 1000))

cat("\nDone.\n")
