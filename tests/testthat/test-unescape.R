# Escape decoding: \'xx (CP1252), \uN (unicode), \ucN (fallback length),
# and parity between the C fast-path and the pure-R fallback.

test_that("rtf_unescape decodes hex, unicode and uc fallbacks", {
  expect_identical(rtf_unescape_r("Mean \\'b1 SD"), "Mean ± SD")
  expect_identical(rtf_unescape_r("\\'93q\\'94"), "“q”")   # CP1252 smart quotes
  expect_identical(rtf_unescape_r("\\'e9t\\'e9"), "été")
  expect_identical(rtf_unescape_r("less \\u8804?"), "less ≤")
  expect_identical(rtf_unescape_r("\\uc1\\u8804\\'3f x"), "≤ x")
  # \uc0: no fallback char; the single space is the control-word delimiter
  expect_identical(rtf_unescape_r("\\uc0\\u8805 next"), "≥next")
  expect_identical(rtf_unescape_r("\\uc2\\u916 XY tail"), "Δ tail")
  expect_identical(rtf_unescape_r("\\u-3844?"), intToUtf8(-3844 + 65536))
  expect_identical(rtf_unescape_r("plain text"), "plain text")
})

test_that("C and R unescape implementations agree", {
  skip_if_not(.c_available(), "C fast-path not compiled")
  cases <- c(
    "Mean \\'b1 SD", "\\'93q\\'94", "\\'e9t\\'e9", "less \\u8804?",
    "\\uc1\\u8804\\'3f x", "\\uc0\\u8805 next", "\\uc2\\u916 XY tail",
    "\\u-3844?", "plain text", "a\\u8211?b"
  )
  for (case in cases) {
    expect_identical(.Call(C_rtf_unescape, case), rtf_unescape_r(case),
                     label = sprintf("C output for %s", deparse(case)))
  }
})

test_that("rtf_cell_to_text strips markup and decodes escapes", {
  cases <- list(
    list(in_ = "Mean \\'b1 SD\\cell",                want = "Mean ± SD"),
    list(in_ = "\\'93quoted\\'94\\cell",             want = "“quoted”"),
    list(in_ = "less \\u8804?\\cell",                want = "less ≤"),
    list(in_ = "less \\uc1\\u8804\\'3f x\\cell",     want = "less ≤ x"),
    list(in_ = "\\uc0\\u8805 next\\cell",            want = "≥next"),
    list(in_ = "{\\b Bold head\\b0}\\cell",          want = "Bold head"),
    list(in_ = "\\ql\\fs20 plain text\\cell",        want = "plain text"),
    list(in_ = "keep\\~this\\cell",                  want = "keepthis"),
    list(in_ = "{\\i \\u945?}\\cell",                want = "α"),
    list(in_ = "{\\*\\fldinst HYPERLINK}text\\cell", want = "text"),
    list(in_ = "{a{\\b b}c}\\cell",                  want = "abc")
  )
  for (case in cases) {
    expect_identical(rtf_cell_to_text_r(case$in_), case$want,
                     label = sprintf("R output for %s", deparse(case$in_)))
    expect_identical(rtf_cell_to_text(case$in_), case$want,
                     label = sprintf("dispatch output for %s", deparse(case$in_)))
  }
})
