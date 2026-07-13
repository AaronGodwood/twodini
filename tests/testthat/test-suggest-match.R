# Fuzzy bookmark <-> table matcher (ported from tests/test_suggest_match.R).

test_that("normalise_name tokenises and builds numeric keys", {
  n <- normalise_name("T_14_1_Chol")
  expect_identical(n$tokens, c("t", "14", "1", "chol"))
  expect_identical(n$number, "14.1")
  expect_identical(normalise_name("demographics")$number, "")
  # Order matters: 14.1 != 1.41
  expect_false(identical(normalise_name("t_14_1")$number,
                         normalise_name("t_1_41")$number))
})

test_that("match_score behaves sensibly", {
  s <- match_score("aaa", "zzz")
  expect_true(s >= 0 && s <= 1)
  expect_gte(match_score("x_14_1", "y_14_1"), 0.9)
  expect_gt(match_score("x_14_1", "y_14_1"), match_score("x_14_1", "x_14_2"))
  expect_equal(match_score("demog", "demog"), 1)
})

test_that("best_match picks the right candidate", {
  r1 <- best_match("T_14_1_Chol", c("table_14_1", "table_14_2", "demographics"))
  expect_identical(r1$match, "table_14_1")

  r2 <- best_match("demog", c("demographics", "ae_summary", "table_14_1"))
  expect_identical(r2$match, "demographics")

  expect_true(is.na(best_match("foo", character())$match))
  expect_true(is.na(best_match("foo", c("", ""))$match))

  # Determinism: ties break toward shared numeric key, then alphabetically
  r3 <- best_match("t_14_1", c("z_14_1", "a_14_1"))
  expect_identical(r3$match, "a_14_1")
})
