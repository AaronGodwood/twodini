# suggest_match.R - fuzzy pairing of bookmark names to RTF table names.
#
# Clinical/SAS outputs number tables (e.g. "14.1") and bookmarks tend to echo
# that number, so the strongest signal is a matching numeric key; shared word
# tokens and raw edit distance are the fallbacks. Base R only (adist for edit
# distance) - no stringdist dependency.

# Split a name into lowercase alphanumeric tokens and an ordered numeric key.
# "T_14_1_Chol" -> list(tokens = c("t","14","1","chol"), number = "14.1")
normalise_name <- function(name) {
  name <- tolower(as.character(name %||% ""))
  tokens <- strsplit(name, "[^a-z0-9]+", perl = TRUE)[[1]]
  tokens <- tokens[nzchar(tokens)]
  nums   <- tokens[grepl("^[0-9]+$", tokens)]
  list(
    tokens = tokens,
    number = if (length(nums) > 0L) paste(nums, collapse = ".") else ""
  )
}

# Jaccard overlap of two token sets (|A n B| / |A u B|), 0 when both empty.
token_jaccard <- function(a, b) {
  a <- unique(a)
  b <- unique(b)
  if (length(a) == 0L && length(b) == 0L) return(0)
  inter <- length(intersect(a, b))
  union <- length(union(a, b))
  if (union == 0L) 0 else inter / union
}

# Normalised edit similarity in [0, 1]: 1 - levenshtein / longer length.
edit_similarity <- function(a, b) {
  a <- tolower(a)
  b <- tolower(b)
  m <- max(nchar(a), nchar(b))
  if (m == 0L) return(0)
  1 - as.numeric(adist(a, b)[1, 1]) / m
}

#' Similarity score between two names in the range 0 to 1
#'
#' A shared numeric key (e.g. both contain "14.1") is treated as a strong signal
#' and floors the score near 1; otherwise the score blends token Jaccard and
#' normalised edit similarity.
#' @param a,b Names to compare
#' @return Numeric score between 0 and 1
match_score <- function(a, b) {
  na <- normalise_name(a)
  nb <- normalise_name(b)

  jac  <- token_jaccard(na$tokens, nb$tokens)
  edit <- edit_similarity(a, b)
  base <- 0.6 * jac + 0.4 * edit

  # Strong signal: identical non-empty numeric keys. Add a small token-overlap
  # bonus so several "14.1" candidates break ties on shared words.
  if (nzchar(na$number) && identical(na$number, nb$number)) {
    return(min(1, 0.9 + 0.1 * jac))
  }

  base
}

#' Best-matching candidate for a name
#'
#' @param name The name to match (e.g. a bookmark)
#' @param candidates Character vector of candidate names (e.g. table names)
#' @return list(match = best candidate or NA, score = its score from 0 to 1).
#'   Ties break toward a shared numeric key, then alphabetically, for
#'   determinism.
best_match <- function(name, candidates) {
  candidates <- candidates[nzchar(candidates)]
  if (length(candidates) == 0L) return(list(match = NA_character_, score = 0))

  scores <- vapply(candidates,
                   function(cand) match_score(name, cand),
                   numeric(1))

  best <- max(scores)
  tied <- candidates[scores == best]
  if (length(tied) > 1L) {
    key <- normalise_name(name)$number
    if (nzchar(key)) {
      shares_key <- vapply(
        tied,
        function(c) identical(normalise_name(c)$number, key),
        logical(1)
      )
      if (any(shares_key)) tied <- tied[shares_key]
    }
    tied <- sort(tied)
  }

  list(match = tied[1], score = best)
}
