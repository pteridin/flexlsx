
## Create a test dataset
create_df <- function(m) {
  m |>
    as.data.frame() |>
    mutate(row_id = 1:dplyr::n()) |>
    tidyr::pivot_longer(cols = -all_of("row_id"),
                        names_to = "name",
                        values_to = "style") |>
    mutate(col_id = rep(1:ncol(m), nrow(m)), style_other = T) |>
    select(-all_of("name"))
}

matrix(
  c(".",".",".",".",".",".",
    "-","-","-","-","-","-",
    "-",".",".",".",".","-",
    "-",".",".",".",".","-",
    "-",".",".",".",".","-",
    "-","-","-","-","-","-",
    "-",".","-","-",".","-",
    "-",".","-","-",".","-",
    ".","-",".",".",".","."),
  ncol = 6,
  byrow = TRUE) -> m

df <- create_df(m)



## recreate matrix
recreate_matrix <- function(df, df_hashes) {
  if ("row_id" %in% names(df)) {
    df$row_from <- df$row_id
    df$row_to <- df$row_id
  }

  df$hash <- factor(df$hash, df_hashes$hash, df_hashes$style) |>
    as.character()


  m2 <- matrix(NA_character_,
               nrow = max(df$row_to),
               ncol = max(df$col_to))
  for (i in seq_len(nrow(df))) {
    m2[df$row_from[i]:df$row_to[i], df$col_from[i]:df$col_to[i]] <- df$hash[i]
  }

  return(m2)
}


test_that("style_to_hash works", {
  df_hashes <- df |>
    style_to_hash()

  df_reference <- tibble::tribble(~ style, ~ style_other, ~ hash, "-", T, 1L, ".", T, 2L)

  attr(df_reference, "cols_to_join") <- c("style", "style_other")

  testthat::expect_equal(df_hashes, df_reference)
})


test_that("get_dim_rowwise works", {
  df_hashes <- df |>
    style_to_hash()

  df_rowwise <- df |>
    get_dim_rowwise(df_hashes)

  expect_equal(recreate_matrix(df_rowwise, df_hashes), m)

  set.seed(123)
  for (i in 1:10) {
    m2 <- matrix(sample(
      c("-", ".", "+"),
      100,
      replace = TRUE,
      prob = c(0.1, 0.8, 0.1)
    ), nrow = 10)
    df2 <- create_df(m2)
    df_hashes <- df2 |>
      style_to_hash()

    df_rowwise <- df2 |>
      get_dim_rowwise(df_hashes)

    expect_equal(recreate_matrix(df_rowwise, df_hashes), m2)
  }
})

test_that("get_dim_colwise works", {
  df_hashes <- df |>
    style_to_hash()

  df_rowwise <- df |>
    get_dim_rowwise(df_hashes)

  df_colwise <- df_rowwise |>
    get_dim_colwise()


  expect_equal(recreate_matrix(df_colwise, df_hashes), m)

  set.seed(123)
  for (i in 1:10) {
    m2 <- matrix(sample(
      c("-", ".", "+"),
      100,
      replace = TRUE,
      prob = c(0.1, 0.8, 0.1)
    ), nrow = 10)
    df2 <- create_df(m2)
    df_hashes <- df2 |>
      style_to_hash()

    df_rowwise <- df2 |>
      get_dim_rowwise(df_hashes)

    df_colwise <- df_rowwise |>
      get_dim_colwise()

    expect_equal(recreate_matrix(df_colwise, df_hashes), m2)
  }

})

test_that("get_dim_ranges works", {
  df_ranges <- df |>
    get_dim_ranges()

  expect_equal(sum(!df_ranges$multi_lines), 2L)
  expect_true(all(df_ranges$style_other))
  expect_true("style" %in% names(df_ranges))
})

