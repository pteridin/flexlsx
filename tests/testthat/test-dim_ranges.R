
## Create a test dataset
matrix(
  c(".",".",".",".",".",".",
    "-","-","-","-","-","-",
    "-",".",".",".",".","-",
    "-",".",".",".",".","-",
    "-",".",".",".",".","-",
    "-","-","-","-","-","-",
    "-",".","-","-",".","-",
    "-",".","-","-",".","-",
    ".",".",".",".",".","."),
  ncol = 6,
  byrow = TRUE) -> m

m |>
  as.data.frame() |>
  mutate(row_id = 1:n()) |>
  tidyr::pivot_longer(cols = -row_id,
                      names_to = "name",
                      values_to = "style") |>
  mutate(col_id = rep(1:6, nrow(m)),
         style_other = T) |>
  select(-name) -> df

## recreate matrix
recreate_matrix <- function(df) {
  if("row_id" %in% names(df)) {
    df$row_from <- df$row_id
    df$row_to <- df$row_id
  }

  df$hash <- factor(df$hash, c(1L,2L), c("-",".")) |>
    as.character()


  m2 <- matrix(NA_character_,
               nrow = max(df$row_to),
               ncol = max(df$col_to))
  for(i in seq_len(nrow(df))) {
    m2[df$row_from[i]:df$row_to[i],
       df$col_from[i]:df$col_to[i]] <- df$hash[i]
  }

  return(m2)
}


test_that("style_to_hash works",
          {

  df_hashes <- df |>
    style_to_hash()

  df_reference <- tibble::tribble(
    ~style, ~style_other, ~hash,
    "-",    T,            1L,
    ".",    T,            2L
  )

  attr(df_reference, "cols_to_join") <- c("style", "style_other")

  testthat::expect_equal(df_hashes,
                         df_reference)
})


test_that("get_dim_rowwise works", {
            df_hashes <- df |>
              style_to_hash()

            df_rowwise <- df |>
              get_dim_rowwise(df_hashes)

            expect_equal(recreate_matrix(df_rowwise), m)
})

test_that("get_dim_colwise works",
          {
            df_hashes <- df |>
              style_to_hash()

            df_rowwise <- df |>
              get_dim_rowwise(df_hashes)

            df_colwise <- df_rowwise |>
              get_dim_colwise()


            expect_equal(recreate_matrix(df_colwise), m)

          })


