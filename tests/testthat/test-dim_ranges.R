
test_that("dim_ranges does not break", {
  tibble::tibble(col_id = c(1, 1, 1, 2, 2, 2),
                 row_id = c(1, 2, 3, 1, 2, 3),
                 style = c("a", "a", "b", "a", "b", "b")) -> df

  get_dim_ranges(df) -> df_ranges
  df_ranges$hash -> hashes

  expect_equal(hashes[1], hashes[3])
  expect_equal(hashes[2], hashes[4])
  expect_equal(df_ranges |>
                 select(-hash),


               tibble::tribble(~col_id,  ~line, ~row_id, ~style, ~row_to, ~dims, ~multi_lines,
                                 1,      1,     1,  "a",  2,   "A1:A2", TRUE,
                                 1,      2,     3,  "b",  3,   "A3:A3", FALSE ,
                                 2,      1,     1,  "a",  1,   "B1:B1", FALSE,
                                 2,      2,     2,  "b",  3,   "B2:B3", TRUE))
})
