
test_that("ft_to_style_tibble does not break", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  ft <- as_flextable(table(mtcars[,1:2]))

  flexlsx:::ft_to_style_tibble(ft,
                     offset_rows = 0L, offset_cols = 0L, offset_caption_rows = 0L) -> x

  # Fix some platforms that use other default fonts & row-heights
  x <- x |> select(-ends_with("family"),
                   -all_of("rowheight"))

  expect_snapshot_value(as.list(x), style = "json2")
})

test_that("ft_to_style_tibble does not break with offsets", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  ft <- as_flextable(table(mtcars[,1:2]))
  flexlsx:::ft_to_style_tibble(ft,
                     offset_rows = 5L, offset_cols = 2L, offset_caption_rows = 8L) -> y

  # Fix some platforms that use other default fonts & row-heights
  y <- y |> select(-ends_with("family"),
                   -all_of("rowheight"))

  expect_snapshot_value(as.list(y), style = "json2")
})
