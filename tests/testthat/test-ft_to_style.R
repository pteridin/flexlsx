
test_that("ft_to_style_tibble does not break", {
  ft <- flextable::as_flextable(table(mtcars[,1:2]))

  flexlsx:::ft_to_style_tibble(ft,
                     offset_rows = 0L, offset_cols = 0L, offset_caption_rows = 0L) -> x
  expect_snapshot_value(as.list(x), style = "json2")
})

test_that("ft_to_style_tibble does not break with offsets", {
  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  flexlsx:::ft_to_style_tibble(ft,
                     offset_rows = 5L, offset_cols = 2L, offset_caption_rows = 8L) -> y
  expect_snapshot_value(as.list(y), style = "json2")
})
