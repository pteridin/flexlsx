test_that("rowwise merge works", {
  skip_if_not_installed("flextable")
  data("mtcars")

  sheet <- "iris"
  ft <- mtcars |>
    head() |>
    flextable::flextable()

  ft <- flextable::merge_h(ft, i = 5)

  wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)

  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = sheet,
                         start_col = 2,
                         start_row = 2)

  df <- openxlsx2::wb_read(wb,
                           sheet=sheet,
                           start_row = 2,
                           start_col = 2,
                           col_names = T)

  df2 <- mtcars |>
    head()
  rownames(df2) <- NULL


  expect_equal(as.numeric(unlist(df)),
               as.numeric(unlist(df2)))

  NULL
})


test_that("columnwise merge works", {
  skip_if_not_installed("flextable")
  data("mtcars")

  sheet <- "iris"
  ft <- mtcars |>
    head() |>
    flextable::flextable()

  ft <- flextable::merge_v(ft, j = ~ vs + am + gear)

  wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)

  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = sheet,
                         start_col = 2,
                         start_row = 2)

  df <- openxlsx2::wb_read(wb,
                           sheet=sheet,
                           start_row = 2,
                           start_col = 2,
                           col_names = T)

  df2 <- mtcars |>
    head()
  rownames(df2) <- NULL


  expect_equal(as.numeric(unlist(df)),
               as.numeric(unlist(df2)))

  NULL
})

test_that("complex merge works", {
  skip_if_not_installed("flextable")
  data("mtcars")

  sheet <- "iris"
  ft <- mtcars |>
    head() |>
    flextable::flextable()

  ft <- flextable::merge_v(ft, j = ~ vs + am + gear)
  ft <- flextable::merge_h(ft, i = 5)

  wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)

  #TODO: Complex merge DOES not work right now!
  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = sheet,
                         start_col = 2,
                         start_row = 2) |>
    expect_error()

  NULL
})
