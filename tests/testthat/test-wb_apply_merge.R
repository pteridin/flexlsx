test_that("rowwise merge works", {
  skip_if_not_installed("flextable")
  data("mtcars")

  sheet <- "iris"
  ft <- mtcars |>
    head() |>
    flextable::flextable()

  ft <- flextable::merge_h(ft, i = 5)

  wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)

  wb <- wb_add_flextable(
    wb = wb,
    ft = ft,
    sheet = sheet,
    start_col = 2,
    start_row = 2
  )

  test_wb_ft(wb, ft, "merge_row")

  df <- openxlsx2::wb_read(wb,
    sheet = sheet,
    start_row = 2,
    start_col = 2,
    col_names = TRUE
  )

  df2 <- mtcars |>
    head()
  rownames(df2) <- NULL
  df2[5, 9] <- NA

  expect_equal(
    as.numeric(unlist(df)),
    as.numeric(unlist(df2))
  )

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

  wb <- wb_add_flextable(
    wb = wb,
    ft = ft,
    sheet = sheet,
    start_col = 2,
    start_row = 2
  )
  test_wb_ft(wb, ft, "merge_col")

  df <- openxlsx2::wb_read(wb,
    sheet = sheet,
    start_row = 2,
    start_col = 2,
    col_names = TRUE
  )

  df2 <- mtcars |>
    head() |>
    dplyr::mutate(dplyr::across(
      c(vs, am, gear),
      ~ ifelse(coalesce(lag(.x), -1) == .x,
        NA, .x
      )
    ))
  rownames(df2) <- NULL

  expect_equal(
    as.numeric(unlist(df)),
    as.numeric(unlist(df2))
  )

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

  wb <- openxlsx2::wb_workbook()$add_worksheet(sheet) |>
    wb_add_flextable(
      ft = ft,
      sheet = sheet,
      start_col = 2,
      start_row = 2
    ) |>
    expect_warning()

  NULL
})
