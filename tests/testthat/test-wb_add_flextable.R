test_that(
  "flextable without header", {
    skip_if_not_installed("flextable")

    sheet <- "iris"
    ft <- datasets::iris |>
      head() |>
      flextable::flextable() |>
      flextable::delete_part(part = "header")
    wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)
    dims <- "B2"


    wb <- wb_add_flextable(wb = wb,
                           ft = ft,
                           sheet = sheet,
                           dims = dims)

    df <- openxlsx2::wb_read(wb,
                             sheet=sheet,
                             start_row = 2,
                             start_col = 2,
                             col_names = F)
    df$F <- NULL
    df2 <- datasets::iris |>
      head()
    df2$Species <- NULL


    expect_equal(as.numeric(unlist(df)),
                 as.numeric(unlist(df2)))

    NULL
  }
)

test_that("Add with numeric offset", {
  skip_if_not_installed("flextable")
  data("mtcars")

  sheet <- "iris"
  ft <- mtcars |>
    head() |>
    flextable::flextable()
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

test_that("Add multi-header", {
  skip_if_not_installed("flextable")

  typology <- data.frame(
    col_keys = c(
      "Sepal.Length", "Sepal.Width", "Petal.Length",
      "Petal.Width", "Species"
    ),
    what = c("Sepal", "Sepal", "Petal", "Petal", "Species"),
    measure = c("Length", "Width", "Length", "Width", "Species"),
    stringsAsFactors = FALSE
  )

  ft_1 <- flextable::flextable(head(iris)) |>
    flextable::set_header_df(mapping = typology, key = "col_keys") |>
    flextable::merge_h(part = "header") |>
    flextable::merge_v( j = "Species", part = "header") |>
    flextable::theme_vanilla() |>
    flextable::fix_border_issues() |>
    flextable::autofit()

  wb <- openxlsx2::wb_workbook()$add_worksheet("multiheader")

  wb <- wb_add_flextable(wb = wb,
                         ft = ft_1,
                         sheet = "multiheader",
                         start_col = 2,
                         start_row = 2)

  expect_equal(openxlsx2::wb_read(wb,
                           sheet="multiheader",
                           start_row = 2,
                           start_col = 2,
                           col_names = T) |>
                colnames(),
               c("Sepal", "Sepal", "Petal", "Petal", "Species"))

  expect_equal(openxlsx2::wb_read(wb,
                                  sheet="multiheader",
                                  start_row = 3,
                                  start_col = 2,
                                  col_names = T) |>
                 colnames(),
               c("Length", "Width", "Length", "Width", "Species"))

  NULL
})

test_that("using openxlsx2::current_sheet() works", {
  skip_if_not_installed("flextable")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))

  openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    flexlsx::wb_add_flextable(
      sheet = openxlsx2::current_sheet(),
      ft = ft,
      dims = "C2"
    ) -> wb

  expect_equal(
    wb$get_sheet_names(),
    c(`Sheet 1` = "Sheet 1")
  )

  openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    flexlsx::wb_add_flextable(
      ft = ft,
      dims = "A1"
    ) -> wb

  expect_equal(
    wb$get_sheet_names(),
    c(`Sheet 1` = "Sheet 1")
  )

  expect_equal(names(wb$to_df(sheet  = "Sheet 1")),
               c("mpg", NA, "cyl", "cyl", "cyl", "cyl"))

})

test_that("When sheet does not exists throws an error", {
  ft <- flextable::as_flextable(table(mtcars[,1:2]))

  expect_error( openxlsx2::wb_workbook() |>
                  flexlsx::wb_add_flextable(
                    sheet = openxlsx2::current_sheet(),
                    ft = ft,
                    dims = "C2"
                  ),
                regexp = "Sheet 'current_sheet' does not exist!")

  expect_error( openxlsx2::wb_workbook() |>
                  flexlsx::wb_add_flextable(
                    sheet = "test",
                    ft = ft,
                    dims = "C2"
                  ),
                regexp = "Sheet 'test' does not exist!")
})
