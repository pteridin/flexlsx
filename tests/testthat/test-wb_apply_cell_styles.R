test_that("error if sheet is non-existant", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(mtcars)
  wb <- openxlsx2::wb_workbook()

  wb_add_flextable(wb = wb,
                   ft = ft,
                   sheet = "nonexistant",
                   dims = "B4") |>
    expect_error()

  wb_apply_cell_styles(wb, "nonexistant", NULL) |>
    expect_error()
})

test_that("add fill", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(mtcars)
  wb <- openxlsx2::wb_workbook()$add_worksheet("fill")


  ft <- flextable::bg(ft,
           i = ~ am == 1,
           j = ~ am,
           bg = "orange",
           part = "body") |>
    flextable::bg(i = ~ hp > 100,
                  j = ~ hp,
                  bg = "red",
                  part = "body")


  wb <- wb_add_flextable(wb = wb,
                           ft = ft,
                           sheet = "fill",
                           dims = "D5")


  expect_true(wb$get_cell_style("fill", dims = "G6") !=
                wb$get_cell_style("fill", dims = "H6"))
  expect_true(wb$get_cell_style("fill", dims = "L6") !=
                wb$get_cell_style("fill", dims = "H6"))
  expect_true(wb$get_cell_style("fill", dims = "L6") !=
                wb$get_cell_style("fill", dims = "G6"))

})
