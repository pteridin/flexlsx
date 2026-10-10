test_that("error if sheet is non-existant", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(mtcars)
  wb <- openxlsx2::wb_workbook()

  wb_add_flextable(
    wb = wb,
    ft = ft,
    sheet = "nonexistant",
    dims = "B4"
  ) |>
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
    j = ~am,
    bg = "orange",
    part = "body"
  ) |>
    flextable::bg(
      i = ~ hp > 100,
      j = ~hp,
      bg = "red",
      part = "body"
    )


  wb <- wb_add_flextable(
    wb = wb,
    ft = ft,
    sheet = "fill",
    dims = "D5"
  )


  expect_true(wb$get_cell_style("fill", dims = "G6") !=
    wb$get_cell_style("fill", dims = "H6"))
  expect_true(wb$get_cell_style("fill", dims = "L6") !=
    wb$get_cell_style("fill", dims = "H6"))
  expect_true(wb$get_cell_style("fill", dims = "L6") !=
    wb$get_cell_style("fill", dims = "G6"))
})

test_that("rotated text keeps its rotation", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(head(mtcars)[, c("mpg", "cyl", "disp")]) |>
    flextable::rotate(j = "mpg", rotation = "btlr", part = "header") |>
    flextable::rotate(j = "cyl", rotation = "tbrl", part = "header")

  wb <- openxlsx2::wb_workbook()$add_worksheet("rotate")
  wb <- wb_add_flextable(
    wb = wb,
    ft = ft,
    sheet = "rotate",
    dims = "A1"
  )

  cell_xf <- function(dims) {
    style_id <- as.integer(wb$get_cell_style("rotate", dims = dims))
    wb$styles_mgr$styles$cellXfs[[style_id + 1L]]
  }

  expect_match(cell_xf("A1"), 'textRotation="90"', fixed = TRUE)
  expect_match(cell_xf("B1"), 'textRotation="180"', fixed = TRUE)
  expect_false(grepl("textRotation", cell_xf("C1"), fixed = TRUE))
})

test_that("fill and border colors with an alpha channel export as opaque RGB", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(head(mtcars)[, c("mpg", "cyl", "disp")]) |>
    flextable::bg(j = "mpg", bg = "#EDB50199", part = "body") |>
    flextable::bg(j = "cyl", bg = "orange", part = "body") |>
    flextable::border_outer(
      border = officer::fp_border(color = "#00000080", width = 1),
      part = "body"
    )

  wb <- openxlsx2::wb_workbook()$add_worksheet("alpha")
  wb <- wb_add_flextable(
    wb = wb,
    ft = ft,
    sheet = "alpha",
    dims = "A1"
  )

  cell_xf <- function(dims) {
    style_id <- as.integer(wb$get_cell_style("alpha", dims = dims))
    wb$styles_mgr$styles$cellXfs[[style_id + 1L]]
  }
  style_of <- function(dims, what) {
    xf <- cell_xf(dims)
    id <- regmatches(xf, regexpr(paste0("(?<=", what, "Id=\")[0-9]+"), xf, perl = TRUE))
    wb$styles_mgr$styles[[paste0(what, "s")]][[as.integer(id) + 1L]]
  }

  # R writes #RRGGBBAA; Excel reads eight hex digits as AARRGGBB, so the
  # alpha byte must be dropped rather than passed through.
  expect_match(style_of("A2", "fill"), 'rgb="FFEDB501"', fixed = TRUE)
  expect_match(style_of("B2", "fill"), 'rgb="FFFFA500"', fixed = TRUE)
  expect_false(grepl("EDB50199", style_of("A2", "fill"), fixed = TRUE))
  expect_match(style_of("A2", "border"), 'rgb="FF000000"', fixed = TRUE)
  expect_false(grepl("00000080", style_of("A2", "border"), fixed = TRUE))
  # An unfilled cell stays unfilled.
  expect_false(grepl("patternType=\"solid\"", style_of("C2", "fill"), fixed = TRUE))
})
