test_that("option string_num is supported", {
  skip_if_not_installed("flextable")
  library(flextable)

  ft <- flextable(airquality[seq_len(10), ])
  ft <- add_header_row(ft,
    colwidths = c(4, 2),
    values = c("Air quality", "Time")
  )
  ft <- theme_vanilla(ft)
  ft <- add_footer_lines(
    ft,
    "Daily air quality measurements in New York, May to September 1973."
  )
  ft <- color(ft, part = "footer", color = "#666666")
  ft <- set_caption(ft, caption = "New York Air Quality Measurements")
  ft

  options("openxlsx2.string_nums" = NULL)
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb <- flexlsx::wb_add_flextable(wb, "mtcars", ft, dims = "C2")

  options("openxlsx2.string_nums" = TRUE)
  wb <- wb$add_worksheet("mtcars numeric")
  wb <- flexlsx::wb_add_flextable(wb, "mtcars numeric", ft, dims = "C2")

  cc <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(cc[cc$r == "C5", "v"], "")

  cc <- wb$worksheets[[2]]$sheet_data$cc
  expect_equal(cc[cc$r == "C5", "v"], "41")

  test_wb_ft(wb, ft, "string_num")

  options("openxlsx2.string_nums" = NULL)
})
