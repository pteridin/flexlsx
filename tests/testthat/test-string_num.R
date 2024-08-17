test_that("option string_num is supported", {

  library(flextable)

  ft <- flextable(airquality[seq_len(10), ])
  ft <- add_header_row(ft,
                       colwidths = c(4, 2),
                       values = c("Air quality", "Time")
  )
  ft <- theme_vanilla(ft)
  ft <- add_footer_lines(ft, "Daily air quality measurements in New York, May to September 1973.")
  ft <- color(ft, part = "footer", color = "#666666")
  ft <- set_caption(ft, caption = "New York Air Quality Measurements")
  ft

  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")

  options("openxlsx2.string_nums" = TRUE)
  wb <- flexlsx::wb_add_flextable(wb, "mtcars", ft, dims = "C2")
  cc <- wb$worksheets[[1]]$sheet_data$cc
  expect_equal(cc[cc$r == "C5", "v"], "41")
  options("openxlsx2.string_nums" = NULL)

})
