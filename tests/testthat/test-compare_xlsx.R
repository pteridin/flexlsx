test_that("compare_xlsx works", {
  tmpfile <- tempfile(fileext = ".xlsx")
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb$save(tmpfile)
  expect_snapshot_value(compare_xlsx(tmpfile))
})
