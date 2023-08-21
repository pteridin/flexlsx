test_that("Generation from start to finish works", {
  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb_add_flextable(wb, "mtcars", ft)$save(tmpfile)

})

test_that("Offsets work", {
  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  wb <- openxlsx2::wb_workbook()$
    add_worksheet("mtcars_offset_B3")$
    add_worksheet("mtcars_offset_52")
  wb <- wb_add_flextable(wb, "mtcars_offset_B3", ft, dims = "B3")
  wb_add_flextable(wb, "mtcars_offset_52", ft, start_col = 5, start_row = 2)$save(tmpfile)
})
