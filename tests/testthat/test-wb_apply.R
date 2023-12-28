
test_that("Generation from start to finish works", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb_add_flextable(wb, "mtcars", ft)$save(tmpfile)

  y <- compare_xlsx(tmpfile)
  expect_snapshot_value(y)
})

test_that("Offsets work", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  wb <- openxlsx2::wb_workbook()$
    add_worksheet("mtcars_offset_B3")$
    add_worksheet("mtcars_offset_52")
  wb <- wb_add_flextable(wb, "mtcars_offset_B3", ft, dims = "B3")
  wb_add_flextable(wb, "mtcars_offset_52", ft, start_col = 5, start_row = 2)$save(tmpfile)

  y <- compare_xlsx(tmpfile)
  expect_snapshot_value(y)
})



test_that("Simple Caption works", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  ft <- flextable::set_caption(ft, "Simple Caption")
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb_add_flextable(wb, "mtcars", ft)$save(tmpfile)

  y <- compare_xlsx(tmpfile)
  expect_snapshot_value(y)
})

test_that("Complex Caption works", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  ft <- flextable::set_caption(ft,
                                       caption = flextable::as_paragraph('a ', flextable::as_b('bold'),
                                                                         " and <br>",
                                                                         flextable::as_i('italic'),
                                                                         ' text <br /> with <br/>',
                                                                         flextable::as_chunk("Variations!", props = flextable::fp_text_default(color = "orange",
                                                                                                                                               font.family = "Courier",
                                                                                                                                               underlined = T))
                                       ))
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb_add_flextable(wb, "mtcars", ft, offset_caption_rows = 1L)$save(tmpfile)

  y <- compare_xlsx(tmpfile)
  expect_snapshot_value(y)
})

test_that("Complex gtsummary works", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")

  data("Titanic")

  tibble::as_tibble(Titanic) |>
    tidyr::uncount(n) |>
    gtsummary::tbl_strata(strata = "Class",
                          .tbl_fun = \(x) {
                            gtsummary::tbl_summary(x,
                                                   by="Sex") |>
                              gtsummary::add_difference(everything() ~ "prop.test")
                          }) |>
    gtsummary::tbl_butcher() |>
    gtsummary::bold_labels() |>
    gtsummary::italicize_levels() |>
    gtsummary::as_flex_table() -> ft


  wb <- openxlsx2::wb_workbook()$add_worksheet("titanic")
  wb_add_flextable(wb, "titanic", ft, offset_caption_rows = 1L)$save(tmpfile)

  y <- compare_xlsx(tmpfile)
  expect_snapshot_value(y)
})


test_that("Illegal XML characters work", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")
  to_check <- c("1 (<0.1%)",
                "1 > 100",
                "&amp; 1 == 1",
                "'Hallo'")

  data.frame(IllegalXML = to_check) |>
    flextable::flextable() -> ft

  wb <- openxlsx2::wb_workbook()$add_worksheet("titanic")
  wb_add_flextable(wb, "titanic", ft, offset_caption_rows = 1L)$save(tmpfile)

  expect_true(T)
})


test_that("Linebreaks work", {
  skip_if_not_installed("flextable")

  tmpfile <- tempfile(fileext = ".xlsx")
  to_check <- c("Hello<br>Linebreak",
                "Hello<br><br><br>Linebreak2",
                "Ein<br>Zeilenbrecher",
                "Zwei<br>Zeilen<br>echer")
  to_check2 = c("Hello kein Linebreak",
                "Hello ein <br>Linebreak",
                "Drei<br><br><br>Zeilenbrecher",
                "Ein<br>Zeilenbrecher")

  data.frame(Linebreak1 = to_check,
             Linebreak2 = to_check2) |>
    flextable::flextable() |>
    flextable::autofit(part = "body")-> ft

  wb <- openxlsx2::wb_workbook()$add_worksheet("titanic")
  wb_add_flextable(wb, "titanic", ft, offset_caption_rows = 1L)$save(tmpfile)

  expect_true(T)
})
