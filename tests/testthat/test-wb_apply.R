
compare_xlsx <- function(ft,
                         tmpfile,
                         sheet_name,
                         start_col = 1,
                         start_row = 1,
                         offset_caption_rows = 0L,
                         dims = NULL) {

  # Retrieve offsets
  if (!is.null(dims)) {
    dims <- openxlsx2::dims_to_rowcol(dims, as_integer = TRUE)
    offset_cols <- min(dims[[1]]) - 1L
    offset_rows <- min(dims[[2]]) - 1L
  } else {
    offset_cols <- start_col - 1L
    offset_rows <- start_row - 1L
  }

  # ignore offset if there is no caption
  if(length(ft$caption$value) == 0) {
    offset_caption_rows <- 0L
  }

  ft_style <- flexlsx:::ft_to_style_tibble(ft,
                                           offset_rows = offset_rows,
                                           offset_cols = offset_cols,
                                           offset_caption_rows = offset_caption_rows)
  wb <- openxlsx2::wb_load(tmpfile)

  ft_range <- ft_style |>
    summarize(across(all_of(c("col_id", "row_id")),
                     list(Min = min, Max = max)),
              .groups = "drop") |>
    mutate(r = paste0(openxlsx2::int2col(col_id_Min), row_id_Min, ":",
                      openxlsx2::int2col(col_id_Max), row_id_Max))


  ## Check content
  diff_content <- compare_content(ft,
                                  ft_style,
                                  ft_range,
                                  wb, sheet_name,
                                  offset_rows + offset_caption_rows,
                                  offset_cols)
  nrow_diff_content <- nrow(diff_content)

  expect_equal(nrow_diff_content, 0)
  if(nrow_diff_content > 0) {
    print(paste0("Differences in content ", sheet_name, ":"))
    print(diff_content)
  }

  ## TODO: More tests (cell style, text style, ...)
}

compare_content <- function(ft,
                            ft_style, ft_range, wb, sheet_name,
                            offset_rows,
                            offset_cols) {

  if(flextable:::has_caption(ft))
    offset_rows <- offset_rows + 1L


  content_target <- select(ft_style,
                           all_of(c("row_id", "col_id", "content"))) |>
    unnest_legacy() |>
    group_by(across(all_of(c("row_id", "col_id")))) |>
    summarize(content = paste0(txt, collapse = ""),
              .groups = "drop")

  content_is <- wb$to_df(sheet = sheet_name,
                         dims = ft_range$r,
                         col_names = FALSE) |>
    mutate(across(everything(), ~ as.character(.x)),
      row_id = dplyr::row_number() + offset_rows) |>
    tidyr::pivot_longer(-all_of("row_id")) |>
    group_by(across(all_of("row_id"))) |>
    mutate(col_id = dplyr::row_number() + offset_cols)

  content_is |>
    dplyr::full_join(content_target, by = c("row_id", "col_id")) |>
    dplyr::mutate(value = dplyr::coalesce(value, ""),
                  content = dplyr::coalesce(content, "")) |>
    filter(value != content)
}


test_that("Generation from start to finish works", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb_add_flextable(wb, "mtcars", ft)$save(tmpfile)

  compare_xlsx(ft, tmpfile, "mtcars")

})

test_that("Offsets work", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  wb <- openxlsx2::wb_workbook()$
    add_worksheet("mtcars_offset_B3")$
    add_worksheet("mtcars_offset_52")
  wb <- wb_add_flextable(wb, "mtcars_offset_B3", ft, dims = "B3")
  wb_add_flextable(wb, "mtcars_offset_52", ft, start_col = 5, start_row = 2)$save(tmpfile)

  compare_xlsx(ft, tmpfile,
               sheet_name = "mtcars_offset_B3",
               dims = "B3")
  compare_xlsx(ft, tmpfile,
               sheet_name = "mtcars_offset_52",
               start_col = 5,
               start_row = 2)
})



test_that("Simple Caption works", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  tmpfile <- tempfile(fileext = ".xlsx")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))
  ft <- flextable::set_caption(ft, "Simple Caption")
  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
  wb_add_flextable(wb, "mtcars", ft)$save(tmpfile)

  compare_xlsx(ft, tmpfile,
               sheet_name = "mtcars")
})

test_that("Complex Caption works", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

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

  compare_xlsx(ft, tmpfile,
               sheet_name = "mtcars",
               offset_caption_rows = 1L)
})

test_that("Complex gtsummary works", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

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

  compare_xlsx(ft, tmpfile,
               sheet_name = "titanic",
               offset_caption_rows = 1L)
})


test_that("Illegal XML characters work", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

  tmpfile <- tempfile(fileext = ".xlsx")
  to_check <- c("1 (<0.1%)",
                "1 > 100",
                "&amp; 1 == 1",
                "'Hallo'")

  data.frame(IllegalXML = to_check) |>
    flextable::flextable() -> ft

  wb <- openxlsx2::wb_workbook()$add_worksheet("titanic")
  wb_add_flextable(wb, "titanic", ft, offset_caption_rows = 1L)$save(tmpfile)

  wb <- openxlsx2::wb_load(tmpfile)

  expect_equal(wb$to_df(sheet = "titanic",
                  dims = "A1:A5",
                  col_names = FALSE)$A,
               c("IllegalXML",
                 "1 (<0.1%)",
                 "1 > 100",
                 "&amp; 1 == 1",
                 "'Hallo'"))
})


test_that("Linebreaks work", {
  skip_if_not_installed("flextable")
  require("flextable", quietly = TRUE)

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

  wb <- openxlsx2::wb_load(tmpfile)

  expect_equal(wb$to_df(sheet = "titanic",
                        dims = "A1:A5",
                        col_names = FALSE)$A,
               c("Linebreak1",
                 "Hello\nLinebreak",
                 "Hello\n\n\nLinebreak2",
                 "Ein\nZeilenbrecher",
                 "Zwei\nZeilen\necher"))
})
