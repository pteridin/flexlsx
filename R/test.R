#' Just a test, ignore it! :)
#'
#' @return a workbook
test_it <- function() {
  ft <- flextable::as_flextable(table(mtcars[,1:2]))

  ft_titanic <- tibble::as_tibble(Titanic)|>
    tidyr::uncount(n) |>
    gtsummary::tbl_strata(strata = Class,
                          .tbl_fun = ~gtsummary::tbl_summary(.x, by = Sex)) |>
    gtsummary::tbl_butcher() |>
    gtsummary::bold_labels() |>
    gtsummary::italicize_levels() |>
    gtsummary::as_flex_table()

  wb <- openxlsx2::wb_workbook()$
    add_worksheet("mtcars")$
    add_worksheet("titanic")

  wb <- wb_add_flextable(wb, "mtcars", ft)
  wb <- wb_add_flextable(wb, "titanic", ft_titanic)

  wb$save("~/text.xlsx")
}
