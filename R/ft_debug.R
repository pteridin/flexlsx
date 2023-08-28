#' Debugs the flextable creation
#'
#' @param ft the flextable
#'
#' @return NULL
#'
ft_debug <- function(ft) {
  tmpfile <- tempfile(fileext = ".xlsx")
  wb <- openxlsx2::wb_workbook()$add_worksheet("flextable")

  # add the flextable
  wb <- wb_add_flextable(wb, "flextable", ft)

  # add the styling
  df_styles <- ft_to_style_tibble(ft)
  wb$add_worksheet("df_styles")
  wb <- wb_add_flextable(wb, "flextable", flextable::flextable(dplyr::select(df_styles, -all_of("content"))))

  wb$add_worksheet("content")
  wb <- wb_add_flextable(wb, "content", flextable::flextable(dplyr::select(df_styles, all_of("content")) |>
                                                               tidyr::unnest_legacy()))

  # apply styling



}
