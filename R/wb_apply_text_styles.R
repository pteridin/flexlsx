
#' Applies the text styles
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb the [workbook][openxlsx2::wbWorkbook]
#' @param sheet the sheet of the workbook
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @importFrom dplyr select all_of
#' @importFrom openxlsx2 wb_color
#'
wb_apply_text_styles <- function(wb, sheet, df_style) {

  wb$validate_sheet(sheet)

  ## aggregate borders
  df_text_styles <- df_style |>
    dplyr::select(dplyr::all_of(c("col_id",
                                  "row_id",
                                  "font.family",
                                  "color",
                                  "font.size",
                                  "bold",
                                  "italic",
                                  "underlined")))

  df_text_styles_aggregated <- get_dim_ranges(df_text_styles)

  for(i in seq_len(nrow(df_text_styles_aggregated))) {
    crow <- df_text_styles_aggregated[i, ]

    wb$add_font(
      dims      = crow$dims,
      name      = crow$font.family,
      color     = openxlsx2::wb_color(crow$color),
      size      = crow$font.size,
      bold      = ifelse(crow$bold, "1", ""),
      italic    = ifelse(crow$italic, "1", ""),
      underline = ifelse(crow$underlined, "1", "")
    )
  }
  return(invisible(NULL))
}
