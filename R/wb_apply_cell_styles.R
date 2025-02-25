#' Applies the cell styles
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb the [workbook][openxlsx2::wbWorkbook]
#' @param sheet the sheet of the workbook
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @importFrom dplyr select all_of mutate
#' @importFrom openxlsx2 wb_color
#' @importFrom rlang .data
#'
wb_apply_cell_styles <- function(wb, sheet, df_style) {
  wb$validate_sheet(sheet)

  ## aggregate borders
  df_cell_styles <- df_style |>
    dplyr::mutate(
      background.color = ifelse(.data$shading.color != "transparent",
        .data$shading.color,
        .data$background.color
      ),
      text.direction = dplyr::case_when(
        .data$text.direction == "tbrl" ~ "180",
        .data$text.direction == "btrl" ~ "90",
        TRUE ~ ""
      )
    ) |>
    dplyr::select(dplyr::all_of(c(
      "col_id",
      "row_id",
      "text.align",
      "vertical.align",
      "text.direction",
      "background.color"
    )))

  df_cell_styles_aggregated <- get_dim_ranges(df_cell_styles)

  for (i in seq_len(nrow(df_cell_styles_aggregated))) {
    crow <- df_cell_styles_aggregated[i, ]

    wb$add_cell_style(
      sheet = sheet,
      dims = crow$dims,
      horizontal = crow$text.align,
      vertical = crow$vertical.align,
      textRotation = crow$text.direction,
      wrapText = "1"
    )

    if (crow$background.color != "transparent") {
      wb$add_fill(
        sheet = sheet,
        dims  = crow$dims,
        color = openxlsx2::wb_color(crow$background.color)
      )
    }
  }
  return(invisible(NULL))
}
