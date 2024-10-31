#' Determines the border width
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param border_width a numeric vector determining the border-width
#'
#' @return a factor of xlsx border styles
#'
ft_to_xlsx_border_width <- function(border_width) {
  cut(border_width,
      c(-Inf, 0, .9999, 1.25, Inf),
      c("no border",
        "hair",
        "medium",
        "thick"))  |>
    as.character()
}

#' Where there is no border return NULL
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param border_width a numeric vector determining the border-width
#'
#' @return border_width or NULL
#'
handle_null_border <- function(border_width) {
  if(border_width == "no border")
    return(NULL)
  return(border_width)
}

#' Applies the border styles
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb the [workbook][openxlsx2::wbWorkbook]
#' @param sheet the sheet of the workbook
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @importFrom dplyr select mutate all_of starts_with across
#' @importFrom dplyr if_else
#' @importFrom purrr pluck
#' @importFrom openxlsx2 wb_color
#' @importFrom rlang .data
#'
wb_apply_border <- function(wb, sheet, df_style) {

  if(!sheet %in% wb$get_sheet_names())
    stop("sheet '", sheet, "' does not exist in wb!")

  ## aggregate borders
  df_borders <- df_style |>
    dplyr::select(dplyr::starts_with("border."),
                  dplyr::all_of(c("col_id",
                                  "row_id"))) |>
    dplyr::mutate(dplyr::across(dplyr::starts_with("border.width"),
                                ~ ft_to_xlsx_border_width(.x)),
                  border.width.top = dplyr::if_else(.data$border.color.top == "transparent",
                                                    "no border",
                                                    .data$border.width.top),
                  border.width.bottom = dplyr::if_else(.data$border.color.bottom == "transparent",
                                                       "no border",
                                                       .data$border.width.bottom),
                  border.width.left = dplyr::if_else(.data$border.color.left == "transparent",
                                                     "no border",
                                                     .data$border.width.left),
                  border.width.right = dplyr::if_else(.data$border.color.right == "transparent",
                                                      "no border",
                                                      .data$border.width.right),
                  dplyr::across(dplyr::starts_with("border.color."),
                                ~ dplyr::if_else(.x == "transparent",
                                                 "black", .x)))


  df_borders_aggregated <- get_dim_ranges(df_borders)

  for(i in seq_len(nrow(df_borders_aggregated))) {
    crow <- df_borders_aggregated[i,]

    crow$border.width.bottom <- handle_null_border(crow$border.width.bottom)
    crow$border.width.left <- handle_null_border(crow$border.width.left)
    crow$border.width.right <- handle_null_border(crow$border.width.right)
    crow$border.width.top <- handle_null_border(crow$border.width.top)

    if(crow$multi_lines) {
      if(is.null(purrr::pluck(crow,"border.width.bottom"))) {
        hgrid_border <- purrr::pluck(crow,"border.width.top")
        hgrid_color <- openxlsx2::wb_color(crow$border.color.top)
      } else {
        hgrid_border <- purrr::pluck(crow,"border.width.bottom")
        hgrid_color <- openxlsx2::wb_color(crow$border.color.bottom)
      }
    }

    wb$add_border(
      sheet = sheet,
      dims = crow$dims,

      bottom_color = openxlsx2::wb_color(crow$border.color.bottom),
      left_color   = openxlsx2::wb_color(crow$border.color.bottom),
      right_color  = openxlsx2::wb_color(crow$border.color.bottom),
      top_color    = openxlsx2::wb_color(crow$border.color.bottom),

      bottom_border = purrr::pluck(crow,"border.width.bottom"),
      left_border   = purrr::pluck(crow,"border.width.left"),
      right_border  = purrr::pluck(crow,"border.width.right"),
      top_border    = purrr::pluck(crow,"border.width.top"),

      inner_hgrid = if(crow$multi_lines) hgrid_border else NULL,
      inner_hcolor = if(crow$multi_lines) hgrid_color else NULL
    )
  }
  return(invisible(NULL))
}
