#' Determines the border style
#'
#' openxlsx2/Excel does handle borders differently than
#' flextable. This function maps the flextable border styles
#' to the Excel border styles.
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param border_color the color of the border
#' @param border_width a numeric vector determining the border-width
#' @param border_style the flextable style name of the border
#'
#' @return a factor of xlsx border styles
#'
#' @importFrom dplyr case_when
#'
ft_to_xlsx_border <- function(border_color,
                               border_width,
                               border_style) {
  dplyr::case_when(
    border_color == "transparent" |
      border_style %in% c("none", "nil") |
      border_width <= 0 ~ "no border",

    border_style == "double" ~ "double", # ?
    border_style == "dotted" ~ "dotted",

    border_style == "dashed" & border_width < 1.25 ~ "dashed",
    border_style == "dotDash" & border_width < 1.25 ~ "dashDot",

    border_style == "dashed" & border_width < 1.25 ~ "mediumDashed",
    border_style == "dotDash" & border_width < 1.25 ~ "mediumDashDot",

    border_style == "dashed" ~ "dashed",
    border_style == "dotDash" ~ "dashDot",

    border_style == "dotDotDash" ~ "dashedDotDot",

    border_width < .5 ~ "hair",
    border_width < 1 ~ "thin",
    border_width < 1.25 ~ "medium",
    T ~ "thick"
  )
}

#' Where there is no border return NULL
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param border_style the openxlsx2 style of the border
#'
#' @return border_style or NULL
#'
handle_null_border <- function(border_style) {
  if(border_style == "no border")
    return(NULL)
  return(border_style)
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

  wb$validate_sheet(sheet)

  ## aggregate borders
  df_borders <- df_style |>
    dplyr::select(dplyr::starts_with("border."),
                  dplyr::all_of(c("col_id",
                                  "row_id"))) |>
    dplyr::mutate(border.style.top = ft_to_xlsx_border(.data$border.color.top,
                                                       .data$border.width.top,
                                                       .data$border.style.top),
                  border.style.bottom = ft_to_xlsx_border(.data$border.color.bottom,
                                                       .data$border.width.bottom,
                                                       .data$border.style.bottom),
                  border.style.left = ft_to_xlsx_border(.data$border.color.left,
                                                       .data$border.width.left,
                                                       .data$border.style.left),
                  border.style.right = ft_to_xlsx_border(.data$border.color.right,
                                                       .data$border.width.right,
                                                       .data$border.style.right),
                  dplyr::across(dplyr::starts_with("border.color."),
                                ~ dplyr::if_else(.x == "transparent",
                                                 "black", .x)))


  df_borders_aggregated <- get_dim_ranges(df_borders)

  for(i in seq_len(nrow(df_borders_aggregated))) {
    crow <- df_borders_aggregated[i,]

    crow$border.style.top <- handle_null_border(crow$border.style.top)
    crow$border.style.bottom <- handle_null_border(crow$border.style.bottom)
    crow$border.style.left <- handle_null_border(crow$border.style.left)
    crow$border.style.right <- handle_null_border(crow$border.style.right)


    # Spans across multiple rows
    if(crow$multi_rows) {
      if(is.null(purrr::pluck(crow,"border.style.bottom"))) {
        hgrid_border <- purrr::pluck(crow,"border.style.top")
        hgrid_color <- openxlsx2::wb_color(crow$border.color.top)
      } else {
        hgrid_border <- purrr::pluck(crow,"border.style.bottom")
        hgrid_color <- openxlsx2::wb_color(crow$border.color.bottom)
      }
    }

    # Spans across multiple cols
    if(crow$multi_cols) {
      if(is.null(purrr::pluck(crow,"border.style.left"))) {
        vgrid_border <- purrr::pluck(crow,"border.style.right")
        vgrid_color <- openxlsx2::wb_color(crow$border.color.right)
      } else {
        vgrid_border <- purrr::pluck(crow,"border.style.left")
        vgrid_color <- openxlsx2::wb_color(crow$border.color.left)
      }
    }

    wb$add_border(
      sheet = sheet,
      dims = crow$dims,

      bottom_color = openxlsx2::wb_color(crow$border.color.bottom),
      left_color   = openxlsx2::wb_color(crow$border.color.left),
      right_color  = openxlsx2::wb_color(crow$border.color.right),
      top_color    = openxlsx2::wb_color(crow$border.color.top),

      bottom_border = purrr::pluck(crow,"border.style.bottom"),
      left_border   = purrr::pluck(crow,"border.style.left"),
      right_border  = purrr::pluck(crow,"border.style.right"),
      top_border    = purrr::pluck(crow,"border.style.top"),

      inner_hgrid = if(crow$multi_rows) hgrid_border else NULL,
      inner_hcolor = if(crow$multi_rows) hgrid_color else NULL,

      inner_vgrid = if(crow$multi_cols) vgrid_border else NULL,
      inner_vcolor = if(crow$multi_cols) vgrid_color else NULL
    )
  }
  return(invisible(NULL))
}
