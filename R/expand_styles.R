

#' Retrieves dims of same style rows within same column
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param df_x styling information incl. col_id & row_id
#'
#' @importFrom dplyr select rowwise do ungroup mutate
#' @importFrom dplyr arrange bind_cols first
#' @importFrom rlang hash .data
#' @importFrom openxlsx2 int2col
#'
#' @return merged styles as a [tibble][tibble::tibble-package]
get_dim_ranges <- function(df_x) {

  ## Hash the style information -------
  df_style_hashed <- df_x |>
    dplyr::select(-dplyr::all_of(c("col_id",
                                    "row_id"))) |>
    dplyr::rowwise() |>
    dplyr::do(hash = rlang::hash(.)) |>
    dplyr::ungroup() |>
    tidyr::unnest_legacy() |>
    dplyr::mutate(hash = as.integer(factor(.data$hash)))

  ## Based on style merge rows -------
  df_aggregated <- df_x |>
    dplyr::bind_cols(df_style_hashed) |>
    dplyr::arrange(.data$col_id, .data$row_id) |>
    dplyr::group_by(.data$col_id) |>
    dplyr::mutate(line = cumsum(dplyr::coalesce(dplyr::lag(.data$hash),0L) - .data$hash != 0)) |>
    dplyr::group_by(.data$col_id, .data$line) |>
    dplyr::mutate(row_to = max(.data$row_id)) |>
    dplyr::do(style = dplyr::first(dplyr::select(., -dplyr::all_of(c("col_id", "line"))))) |>
    dplyr::ungroup() |>
    tidyr::unnest_legacy()

  ## Prepare ranges --------
  df_aggregated <- df_aggregated |>
    dplyr::mutate(dims = paste0(openxlsx2::int2col(.data$col_id), .data$row_id, ":",
                                openxlsx2::int2col(.data$col_id), .data$row_to))

  df_aggregated$multi_lines <- df_aggregated$row_id != df_aggregated$row_to

  return(df_aggregated)
}
