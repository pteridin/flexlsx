

#' Retrieves dims of same style rows within same column
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param df_x styling information incl. col_id & row_id
#'
#' @importFrom dplyr select do ungroup mutate summarize group_by all_of
#' @importFrom dplyr arrange bind_cols first everything slice_min
#' @importFrom openxlsx2 int2col
#'
#' @return merged styles as a [tibble][tibble::tibble-package]
get_dim_ranges <- function(df_x) {

  ## Hash the style information -------
  df_style_hashed <- df_x |>
    dplyr::group_by(dplyr::across(-dplyr::all_of(c("col_id",
                                                   "row_id")))) |>
    dplyr::summarize(hash = 1L,
                     .groups = "drop") |>
    dplyr::mutate(hash = dplyr::row_number())

  cols_to_join <- names(df_style_hashed)[names(df_style_hashed) != "hash"]


  ## Aggregate based on style, merge rows -------
  df_aggregated <- df_x |>
    dplyr::left_join(df_style_hashed,
                     by = cols_to_join) |>
    dplyr::arrange(.data$col_id, .data$row_id) |>
    dplyr::group_by(.data$col_id) |>
    dplyr::mutate(line = cumsum(dplyr::coalesce(dplyr::lag(.data$hash),0L) - .data$hash != 0)) |>
    dplyr::group_by(.data$col_id, .data$line) |>
    dplyr::mutate(row_to = max(.data$row_id)) |>
    dplyr::slice_min(.data$row_id) |>
    dplyr::ungroup()

  ## Prepare ranges --------
  df_aggregated <- df_aggregated |>
    dplyr::mutate(dims = paste0(openxlsx2::int2col(.data$col_id), .data$row_id, ":",
                                openxlsx2::int2col(.data$col_id), .data$row_to)) |>
    dplyr::select(all_of(c("col_id", "line",   "row_id")),
                  dplyr::everything())

  df_aggregated$multi_lines <- df_aggregated$row_id != df_aggregated$row_to


  return(df_aggregated)
}
