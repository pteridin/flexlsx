#' Retrieves dims of same style rows within same column
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param df_x styling information incl. col_id & row_id
#'
#' @importFrom dplyr select group_by summarize mutate
#' @importFrom dplyr arrange bind_cols first everything slice_min
#' @importFrom openxlsx2 int2col
#' @importFrom rlang .data
#'
#' @return merged styles as a [tibble][tibble::tibble-package]
get_dim_ranges <- function(df_x) {
  df_style_hashed <- df_x |>
    style_to_hash()

  df_rowwise <- df_x |>
    get_dim_rowwise(df_style_hashed)

  df_colwise <- df_rowwise |>
    get_dim_colwise()

  df_aggregated <- df_colwise |>
    mutate(
      dims = paste0(
        openxlsx2::int2col(.data$col_from),
        .data$row_from,
        ":",
        openxlsx2::int2col(.data$col_to),
        .data$row_to
      ),
      multi_lines = .data$row_to != .data$row_from |
        .data$col_to != .data$col_from
    ) |>
    left_join(df_style_hashed, by = "hash")

  return(df_aggregated)
}

#' Retrieves hashed style information
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' Converts each style to an individual integer hash
#' for easy comparison and aggregation.
#'
#' @inheritParams get_dim_ranges
#'
#' @return hashed style information as a [tibble][tibble::tibble-package]
#'
#' @importFrom dplyr arrange group_by summarize mutate all_of
#' @importFrom dplyr across row_number
#'
style_to_hash <- function(df_x) {
  df_style_hashed <- df_x |>
    arrange(across(all_of(c(
      "row_id", "col_id"
    )))) |>
    group_by(across(-all_of(c(
      "col_id", "row_id"
    )))) |>
    summarize(hash = 1L, .groups = "drop") |>
    mutate(hash = row_number())

  cols_to_join <- names(df_style_hashed)[names(df_style_hashed) != "hash"]
  attr(df_style_hashed, "cols_to_join") <- cols_to_join
  return(df_style_hashed)
}


#' Groups each column with same style each row
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @inheritParams get_dim_ranges
#' @param df_style_hashed [tibble][tibble::tibble-package] of hashed style
#' information
#'
#' @return [tibble][tibble::tibble-package] of row-wise aggregates style
#' information
#'
#' @importFrom dplyr left_join select arrange group_by
#' @importFrom dplyr mutate summarize first last
#' @importFrom dplyr all_of first last across lag
#' @importFrom rlang .data
#'
get_dim_rowwise <- function(df_x, df_style_hashed) {
  df_rows <- df_x |>
    left_join(df_style_hashed, by = attr(df_style_hashed, "cols_to_join")) |>
    select(all_of(c("row_id", "col_id", "hash"))) |>
    arrange(across(all_of(c(
      "row_id", "col_id"
    )))) |>
    group_by(across(all_of("row_id"))) |>
    mutate(col_change = cumsum(.data$hash !=
                                 lag(.data$hash,
                                     default = first(.data$hash)))) |>
    group_by(across(all_of(c(
      "row_id", "hash", "col_change"
    )))) |>
    summarize(
      col_from = min(.data$col_id),
      col_to   = max(.data$col_id),
      .groups  = "drop"
    ) |>
    select(-all_of("col_change")) |>
    arrange(across(all_of(c(
      "row_id", "col_from"
    ))))

  return(df_rows)
}


#' Groups each row with same style each column
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param df_rows [tibble][tibble::tibble-package] of row-wise aggregates style
#'
#' @return [tibble][tibble::tibble-package] of column-wise aggregates style
#'
#' @importFrom dplyr arrange group_by summarize mutate all_of
#' @importFrom dplyr across lag
#' @importFrom rlang .data
#'
get_dim_colwise <- function(df_rows) {
  df_rows |>
    arrange(across(all_of(c(
      "col_from", "col_to", "row_id"
    )))) |>
    group_by(across(all_of(c(
      "col_from", "col_to"
    )))) |>
    mutate(
      row_change = .data$row_id != lag(.data$row_id,
                                       default = min(.data$row_id)) + 1L,
      style_change = .data$hash != lag(.data$hash,
                                       default = min(.data$hash)),
      change = cumsum(.data$row_change | .data$style_change)
    ) |>
    group_by(across(all_of(
      c("hash", "col_from", "col_to", "change")
    ))) |>
    summarize(
      row_from = min(.data$row_id),
      row_to = max(.data$row_id),
      .groups = "drop"
    ) |>
    select(-all_of("change")) |>
    arrange(across(all_of(c(
      "row_from", "col_from"
    ))))
}
