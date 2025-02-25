#' Determine problematic merges
#'
#' @param df_to_merge The data.frame containing information about the cells to
#' merge
#'
#' @return df_to_merge is extended by is_encapsulated and is_need_resolve
#'
#' @importFrom rlang .data
#' @importFrom dplyr mutate arrange
#'
merge_resolve_type <- function(df_to_merge) {
  n_x <- nrow(df_to_merge)
  df_to_merge <- df_to_merge |>
    dplyr::mutate(merge_type = dplyr::case_when(
      .data$span.rows > 0 &
        .data$span.cols > 0 ~ 1L,
      .data$span.rows > 0 ~ 2L,
      TRUE ~ 3L
    )) |>
    dplyr::arrange(
      .data$merge_type,
      .data$row_id,
      .data$col_id
    )

  df_to_merge <- cpp_merge_resolve_type(df_to_merge)

  return(df_to_merge)
}

#' Merges cells
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @inheritParams wb_apply_border
#'
#' @return df_style tibble
#'
#' @importFrom dplyr select all_of mutate filter
#' @importFrom openxlsx2 wb_color
#' @importFrom rlang .data
#'
wb_apply_merge <- function(wb, sheet, df_style) {
  df_merges <- df_style |>
    dplyr::mutate(
      span.rows = pmax(.data$span.rows - 1, 0),
      span.cols = pmax(.data$span.cols - 1, 0)
    ) |>
    dplyr::filter(.data$span.rows > 0 |
      .data$span.cols > 0) |>
    dplyr::mutate(
      row_end = .data$row_id + .data$span.cols,
      col_end = .data$col_id + .data$span.rows,
      dims = paste0(
        openxlsx2::int2col(.data$col_id), .data$row_id, ":",
        openxlsx2::int2col(.data$col_end), .data$row_end
      )
    ) |>
    dplyr::select(dplyr::all_of(c(
      "span.rows",
      "span.cols",
      "row_id",
      "row_end",
      "col_id",
      "col_end",
      "dims"
    ))) |>
    merge_resolve_type() |>
    dplyr::filter(!.data$is_encapsulated)

  if (sum(df_merges$is_need_resolve) > 0) {
    warning("Found ", sum(df_merges$is_need_resolve), " overlapping merges!
  Conflicting merges are removed;
  Styling might not fully resemble the flextable!")
    df_merges <- df_merges |>
      dplyr::filter(!.data$is_need_resolve)
  }

  ## Apply merges
  for (i in seq_len(nrow(df_merges))) {
    df_style_def <- df_merges[i, ]
    wb$merge_cells(
      sheet = sheet,
      dims = df_style_def$dims,
      solve = df_style_def$is_need_resolve
    )
  }

  return(df_style)
}
