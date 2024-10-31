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

  ## cols (& rows merge)
  df_cols_to_merge <- df_style |>
    dplyr::filter(.data$span.rows > 1)  |>
    dplyr::select(dplyr::all_of(c("span.rows",
                                  "span.cols",
                                  "row_id",
                                  "col_id"))) |>
    dplyr::mutate(
      span.rows = pmax(.data$span.rows - 1, 0),
      span.cols = pmax(.data$span.cols - 1, 0),

      dims = paste0(
        openxlsx2::int2col(.data$col_id), .data$row_id, ":",
        openxlsx2::int2col(.data$col_id + .data$span.rows),
        .data$row_id + .data$span.cols
      ))

  for(i in seq_len(nrow(df_cols_to_merge))) {
    df_style_def <- df_cols_to_merge[i,]

    # Override style of merged columns with merging row
    df_style_info <- df_style[df_style$row_id == df_style_def$row_id &
                                df_style$col_id == df_style_def$col_id,] |>
      select(-all_of(c("row_id", "col_id", "content")))

    df_style[df_style$row_id >= df_style_def$row_id &
               df_style$row_id <= df_style_def$row_id + df_style_def$span.cols &
               df_style$col_id >= df_style_def$col_id &
               df_style$col_id <= df_style_def$col_id + df_style_def$span.rows ,
             -which(names(df_style) %in% c("row_id", "col_id", "content"))] <- df_style_info

    # Apply merge
    wb$merge_cells(sheet = sheet,
                   dims = df_style_def$dims)
  }


  ## rows merge only!
  df_rows_to_merge <- df_style |>
    dplyr::filter(.data$span.cols > 1,
                  .data$span.rows <= 1) |>
    dplyr::select(dplyr::all_of(c("span.rows",
                                  "span.cols",
                                  "row_id",
                                  "col_id"))) |>
    dplyr::mutate(
      span.rows = pmax(.data$span.rows - 1, 0),
      span.cols = pmax(.data$span.cols - 1, 0),

      dims = paste0(
        openxlsx2::int2col(.data$col_id), .data$row_id, ":",
        openxlsx2::int2col(.data$col_id + .data$span.rows),
        .data$row_id + .data$span.cols
      ))

  for(i in seq_len(nrow(df_rows_to_merge))) {
    df_style_def <- df_rows_to_merge[i,]

    # Override style of merged columns with merging row
    df_style_info <- df_style[df_style$row_id == df_style_def$row_id &
                                df_style$col_id == df_style_def$col_id,] |>
      select(-all_of(c("row_id", "col_id", "content")))

    df_style[df_style$row_id >= df_style_def$row_id &
               df_style$row_id <= df_style_def$row_id + df_style_def$span.cols &
               df_style$col_id >= df_style_def$col_id &
               df_style$col_id <= df_style_def$col_id + df_style_def$span.rows,
             -which(names(df_style) %in% c("row_id", "col_id", "content"))] <- df_style_info

    # Apply merge
    wb$merge_cells(sheet = sheet,
                   dims = df_style_def$dims)
  }

  return(df_style)
}
