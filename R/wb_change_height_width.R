#' Changes the cell width
#'
#' @inheritParams wb_add_caption
#'
#' @return NULL
#'
wb_change_cell_width <- function(wb, sheet, ft, offset_cols) {
  # Tell me why?
  cwidths <- rbind(
    ft$header$colwidths,
    ft$body$colwidths,
    ft$footer$colwidths
  ) |>
    apply(2, max) * 2.54 * 4 / 16 * 20 # Ain't nothing but a constant


  wb$set_col_widths(
    sheet = sheet,
    cols = paste0(
      openxlsx2::int2col(1 + offset_cols), ":",
      openxlsx2::int2col(length(cwidths) + offset_cols)
    ),
    widths = cwidths
  )

  return(invisible(NULL))
}

#' Changes the row height
#'
#' @inheritParams wb_add_caption
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @return NULL
#'
#' @importFrom dplyr select mutate group_by all_of summarize
#' @importFrom rlang .data
#' @importFrom stringi stri_count
#'
wb_change_row_height <- function(wb, sheet, df_style) {
  font_sizes <- vapply(df_style$content,
    \(x) {
      ifelse(all(is.na(x$font.size)),
        NA_real_,
        max(x$font.size, na.rm = TRUE)
      )
    },
    FUN.VALUE = numeric(1)
  )

  newline_counts <- vapply(df_style$content,
    \(x) {
      sum(stringi::stri_count(x$txt, regex = "<br */{0,1}>") +
        stringi::stri_count(x$txt, regex = "\n"))
    },
    FUN.VALUE = numeric(1)
  ) + 1

  row_heights <- newline_counts *
    coalesce(font_sizes, df_style$font.size) / 11 * 15

  df_row_heights <- df_style |>
    dplyr::select(dplyr::all_of("row_id")) |>
    dplyr::mutate(rh = row_heights) |>
    dplyr::group_by(.data$row_id) |>
    dplyr::summarize(
      row_heights = max(.data$rh),
      .groups = "drop"
    )

  wb$set_row_heights(
    sheet = sheet,
    rows = df_row_heights$row_id,
    heights = df_row_heights$row_heights
  )
}
