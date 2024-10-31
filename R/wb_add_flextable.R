#' Adds a flextable to an openxlsx2 workbook sheet
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb an openxlsx2 workbook
#' @param sheet an openxlsx2 workbook sheet
#' @param ft a flextable
#' @param start_col a vector specifying the starting column to write to.
#' @param start_row a vector specifying the starting row to write to.
#' @param dims Spreadsheet dimensions that will determine start_col and start_row: "A1", "A1:B2", "A:B"
#' @param offset_caption_rows number of rows to offset the caption by
#'
#' @return an openxlsx2 workbook
#' @export
#'
#' @importFrom openxlsx2 dims_to_rowcol
#'
#' @examples
#'
#' if(requireNamespace("flextable", quietly = TRUE)) {
#'   # Create a flextable
#'   ft <- flextable::as_flextable(table(mtcars[,c("am","cyl")]))
#'
#'   # Create a workbook
#'   wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
#'
#'   # Add flextable to workbook
#'   wb <- wb_add_flextable(wb, "mtcars", ft)
#'
#'   # Workbook can now be saved wb$save(),
#'   # opened wb$open() - or removed
#'   rm(wb)
#'  }
#'
wb_add_flextable <- function(wb, sheet, ft,
                             start_col = 1,
                             start_row = 1,
                             offset_caption_rows = 0L,
                             dims = NULL) {
  # Check input
  stopifnot("wbWorkbook" %in% class(wb))
  stopifnot((is.character(sheet) &&
              nchar(sheet) > 0) ||
              is.numeric(sheet) &&
              sheet == as.integer(sheet))
  stopifnot("flextable" %in% class(ft))

  # Retrieve offsets
  if (!is.null(dims)) {
    dims <- openxlsx2::dims_to_rowcol(dims, as_integer = TRUE)
    offset_cols <- min(dims[[1]]) - 1
    offset_rows <- min(dims[[2]]) - 1
  } else {
    stopifnot(is.numeric(start_col),
              start_col >= 1,
              as.integer(start_col) == start_col,
              length(start_col) == 1)
    stopifnot(is.numeric(start_row) &&
                start_row >= 1 &&
                as.integer(start_col) == start_col,
              length(start_col) == 1)

    offset_cols <- start_col - 1
    offset_rows <- start_row - 1
  }

  # ignore offset if there is no caption
  if(length(ft$caption$value) == 0) {
    offset_caption_rows <- 0L
  }

  wb <- wb$clone()

  df_style <- ft_to_style_tibble(ft,
                                 offset_rows=offset_rows,
                                 offset_cols=offset_cols,
                                 offset_caption_rows=offset_caption_rows)

  # Apply styles & add content
  if(length(ft$caption$value) > 0)
    wb_add_caption(wb, sheet = sheet, ft = ft,
                   offset_rows=offset_rows,
                   offset_cols=offset_cols)

  df_style <- wb_apply_merge(wb, sheet, df_style)
  wb_apply_border(wb, sheet, df_style)
  wb_apply_text_styles(wb, sheet, df_style)
  wb_apply_cell_styles(wb, sheet, df_style)
  wb_apply_content(wb, sheet, df_style)
  wb_change_cell_width(wb, sheet, ft, offset_cols)
  wb_change_row_height(wb, sheet, df_style)

  return(wb)
}



