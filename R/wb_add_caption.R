#' Adds a caption to an excel file
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @inheritParams wb_add_flextable
#' @param offset_rows zero-based row offset
#' @param offset_cols zero-based column offset
#'
#' @importFrom openxlsx2 fmt_txt
#' @importFrom purrr map_chr
#' @importFrom stringi stri_count
#'
#' @return NULL
#'
wb_add_caption <- function(wb, sheet,
                           ft,
                           offset_rows=offset_rows,
                           offset_cols=offset_cols) {
  idims <- dim(ft$body$content$data)

  # Default values from header
  lapply(ft$header$styles$text,
         \(x) {
           if("default" %in% names(x))
             return(as.vector(x$data))
           return(NULL)
         }) |>
    data.frame() -> df_styles_default
  df_styles_default <- df_styles_default[1,]

  # create content
  if(ft$caption$simple_caption) {
    # Replace <br> in flextables with newlines
    ft$caption$value <- gsub("<br *\\/{0,1}>", "\n", ft$caption$value)

    content <- openxlsx2::fmt_txt(
      ft$caption$value,
      bold = df_styles_default$bold,
      italic = df_styles_default$italic,
      underline = df_styles_default$underlined,
      size = df_styles_default$font.size,
      color = openxlsx2::wb_color(df_styles_default$color),
      font = df_styles_default$font.family,
      vert_align = df_styles_default$vertical.align
    )
  } else {
    ft$caption$value$txt <- gsub("<br *\\/{0,1}>", "\n", ft$caption$value$txt)
    content <- purrr::map_chr(1:nrow(ft$caption$value),
                              \(i) {
                                ft$caption$value[i,] -> x
                                openxlsx2::fmt_txt(
                                  x$txt,
                                  bold = x$bold,
                                  italic = x$italic,
                                  underline = x$underlined,
                                  size = x$font.size,
                                  color = openxlsx2::wb_color(x$color),
                                  font = x$font.family,
                                  vert_align = x$vertical.align
                                ) |>
                                  paste0()
                              })
  }

  # wrap text if necessary
  to_apply_text_wrap <- ifelse(ft$caption$simple_caption,
                               stringi::stri_count(ft$caption$value, regex = "\\n"),
                               sum(stringi::stri_count(ft$caption$value$txt, regex = "\\n"))) + 1
  if(to_apply_text_wrap > 0) {
    wb$add_cell_style(sheet = sheet,
                      wrap_text = T,
                      dims = paste0(int2col(offset_cols + 1),
                                    offset_rows + 1))

    wb$set_row_heights(sheet = sheet,
                       heights = to_apply_text_wrap*15,
                       rows = offset_rows + 1)
  }

  # add to wb & merge
  wb$add_data(sheet = sheet,
              x = paste0(content, collapse = ""),
              dims = paste0(int2col(offset_cols + 1),
                            offset_rows + 1))
  wb$merge_cells(sheet = sheet, dims = paste0(int2col(offset_cols + 1),
                                              offset_rows + 1,
                                              ":",
                                              int2col(offset_cols + idims[2]),
                                              offset_rows + 1))

  return(invisible(NULL))
}
