

#' Converts a flextable-part to a tibble styles
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param ft_part the part of the flextable to extract the style from
#' @param part the name of the part
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @return a [tibble][tibble::tibble-package]
#'
#' @importFrom dplyr bind_cols select rename all_of arrange
#' @importFrom openxlsx2 int2col
#' @importFrom rlang .data
#'
ftpart_to_style_tibble <- function(ft_part,
                             part = c("header",
                                      "body",
                                      "footer")) {
  ## map styles to data.frames

  # Cells
  lapply(ft_part$styles$cells,
         \(x) {
           if("data" %in% names(x))
             return(as.vector(x$data))
           return(NULL)
         }) |>
    data.frame() -> df_styles_cells
  df_styles_cells$rowheight <- round(ft_part$rowheights * 91.4400, 0)

  # Pars
  lapply(ft_part$styles$pars,
         \(x) {
           if("data" %in% names(x))
             return(as.vector(x$data))
           return(NULL)
         }) |>
    data.frame() -> df_styles_pars

  # Text
  lapply(ft_part$styles$text,
         \(x) {
           if("data" %in% names(x))
             return(as.vector(x$data))
           return(NULL)
         }) |>
    data.frame() -> df_styles_text

  # Merge
  df_styles <- dplyr::bind_cols(df_styles_cells,
                                dplyr::rename(df_styles_text, "text.vertical.align" = dplyr::all_of("vertical.align")),
                                dplyr::select(df_styles_pars, dplyr::all_of("text.align")))

  # Determine spans
  df_styles$span.rows <- ft_part$spans$rows |> as.vector()
  df_styles$span.cols <- ft_part$spans$columns |> as.vector()

  # Add row and col id
  idims <- dim(ft_part$content$data)
  df_styles$col_id <- sort(rep(seq_len(idims[2]), idims[1]))
  df_styles$row_id <- rep(seq_len(idims[1]), idims[2])

  # Add content
  df_styles$content <- lapply(seq_len(nrow(df_styles)), function(i) {
    ft_part$content$data[[df_styles$row_id[i], df_styles$col_id[i]]]
  })

  # Arrange
  df_styles <- dplyr::arrange(df_styles, .data$row_id, .data$col_id)

  return(df_styles)
}

#' Converts a flextable to a tibble with style information
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param ft a [flextable][flextable::flextable-package]
#' @param offset_rows offsets the start-row
#' @param offset_cols offsets the start-columns
#' @param offset_caption_rows number of rows to offset the caption by
#'
#' @return a [tibble][tibble::tibble-package]
#'
#' @importFrom dplyr bind_rows
#' @importFrom openxlsx2 int2col
#'
ft_to_style_tibble <- function(ft, offset_rows = 0L, offset_cols = 0L, offset_caption_rows = 0L) {
  has_caption <- length(ft$caption$value) > 0
  has_footer <- length(ft$footer$content) > 0

  # Caption
  df_caption <- if(has_caption) tibble::tibble(row_id = 1, col_id = 1) else tibble::tibble()

  # Header
  df_header <- ftpart_to_style_tibble(ft$header)
  # Offset row-id based on caption rows
  if(has_caption)
    df_header$row_id <- df_header$row_id + max(df_caption$row_id)

  # Body
  df_body <- ftpart_to_style_tibble(ft$body)
  df_body$row_id <- df_body$row_id + max(df_header$row_id,0L)

  # Footer
  if(has_footer) {
    df_footer <- ftpart_to_style_tibble(ft$footer)
    df_footer$row_id <- df_footer$row_id + max(df_body$row_id)
  } else {
    df_footer <- tibble::tibble()
  }

  df_style <- dplyr::bind_rows(df_caption,
                                 df_header,
                                 df_body,
                                 df_footer)

  # offset the rows
  df_style$row_id <- df_style$row_id + offset_rows + offset_caption_rows
  df_style$col_id <- df_style$col_id + offset_cols

  df_style$col_name <- paste0(openxlsx2::int2col(df_style$col_id),
                            df_style$row_id)

  if(has_caption)
    df_style <- df_style[-1,]

  return(df_style)
}
