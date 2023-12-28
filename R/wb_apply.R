
#' Bugfix for illegal XML chars
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' Only applies for openxlsx versions <= 0.8.0.9000! Otherwise return
#' x
#'
#' Adopted from `openxlsx2`
#'
#' For Bug reports, see:
#' - [Bugreport #2](https://github.com/pteridin/flexlsx/issues/2)
#' - [Upstream](https://github.com/JanMarvin/openxlsx2/pull/757)
#'
#' @param x content with potentially "legal" XML chars
#'
#' @importFrom stringi stri_replace_all_fixed stri_conv
#' @importFrom utils packageVersion
#'
#' @return a character vector
#'
bugfix_openxlsx2_fmt_text <- function(x) {

  if(utils::packageVersion("openxlsx2") <= "0.8.0.9000") {
    x <- as.character(x)
    bad <- Encoding(x) != "UTF-8"

    if (any(bad)) {
      x[bad] <- stringi::stri_conv(x[bad], from = "", to = "UTF-8")
    }

    legal_chars <- c("&",  "\"", "'",  "<",  ">",  "\a", "\b", "\v", "\f")
    legal_sub <- c("&amp;",  "&quot;", "&apos;", "&lt;",   "&gt;",   "",       "",       "",       ""      )

    return(stringi::stri_replace_all_fixed(x, legal_chars, legal_sub, vectorize_all = FALSE))
  }

  return(x)
}


#' Determines the border width
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param border_width a numeric vector determining the border-width
#'
#' @return a factor of xlsx border styles
#'
ft_to_xlsx_border_width <- function(border_width) {
  cut(border_width,
      c(-Inf, 0, .9999, 1.25, Inf),
      c("no border",
        "hair",
        "medium",
        "thick"))  |>
    as.character()
}

#' Where there is no border return NULL
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param border_width a numeric vector determining the border-width
#'
#' @return border_width or NULL
handle_null_border <- function(border_width) {
  if(border_width == "no border")
    return(NULL)
  return(border_width)
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

  if(!sheet %in% wb$get_sheet_names())
    stop("sheet '", sheet, "' does not exist in wb!")

  ## aggregate borders
  df_borders <- df_style |>
    dplyr::select(dplyr::starts_with("border."),
                  dplyr::all_of(c("col_id",
                                  "row_id"))) |>
    dplyr::mutate(dplyr::across(dplyr::starts_with("border.width"),
                                ~ ft_to_xlsx_border_width(.x)),
                  border.width.top = dplyr::if_else(.data$border.color.top == "transparent",
                                                    "no border",
                                                    .data$border.width.top),
                  border.width.bottom = dplyr::if_else(.data$border.color.bottom == "transparent",
                                                       "no border",
                                                       .data$border.width.bottom),
                  border.width.left = dplyr::if_else(.data$border.color.left == "transparent",
                                                       "no border",
                                                     .data$border.width.left),
                  border.width.right = dplyr::if_else(.data$border.color.right == "transparent",
                                                       "no border",
                                                      .data$border.width.right),
                  dplyr::across(dplyr::starts_with("border.color."),
                                ~ dplyr::if_else(.x == "transparent",
                                                 "black", .x)))


  df_borders_aggregated <- get_dim_ranges(df_borders)

  for(i in seq_len(nrow(df_borders_aggregated))) {
    crow <- df_borders_aggregated[i,]

    crow$border.width.bottom <- handle_null_border(crow$border.width.bottom)
    crow$border.width.left <- handle_null_border(crow$border.width.left)
    crow$border.width.right <- handle_null_border(crow$border.width.right)
    crow$border.width.top <- handle_null_border(crow$border.width.top)

    if(crow$multi_lines) {
      if(is.null(purrr::pluck(crow,"border.width.bottom"))) {
        hgrid_border <- purrr::pluck(crow,"border.width.top")
        hgrid_color <- openxlsx2::wb_color(crow$border.color.top)
      } else {
        hgrid_border <- purrr::pluck(crow,"border.width.bottom")
        hgrid_color <- openxlsx2::wb_color(crow$border.color.bottom)
      }
    }

    wb$add_border(
      sheet = sheet,
      dims = crow$dims,

      bottom_color = openxlsx2::wb_color(crow$border.color.bottom),
      left_color   = openxlsx2::wb_color(crow$border.color.bottom),
      right_color  = openxlsx2::wb_color(crow$border.color.bottom),
      top_color    = openxlsx2::wb_color(crow$border.color.bottom),

      bottom_border = purrr::pluck(crow,"border.width.bottom"),
      left_border   = purrr::pluck(crow,"border.width.left"),
      right_border  = purrr::pluck(crow,"border.width.right"),
      top_border    = purrr::pluck(crow,"border.width.top"),

      inner_hgrid = if(crow$multi_lines) hgrid_border else NULL,
      inner_hcolor = if(crow$multi_lines) hgrid_color else NULL
    )
  }
  return(invisible(NULL))
}

#' Applies the text styles
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb the [workbook][openxlsx2::wbWorkbook]
#' @param sheet the sheet of the workbook
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @importFrom dplyr select all_of
#' @importFrom openxlsx2 wb_color
#'
wb_apply_text_styles <- function(wb, sheet, df_style) {

  if(!sheet %in% wb$get_sheet_names())
    stop("sheet '", sheet, "' does not exist in wb!")

  ## aggregate borders
  df_text_styles <- df_style |>
    dplyr::select(dplyr::all_of(c("col_id",
                                  "row_id",
                                  "font.family",
                                  "color",
                                  "font.size",
                                  "bold",
                                  "italic",
                                  "underlined")))

  df_text_styles_aggregated <- get_dim_ranges(df_text_styles)

  for(i in seq_len(nrow(df_text_styles_aggregated))) {
    crow <- df_text_styles_aggregated[i, ]

    wb$add_font(
      dims      = crow$dims,
      name      = crow$font.family,
      color     = openxlsx2::wb_color(crow$color),
      size      = crow$font.size,
      bold      = ifelse(crow$bold, "1", ""),
      italic    = ifelse(crow$italic, "1", ""),
      underline = ifelse(crow$underlined, "1", "")
    )
  }
  return(invisible(NULL))
}

#' Applies the cell styles
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb the [workbook][openxlsx2::wbWorkbook]
#' @param sheet the sheet of the workbook
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @importFrom dplyr select all_of mutate
#' @importFrom openxlsx2 wb_color
#' @importFrom rlang .data
#'
wb_apply_cell_styles <- function(wb, sheet, df_style) {

  if(!sheet %in% wb$get_sheet_names())
    stop("sheet '", sheet, "' does not exist in wb!")

  ## aggregate borders
  df_cell_styles <- df_style |>
    dplyr::mutate(background.color = ifelse(.data$shading.color != "transparent",
                                            .data$shading.color,
                                            .data$background.color),
                  text.direction = dplyr::case_when(.data$text.direction == "tbrl" ~ "180",
                                                    .data$text.direction == "btrl" ~ "90",
                                                    T ~ "")) |>
    dplyr::select(dplyr::all_of(c("col_id",
                                  "row_id",
                                  "text.align",
                                  "vertical.align",
                                  "text.direction",
                                  "background.color")))

  df_cell_styles_aggregated <- get_dim_ranges(df_cell_styles)

  for(i in seq_len(nrow(df_cell_styles_aggregated))) {
    crow <- df_cell_styles_aggregated[i, ]

    wb$add_cell_style(
      sheet = sheet,
      dims       = crow$dims,
      horizontal = crow$text.align,
      vertical   = crow$vertical.align,
      textRotation = crow$text.direction,
      wrapText   = "1"
    )

    if(crow$background.color != "transparent")
      wb$add_fill(
        sheet = sheet,
        dims  = crow$dims,
        color = openxlsx2::wb_color(crow$background.color)
      )

  }
  return(invisible(NULL))
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


#' Prepares the color for content style
#'
#' Converts a color name to the hexadecimal RGB-value
#' Removes "transparent" color
#'
#' @param color_name The name of the color
#'
#' @return The hexadecimal RGB-value
#'
#' @importFrom grDevices col2rgb rgb
#' @importFrom dplyr if_else
#'
prepare_color <- function(color_name) {

  color_name <- dplyr::if_else(color_name == "transparent",
                               NA_character_,
                               color_name)

  colors <- grDevices::col2rgb(color_name)/255
  colors <-  grDevices::rgb(red   = colors[1,],
                              green = colors[2,],
                              blue  = colors[3,])
  colors[is.na(color_name)] <- NA_character_
  return(colors)
}

#' Applies the content
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param wb the [workbook][openxlsx2::wbWorkbook]
#' @param sheet the sheet of the workbook
#' @param df_style the styling tibble from [ft_to_style_tibble]
#'
#' @importFrom dplyr select all_of mutate filter coalesce
#' @importFrom dplyr group_by summarize arrange left_join
#' @importFrom dplyr rowwise
#' @importFrom openxlsx2 wb_color
#' @importFrom rlang .data
#' @importFrom tidyr unnest_legacy
#'
wb_apply_content <- function(wb, sheet, df_style) {

  if(!sheet %in% wb$get_sheet_names())
    stop("sheet '", sheet, "' does not exist in wb!")

  df_content <- dplyr::select(df_style,
                              dplyr::all_of(c("row_id",
                                              "col_id",
                                              "span.rows",
                                              "span.cols",
                                              "font.size",
                                              "font.family",
                                              "color",
                                              "italic",
                                              "bold",
                                              "underlined",
                                              "content",
                                              "vertical.align")))

  ## unnest the content
  df_content_rows <- dplyr::select(df_style,
                                   dplyr::all_of(c("row_id",
                                                   "col_id",
                                                   "content"))) |>
    tidyr::unnest_legacy()

  ## join to the "default" options & replace nas
  df_content <- dplyr::select(df_content, -all_of("content")) |>
    dplyr::left_join(df_content_rows,
                     by = c("row_id", "col_id"),
                     relationship = "one-to-many")

  df_content <- dplyr::mutate(df_content,

                              italic.y = dplyr::coalesce(.data$italic.y,
                                                         .data$italic.x),
                              bold.y = dplyr::coalesce(.data$bold.y,
                                                       .data$bold.x),
                              underlined.y = dplyr::coalesce(.data$underlined.y,
                                                             .data$underlined.x),

                              # colors, font-size, font-family & vertical align will only be applied when different from the default
                              dplyr::across(dplyr::all_of(c("color.x","color.y")),
                                     ~ prepare_color(.x)),

                              color.y = dplyr::coalesce(.data$color.y, .data$color.x),
                              color.y = dplyr::if_else(.data$color.y == "#000000" & .data$color.x == "#000000",
                                                        NA_character_,
                                                        .data$color.y))





  # Replace <br> in flextables with newlines
  df_content$txt <- gsub("<br *\\/{0,1}>", "\n", df_content$txt)

  df_content |>
    mutate(txt = bugfix_openxlsx2_fmt_text(.data$txt)) |>
    dplyr::rowwise() |>
    dplyr::mutate(txt = paste0(openxlsx2::fmt_txt(
                                    .data$txt,
                                    bold = .data$bold.y,
                                    italic = .data$italic.y,
                                    underline = .data$underlined.y,
                                    size = if(is.na(.data$font.size.y)) NULL else .data$font.size.y[[1]],
                                    color = if(is.na(.data$color.y)) NULL else openxlsx2::wb_color(.data$color.y[[1]]),
                                    font = if(is.na(.data$font.family.y)) NULL else .data$font.family.y[[1]],
                                    vert_align = if(is.na(.data$vertical.align.y)) NULL else .data$vertical.align.y[[1]]
                                    ))) |>
    dplyr::ungroup() |>
    dplyr::mutate(txt = ifelse(.data$span.rows == 0 | .data$span.cols == 0,
                               "", .data$txt)) |>
    dplyr::group_by(.data$col_id,.data$row_id) |>
    dplyr::summarize(txt = paste0(.data$txt, collapse = ""),
                     max_font_size = max(coalesce(.data$font.size.y, .data$font.size.x),
                                         na.rm=T),
                     .groups = "drop")  -> df_content

  min_col_id <- min(df_content$col_id)
  max_col_id <- max(df_content$col_id)
  min_row_id <- min(df_content$row_id)
  max_row_id <- max(df_content$row_id)

  dims <- paste0(openxlsx2::int2col(min_col_id),
                 min_row_id, ":",
                 openxlsx2::int2col(max_col_id),
                 max_row_id)

  wb$add_data(sheet = sheet,
              x = matrix(df_content$txt,
                         nrow = max_row_id - min_row_id + 1,
                         ncol = max_col_id - min_col_id + 1),
              dims = dims,
              col_names = F)

  wb$add_ignore_error(dims = dims, number_stored_as_text = TRUE)

  return(invisible(NULL))
}



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
  idims <- dim(ft$body$content$content$data)

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


#' Changes the cell width
#'
#' @inheritParams wb_add_caption
#'
#' @return NULL
#'
wb_change_cell_width <- function(wb, sheet, ft, offset_cols) {

  # Tell me why?
  cwidths <- rbind(ft$header$colwidths,
                    ft$body$colwidths,
                    ft$footer$colwidths) |>
                apply(2,max) * 2.54 * 4 / 16 * 20 # Ain't nothing but a constant


  wb$set_col_widths(sheet = sheet,
                    cols = paste0(openxlsx2::int2col(1 + offset_cols), ":",
                                  openxlsx2::int2col(length(cwidths) + offset_cols)),
                    widths = cwidths)

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
                       \(x) ifelse(all(is.na(x$font.size)),
                                   NA_real_,
                                   max(x$font.size, na.rm=T)),
                       FUN.VALUE = numeric(1))

  newline_counts <- vapply(df_style$content,
                           \(x) sum(stringi::stri_count(x$txt, regex = "<br */{0,1}>") +
                                      stringi::stri_count(x$txt, regex = "\n")),
                           FUN.VALUE = numeric(1)) + 1

  row_heights <- newline_counts*coalesce(font_sizes,df_style$font.size)/11*15

  df_row_heights <- df_style |>
    dplyr::select(dplyr::all_of("row_id")) |>
    dplyr::mutate(rh = row_heights) |>
    dplyr::group_by(.data$row_id) |>
    dplyr::summarize(row_heights = max(.data$rh),
              .groups = "drop")

  wb$set_row_heights(sheet = sheet,
                     rows = df_row_heights$row_id,
                     heights = df_row_heights$row_heights)
}


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
#' if(requireNamespace("flextable", quietly = TRUE)) {
#'    ft <- flextable::as_flextable(table(mtcars[,c("am","cyl")]))
#'    wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")
#'    wb_add_flextable(wb, "mtcars", ft)$save("~/text.xlsx")
#' }
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



