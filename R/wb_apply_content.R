
apply_sub_or_parent <- function(sub, parent, .fn = identity) {
  if(is.na(sub) &&
      is.na(parent))
    return(NULL)

  if (!is.na(sub))
    return(.fn(sub))
  return(.fn(parent))
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
  wb$validate_sheet(sheet)

  df_content <- dplyr::select(
    df_style,
    dplyr::all_of(c(
      "row_id",
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
      "vertical.align"
    ))
  )

  ## unnest the content
  df_content_rows <- dplyr::select(
    df_style,
    dplyr::all_of(c(
      "row_id",
      "col_id",
      "content"
    ))
  ) |>
    tidyr::unnest_legacy()

  ## join to the "default" options & replace nas
  df_content <- dplyr::select(df_content, -all_of("content")) |>
    dplyr::left_join(df_content_rows,
      by = c("row_id", "col_id"),
      relationship = "one-to-many"
    )

  df_content <- dplyr::mutate(df_content,
    italic.y = dplyr::coalesce(
      .data$italic.y,
      .data$italic.x
    ),
    bold.y = dplyr::coalesce(
      .data$bold.y,
      .data$bold.x
    ),
    underlined.y = dplyr::coalesce(
      .data$underlined.y,
      .data$underlined.x
    ),

    # colors, font-size, font-family & vertical align will only be applied when
    # different from the default
    dplyr::across(
      dplyr::all_of(c("color.x", "color.y")),
      ~ prepare_color(.x)
    ),
    color.y = dplyr::coalesce(.data$color.y, .data$color.x),
    color.y = dplyr::if_else(.data$color.y == "#000000" &
      .data$color.x == "#000000",
    NA_character_,
    .data$color.y
    )
  )

  # Replace <br> in flextables with newlines
  df_content$txt <- gsub("<br *\\/{0,1}>", "\n", df_content$txt)

  df_content <- df_content |>
    dplyr::rowwise() |>
    dplyr::mutate(txt = paste0(openxlsx2::fmt_txt(
      .data$txt,
      bold = apply_sub_or_parent(.data$bold.y,
                                 .data$bold.x),
      italic = apply_sub_or_parent(.data$italic.y,
                                   .data$italic.x),
      underline = apply_sub_or_parent(.data$underlined.y,
                                      .data$underlined.x),
      size = apply_sub_or_parent(.data$font.size.y,
                                 .data$font.size.x),
      color = apply_sub_or_parent(.data$color.y,
                                  .data$color.x,
                                  .fn = openxlsx2::wb_color),
      font = apply_sub_or_parent(.data$font.family.y,
                                  .data$font.family.x),
      vert_align = apply_sub_or_parent(.data$vertical.align.y,
                                       .data$vertical.align.x)
    ))) |>
    dplyr::ungroup() |>
    dplyr::mutate(txt = ifelse(.data$span.rows == 0 | .data$span.cols == 0,
      "", .data$txt
    )) |>
    dplyr::group_by(.data$col_id, .data$row_id) |>
    dplyr::summarize(
      txt = paste0(.data$txt, collapse = ""),
      max_font_size = max(coalesce(.data$font.size.y, .data$font.size.x),
        na.rm = TRUE
      ),
      .groups = "drop"
    )

  min_col_id <- min(df_content$col_id)
  max_col_id <- max(df_content$col_id)
  min_row_id <- min(df_content$row_id)
  max_row_id <- max(df_content$row_id)

  dims <- paste0(
    openxlsx2::int2col(min_col_id),
    min_row_id, ":",
    openxlsx2::int2col(max_col_id),
    max_row_id
  )

  df <- matrix(df_content$txt,
    nrow = max_row_id - min_row_id + 1,
    ncol = max_col_id - min_col_id + 1
  ) |>
    as.data.frame()

  if (getOption("openxlsx2.string_nums", default = FALSE)) {
    # convert from styled character to numeric
    xml_to_num <- function(x) {
      val <- vapply(x,
        \(x) {
          ifelse(x == "", NA_character_,
            openxlsx2::xml_value(x, "r", "t")
          )
        },
        FUN.VALUE = character(1),
        USE.NAMES = FALSE
      )
      suppressWarnings(got <- as.numeric(val))
      sel <- !is.na(val) & !is.na(got)
      x[sel] <- got[sel]
      x
    }

    df[] <- lapply(df, xml_to_num)
  }

  wb$add_data(
    sheet = sheet,
    x = df,
    dims = dims,
    col_names = FALSE
  )

  wb$add_ignore_error(dims = dims, number_stored_as_text = TRUE)

  return(invisible(NULL))
}
