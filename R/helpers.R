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
