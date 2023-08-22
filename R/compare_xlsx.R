#' Hashes the contents of xlsx files
#'
#' @description
#' `r lifecycle::badge("experimental")`
#'
#' @param tmpfile the file to be hashed
#' @return A list of hashed values for certain xml files
#' @importFrom rlang hash
compare_xlsx <- function(tmpfile) {

  lapply(c("xl/workbook.xml",
           "xl/styles.xml",
           "xl/worksheets/sheet1.xml"),
         \(x) {
           c <- unz(tmpfile, x)
           x <- rlang::hash(readChar(c, nchars = 90000))
           close(c)
           x
         })
}
