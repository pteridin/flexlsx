test_wb_ft <- function(wb, ft, filename) {
  test_path <- Sys.getenv("flexlsxtestdir")

  # For local development testing only
  if (test_path != "") {
    wb$save(paste0(
      test_path,
      filename,
      ".xlsx"
    ))
    flextable::save_as_html(ft, path = paste0(
      test_path,
      filename,
      ".html"
    ))
  }
}
