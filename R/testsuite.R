test_wb_ft <- function(wb, ft, filename) {
  # For local development testing only
  if(exists("testsuite5645454652165") &&
     testsuite5645454652165 == T) {
    wb$save(paste0("testsuite/",
                   filename,
                   ".xlsx"))
    flextable::save_as_html(ft, path = paste0("testsuite/",
                                              filename,
                                              ".html"))
  }
}
