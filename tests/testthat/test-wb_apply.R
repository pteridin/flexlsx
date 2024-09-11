test_that(
  "flextable without header", {

    sheet <- "iris"
    ft <- datasets::iris |>
      head() |>
      flextable::flextable() |>
      flextable::delete_part(part = "header")
    wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)
    dims <- "B2"


    wb <- wb_add_flextable(wb = wb,
                       ft = ft,
                       sheet = sheet,
                       dims = dims)

    df <- openxlsx2::wb_read(wb,
                       sheet=sheet,
                       start_row = 2,
                       start_col = 2,
                       col_names = F)
    df$F <- NULL
    df2 <- datasets::iris |>
      head()
    df2$Species <- NULL


    expect_equal(as.numeric(unlist(df)),
                 as.numeric(unlist(df2)))

    NULL
  }
)
