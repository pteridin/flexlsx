test_that(
  "flextable without header", {
    skip_if_not_installed("flextable")

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
    test_wb_ft(wb,ft, "without_header")

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

test_that("Add with numeric offset", {
  skip_if_not_installed("flextable")
  data("mtcars")

  sheet <- "iris"
  ft <- mtcars |>
    head() |>
    flextable::flextable()
  wb <- openxlsx2::wb_workbook()$add_worksheet(sheet)

  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = sheet,
                         start_col = 2,
                         start_row = 2)
  test_wb_ft(wb,ft, "numeric_offset")

  df <- openxlsx2::wb_read(wb,
                           sheet=sheet,
                           start_row = 2,
                           start_col = 2,
                           col_names = T)
  df2 <- mtcars |>
    head()
  rownames(df2) <- NULL


  expect_equal(as.numeric(unlist(df)),
               as.numeric(unlist(df2)))

  NULL
})

test_that("Add multi-header", {
  skip_if_not_installed("flextable")

  typology <- data.frame(
    col_keys = c(
      "Sepal.Length", "Sepal.Width", "Petal.Length",
      "Petal.Width", "Species"
    ),
    what = c("Sepal", "Sepal", "Petal", "Petal", "Species"),
    measure = c("Length", "Width", "Length", "Width", "Species"),
    stringsAsFactors = FALSE
  )

  ft_1 <- flextable::flextable(head(iris)) |>
    flextable::set_header_df(mapping = typology, key = "col_keys") |>
    flextable::merge_h(part = "header") |>
    flextable::merge_v( j = "Species", part = "header") |>
    flextable::theme_vanilla() |>
    flextable::fix_border_issues() |>
    flextable::autofit()

  wb <- openxlsx2::wb_workbook()$add_worksheet("multiheader")

  wb <- wb_add_flextable(wb = wb,
                         ft = ft_1,
                         sheet = "multiheader",
                         start_col = 2,
                         start_row = 2)
  test_wb_ft(wb,ft, "multi_header")

  expect_equal(openxlsx2::wb_read(wb,
                           sheet="multiheader",
                           start_row = 2,
                           start_col = 2,
                           col_names = T) |>
                colnames(),
               c("Sepal", NA, "Petal", NA, "Species"))

  expect_equal(openxlsx2::wb_read(wb,
                                  sheet="multiheader",
                                  start_row = 3,
                                  start_col = 2,
                                  col_names = T) |>
                 colnames(),
               c("Length", "Width", "Length", "Width", NA))

  NULL
})

test_that("using openxlsx2::current_sheet() works", {
  skip_if_not_installed("flextable")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))

  openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    flexlsx::wb_add_flextable(
      sheet = openxlsx2::current_sheet(),
      ft = ft,
      dims = "C2"
    ) -> wb

  expect_equal(
    wb$get_sheet_names(),
    c(`Sheet 1` = "Sheet 1")
  )

  openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    flexlsx::wb_add_flextable(
      ft = ft,
      dims = "A1"
    ) -> wb

  expect_equal(
    wb$get_sheet_names(),
    c(`Sheet 1` = "Sheet 1")
  )

  expect_equal(names(wb$to_df(sheet  = "Sheet 1")),
               c("mpg", NA, "cyl", NA, NA, NA))

})

test_that("When sheet does not exists throws an error", {
  skip_if_not_installed("flextable")

  ft <- flextable::as_flextable(table(mtcars[,1:2]))

  expect_error( openxlsx2::wb_workbook() |>
                  flexlsx::wb_add_flextable(
                    sheet = openxlsx2::current_sheet(),
                    ft = ft,
                    dims = "C2"
                  ),
                regexp = "Sheet 'current_sheet' does not exist!")

  expect_error( openxlsx2::wb_workbook() |>
                  flexlsx::wb_add_flextable(
                    sheet = "test",
                    ft = ft,
                    dims = "C2"
                  ),
                regexp = "Sheet 'test' does not exist!")
})

test_that("Complex FT", {
  skip_if_not_installed("officer")
  skip_if_not_installed("flextable")

  library(flextable)

  # -----------------------------------------------------------------------------
  # 1. Create a sample data frame.
  #    Each column name identifies what aspect is being tested.
  df <- data.frame(
    id         = c("Row1", "Row2", "Row3", "Row4"),
    chunk_test = c("Format", "Format", "Format", "Format"),
    para_test  = c("Paragraph", "Paragraph", "Paragraph", "Paragraph"),
    h1         = c("Merge", "A", "B", "C"),      # For horizontal merging test
    h2         = c("Merge", "X", "Y", "Z"),      # For horizontal merging test
    v1         = c("Unique", "MergeV", "MergeV", "Unique"),  # For vertical merging test
    append_test= c("Original", "Original", "Original", "Original"),
    stringsAsFactors = FALSE
  )

  # -----------------------------------------------------------------------------
  # 2. Create the flextable object from the data frame.
  ft <- flextable(df)

  # -----------------------------------------------------------------------------
  # 3. Set custom column widths (in inches) and row heights.
  ft <- width(ft, j = 1, width = 0.7)   # 'id'
  ft <- width(ft, j = 2, width = 2)     # 'chunk_test'
  ft <- width(ft, j = 3, width = 2)     # 'para_test'
  ft <- width(ft, j = 4, width = 1)     # 'h1'
  ft <- width(ft, j = 5, width = 1)     # 'h2'
  ft <- width(ft, j = 6, width = 1)     # 'v1'
  ft <- width(ft, j = 7, width = 2)     # 'append_test'

  ft <- height(ft, i = 1:4, height = 0.8, part = "body")

  # -----------------------------------------------------------------------------
  # 4. Add a header row (spanning all columns), a caption, and a footer.
  ft <- add_header_row(ft, values = c("Advanced Test Table"), colwidths = ncol(df))
  ft <- merge_h(ft, part = "header")  # Merge the header row into one cell
  ft <- set_caption(ft, caption = "Advanced Flextable Test: Chunks, Paragraphs, Merges & Borders")
  ft <- add_footer_lines(ft, values = "Footer: End of Advanced Test")

  # -----------------------------------------------------------------------------
  # 5. Use sugar functions to style chunks in the 'chunk_test' column.
  #    Compose a paragraph with multiple formatted chunks.
  ft <- compose(ft, j = "chunk_test", i = 1,
                value = as_paragraph(
                  "Normal text, ",
                  as_b("Bold text, "),
                  as_i("Italic text, "),
                  as_sub("Subscript, "),
                  as_sup("Superscript")
                ))
  # For rows 2-4, show a simple composition with inline formatting.
  for(i in 2:4){
    ft <- compose(ft, j = "chunk_test", i = i,
                  value = as_paragraph("Row ", i, ": ", as_b("Bold"), ", ", as_i("Italic")))
  }

  # -----------------------------------------------------------------------------
  # 6. Compose multi-line paragraphs in the 'para_test' column.
  #    Here we mix plain text with formatted chunks.
  ft <- compose(ft, j = "para_test", i = 1,
                value = as_paragraph(
                  "Line1", "\n",
                  as_b("Line2 Bold"), "\n",
                  as_i("Line3 Italic")
                ))

  # -----------------------------------------------------------------------------
  # 7. Apply different alignments.
  ft <- align(ft, j = "chunk_test", align = "left", part = "all")
  ft <- align(ft, j = "para_test", align = "center", part = "all")
  ft <- align(ft, j = "h1", align = "right", part = "all")

  # -----------------------------------------------------------------------------
  # 8. Prepend and append content in the 'append_test' column.
  #    Prepend a label and then append a suffix.
  ft <- compose(ft, j = "append_test",
                value = as_paragraph("Prepended: ", as_chunk(append_test)))
  ft <- append_chunks(ft, j = "append_test",
                      value = as_chunk(" :Appended"))

  # -----------------------------------------------------------------------------
  # 9. Set inner and outer borders with different colors and sizes.
  outer_border <- officer::fp_border(color = "red", width = 2)
  inner_border <- officer::fp_border(color = "#3333BB", width = 1)
  ft <- border_outer(ft, border = outer_border, part = "all")
  ft <- border_inner(ft, border = inner_border, part = "all")

  # -----------------------------------------------------------------------------
  # 10. Set padding and line spacing.
  ft <- padding(ft, padding = 5, part = "all")
  ft <- line_spacing(ft, space = 1.5, part = "all")

  # -----------------------------------------------------------------------------
  # 11. Merge cells horizontally and vertically.
  #     a) Horizontal merge in the body:
  #        In row 1, columns 'h1' and 'h2' share identical content ("Merge") so merge them.
  ft <- merge_h(ft, i = 1, part = "body")

  #     b) Vertical merge in column 'v1' for rows 2 and 3 (they are identical).
  ft <- merge_v(ft, j = "v1", part = "body")

  #     c) Simultaneous horizontal and vertical merging:
  #        For demonstration, force rows 2 and 3 in columns 'chunk_test' and 'para_test'
  #        to have identical content, then merge horizontally (across these two columns)
  #        and vertically (across rows 2 and 3).
  ft <- compose(ft, i = 2:3, j = c("chunk_test", "para_test"),
                value = as_paragraph("Combined"))

  # Define new border styles using fp_border:
  dashed_border <- officer::fp_border(color = "darkgreen", width = 1, style = "dashed")
  dotted_border <- officer::fp_border(color = "orange", width = 1.5, style = "dotted")
  double_border <- officer::fp_border(color = "purple", width = 3, style = "double")

  # Apply a double border to the bottom edge of the header row:
  ft <- border(ft, i = 1, border.bottom = double_border, part = "header")

  # Apply a dashed border on the left side of the "id" column in the body:
  ft <- border(ft, j = "id", border.left = dashed_border, part = "body")

  # Apply a dotted border on the right side of the "append_test" column in the body:
  ft <- border(ft, j = "append_test", border.right = dotted_border, part = "body")

  # Apply a dotted border on the bottom side of the "h1" column in the body:
  ft <- border(ft, j = "h1", border.bottom = dotted_border, part = "body")

  # Apply a double border on all sides of the "chunk_test" column in the body:
  ft <- border(ft, j = "chunk_test", border = double_border, part = "body")

  # For merged cells in the "chunk_test" and "para_test" columns (rows 2 and 3),
  # apply a combination: dashed border on the top and dotted border on the bottom.
  ft <- border(ft, i = 2:3, j = c("chunk_test", "para_test"),
               border.top = dashed_border, border.bottom = dotted_border,
               part = "body")

  # Apply a light cyan background to the 'id' column
  ft <- bg(ft, j = "id", bg = "#AAFAFA", part = "all")

  # Apply a light blueish background to the 'chunk_test' column
  ft <- bg(ft, j = "chunk_test", bg = "#ABBBFA", part = "all")

  # Apply an orange background to the 'para_test' column
  ft <- bg(ft, j = "para_test", bg = "orange", part = "all")

  # Merge combined
  ft <- merge_v(ft, j = 2:3)
  ft <- merge_h(ft, i = 2:3)

  ## Add colinfo
  ft <- add_header_row(ft, values = LETTERS[2:8])

  expect_no_warning(wb <- openxlsx2::wb_workbook() |>
    openxlsx2::wb_add_worksheet() |>
    flexlsx::wb_add_flextable(
      ft = ft,
      dims = "B2"
    ))


  test_wb_ft(wb,ft, "complex ft")


})


test_that("MeganMcAuliffe test", {
  skip_if_not_installed("flextable")
  skip_if(Sys.getenv("flexlsxtestdir")=="")


  ft <- readRDS(paste0(Sys.getenv("flexlsxtestdir"),
                       "ft_list_element.rds")) |>
    flextable::autofit()

  expect_no_warning(wb <- openxlsx2::wb_workbook() |>
                      openxlsx2::wb_add_worksheet() |>
                      flexlsx::wb_add_flextable(
                        ft = ft,
                        dims = "B2"
                      ))

  test_wb_ft(wb,ft, "MeganMcAuliffe ft")


})
