test_that("flextable with caption works", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(mtcars)
  wb <- openxlsx2::wb_workbook()$add_worksheet("wo caption")

  ## Without caption
  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = "wo caption",
                         dims = "B4")

  expect_equal(names(wb$to_df("wo caption")),
               c(NA,
                 names(mtcars)))
  ## With caption
  wb$add_worksheet("with caption")
  ft <- flextable::set_caption(ft, "This is a caption")
  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = "with caption",
                         dims = "B4")

  test_wb_ft(wb,ft, "caption")

  df <- wb$to_df("with caption")
  expect_equal(names(df)[2],
                 "This is a caption")

  expect_equal(unlist(df[1,]) |>
                 as.vector(),
               c(NA,
                 names(mtcars)))
})

test_that("flextable with complex caption works", {
  skip_if_not_installed("flextable")
  data("mtcars")

  ft <- flextable::flextable(mtcars)
  wb <- openxlsx2::wb_workbook()$add_worksheet("with caption")

  ## With caption
  caption <- flextable::as_paragraph(
    flextable::as_chunk("This is a complex caption",
             props = flextable::fp_text_default(font.family = "Cambria",
                                                 font.size = 14,
                                                 bold = TRUE,
                                                 italic = TRUE,
                                                 underlined = TRUE)))

  ft <- flextable::set_caption(ft, caption)

  wb <- wb_add_flextable(wb = wb,
                         ft = ft,
                         sheet = "with caption",
                         dims = "B4")
  test_wb_ft(wb,ft, "complex caption")

  df <- wb$to_df("with caption")
  expect_equal(names(df)[2],
               "This is a complex caption")

  expect_equal(unlist(df[1,]) |>
                 as.vector(),
               c(NA,
                 names(mtcars)))
})

