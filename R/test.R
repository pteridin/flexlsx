
test <- function() {
  library(flextable)
  library(flexlsx)
  library(openxlsx2)

  theme_design <- function(x) {
    x <- border_remove(x)
    std_border <- fp_border_default(width = 4, color = "white")
    x <- fontsize(x, size = 10, part = "all")
    x <- font(x, fontname = "Courier", part = "all")
    x <- align(x, align = "center", part = "all")
    x <- bold(x, bold = TRUE, part = "all")
    x <- bg(x, bg = "#475f77", part = "body")
    x <- bg(x, bg = "#eb5555", part = "header")
    x <- bg(x, bg = "#1bbbda", part = "footer")
    x <- color(x, color = "white", part = "all")
    x <- padding(x, padding = 6, part = "all")
    x <- border_outer(x, part="all", border = std_border )
    x <- border_inner_h(x, border = std_border, part="all")
    x <- border_inner_v(x, border = std_border, part="all")
    x <- set_table_properties(x, layout = "fixed")
    x
  }

  ft <- flextable(head(airquality)) %>%
    add_footer_lines(
      c("Daily air quality measurements in New York, May to September 1973.",
        "Hummm, non non rien.")) %>%
    autofit() %>%
    add_header_lines("New York Air Quality Measurements") %>%
    theme_design()
  ft


  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars", grid_lines = FALSE)

  # add the flextable ft to the workbook, sheet "mtcars"
  # offset the table to cell 'C2'
  wb <- wb_add_flextable(wb, "mtcars", ft, dims = "C2")

  wb$open()

}

test2 <- function() {
  library(flextable)
  library(flexlsx)
  library(openxlsx2)

  library(officer)
  library(dplyr)
  library(safetyData)

  use_df_printer()

  set_flextable_defaults(
    border.color = "#AAAAAA", font.family = "Arial",
    font.size = 10, padding = 2, line_spacing = 1.5
  )
  adsl <- select(adam_adsl, AGE, SEX, BMIBLGR1, DURDIS, ARM)
  # adsl


  dat <- summarizor(adsl, by = "ARM")
  # dat

  ft <- as_flextable(dat, spread_first_col = TRUE, separate_with = "variable") %>%
    bold(i = ~ !is.na(variable), j = 1, bold = TRUE) %>%
    set_caption(
      autonum = officer::run_autonum(seq_id = "tab", bkm = "demo_tab", bkm_all = FALSE),
      fp_p = officer::fp_par(text.align = "left", padding = 5),
      align_with_table = FALSE,
      caption = as_paragraph(
        "Demographic Characteristics",
        "\nx.x: Study Subject Data"
      )
    ) %>%
    add_footer_lines("Source: ADaM adsl data frame from r package 'safetyData'") %>%
    fix_border_issues() %>%
    autofit()

  ft

  wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars", grid_lines = FALSE)

  # add the flextable ft to the workbook, sheet "mtcars"
  # offset the table to cell 'C2'
  wb <- wb_add_flextable(wb, "mtcars", ft, dims = "C2")

  wb$open()
}

minimal_color_fuckup <- function() {
  wb <- openxlsx2::wb_workbook()$add_worksheet("color_fu")

  black_style <- create_dxfs_style(font_color = wb_color("black"))

  wb$add_data(sheet = "color_fu",
              fmt_txt("This one goes bad",
                      color = "#000000") |>
                paste0(),
              dims = "A1",
              col_names = F)
  wb$save(tempfile(fileext = ".xlsx"))$open()

}


