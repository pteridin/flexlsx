
<!-- README.md is generated from README.Rmd. Please edit that file -->

# flexlsx <img src="man/figures/logo.png" align="right" height="126"/>

<!-- badges: start -->

[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
[![Codecov test
coverage](https://codecov.io/gh/pteridin/flexlsx/branch/main/graph/badge.svg)](https://app.codecov.io/gh/pteridin/flexlsx?branch=main)
[![R-CMD-check](https://github.com/pteridin/flexlsx/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/pteridin/flexlsx/actions/workflows/R-CMD-check.yaml)
<!-- badges: end -->

The primary objective of `flexlsx` is to offer an effortless interface
for exporting `flextable` objects directly to Microsoft Excel. Building
upon the robust foundation provided by `openxlsx2` and `flextable`,
`flexlsx` ensures compatibility, precision, and efficiency when working
with both trivial and complex tables.

## Installation

You can install the development version of `flexlsx` like so:

``` r
# install.packages("remotes")
remotes::install_github("pteridin/flexlsx")
```

Or install the CRAN release like so:

``` r
install.packages("flexlsx")
```

## Example

This is a basic example which shows you how to solve a common problem:

``` r
library(flexlsx)

# Create a flextable and an openxlsx2 workbook
ft <- flextable::as_flextable(table(mtcars[,1:2]))
wb <- openxlsx2::wb_workbook()$add_worksheet("mtcars")

# add the flextable ft to the workbook, sheet "mtcars"
# offset the table to cell 'C2'
wb <- wb_add_flextable(wb, "mtcars", ft, dims = "C2")

# save the workbook to a temporary xlsx file
tmpfile <- tempfile(fileext = ".xlsx")
wb$save(tmpfile)
```
