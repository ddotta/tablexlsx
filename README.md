<!-- badges: start -->
![GitHub top
language](https://img.shields.io/github/languages/top/ddotta/tablexlsx)
[![R check
status](https://github.com/ddotta/tablexlsx/workflows/R-CMD-check/badge.svg)](https://github.com/ddotta/tablexlsx/actions/workflows/check-release.yaml)
[![codecov](https://codecov.io/gh/ddotta/tablexlsx/branch/main/graph/badge.svg?token=UWLXVupq1C)](https://codecov.io/gh/ddotta/tablexlsx)
[![CodeFactor](https://www.codefactor.io/repository/github/ddotta/tablexlsx/badge)](https://www.codefactor.io/repository/github/ddotta/tablexlsx)
<!-- badges: end -->

:package: Package `tablexlsx` <img src="man/figures/hex_tablexlsx.png" width=110 align="right"/>
======================================

R package that allows to export data frames from R to xslx workbooks

## Installation

``` r
remotes::install_github("ddotta/tablexlsx")
```

``` r
library(tablexlsx)
```

## Why this package ?

This package is a  wrapper of some functions from the great [openxlsx](https://github.com/ycphs/openxlsx) package.  

The purpose of this package is to meet the needs of R users who want to export data frames in `xlsx` files to share their data and results with other users not necessarily R users.  

The goal is to get an API that is as simple as possible to use. The package will be improved over time and will include features such as the possibility to highlight in excel files some remarkable rows/columns (for example totals or coefficients...).

Some advantages of using this package :  

- A simpler syntax for common export operations of excel files;  
- The possibility to write several data frames in the same sheet one below the other;  
- The possibility to merge modalities for one or several columns

:point_right: [See examples gallery in vignette](https://ddotta.github.io/tablexlsx/articles/aa-examples.html)
