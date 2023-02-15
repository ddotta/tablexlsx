<!-- badges: start -->
[![version](http://www.r-pkg.org/badges/version/tablexlsx)](https://CRAN.R-project.org/package=tablexlsx)
[![cranlogs](http://cranlogs.r-pkg.org/badges/tablexlsx)](https://CRAN.R-project.org/package=tablexlsx)
[![R check
status](https://github.com/ddotta/tablexlsx/workflows/R-CMD-check/badge.svg)](https://github.com/ddotta/tablexlsx/actions/workflows/check-release.yaml)
[![codecov](https://codecov.io/gh/ddotta/tablexlsx/branch/main/graph/badge.svg?token=UWLXVupq1C)](https://app.codecov.io/gh/ddotta/tablexlsx)
[![CodeFactor](https://www.codefactor.io/repository/github/ddotta/tablexlsx/badge)](https://www.codefactor.io/repository/github/ddotta/tablexlsx)
<!-- badges: end -->

:package: Package `tablexlsx` <img src="man/figures/hex_tablexlsx.png" width=110 align="right"/>
======================================

R package that allows to export data frames from R to xslx workbooks.

## Installation

To install `tablexlsx` from CRAN :  

``` r
install.packages("tablexlsx")
```

Or alternatively to install the development version from GitHub :  

``` r
remotes::install_github("ddotta/tablexlsx")
```

Then to load it :  

``` r
library(tablexlsx)
```

## Why this package ?

This package is a  wrapper of some functions from the great [openxlsx](https://github.com/ycphs/openxlsx) package.  

The purpose of this package is to meet the needs of R users who want to export data frames in `xlsx` files to share their data and results with other users not necessarily R users.  

The goal is to get an API that is as simple as possible to use. **It's intended to meet the needs of users who want to create clean, well-documented `xlsx` files (with a title and footnotes)**.  

The package will be improved over time and will include new features based on the needs expressed by users. To do so, feel free to open an issue [here](https://github.com/ddotta/tablexlsx/issues). 

## Advantages

Some advantages of using this package :  

- A simpler syntax for common export operations of excel files;  
- It can write several data frames in the same sheet one below the other;  
- It can merge modalities for one or several columns;  
- It can automatically open files as soon as they are created so you can inspect your workbook. 
- Data frames to be exported in `xlsx` files can be passed as an argument in functions as simple tables or list of tables.  
- The main function of the package [toxlsx()](https://ddotta.github.io/tablexlsx/reference/toxlsx.html) works interchangeably with `%>%` and `|>` operators and these 3 syntaxes below are equivalent :  

  
```
toxlsx(object = iris, path = mypath)
iris |> toxlsx(path = mypath)
iris %>% toxlsx(path = mypath)
```
  
:point_right: For more, [see examples gallery in vignette](https://ddotta.github.io/tablexlsx/articles/aa-examples.html) :mag_right:
