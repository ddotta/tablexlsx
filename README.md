:package: Package `tablexlsx` <img src="man/figures/hex_tablexlsx.png" width=110 align="right"/>
======================================

R package that allows to convert export data frames from R to xslx workbooks

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


## Examples

This package will allow you to make exports in xlsx format, whether they are simple or very customized :

- Simply export a data frame to an xlsx file 

``` r
iris %>% toxlsx()
```

- Export a list of data frames to an xlsx file by specifying which data frame goes in which sheet, styling each column, giving a title and footnotes...

``` r
list(iris,cars) %>%
  toxlsx(tosheet = list("iris" = "first",
                        "cars" = "second"),
         title = list("iris" = "Presentation of iris",
                      "cars" = "Data about cars"),
         columnstyle = list("iris" = list("c1" = "number",
                                          "c2" = "number",
                                          "c3" = "number",
                                          "c4" = "number",
                                          "c5" = "character"),
                            "cars" =  list("c1" = "number",
                                           "c2" = "number")),
         footnote1 = list("iris" = "The data set contains 3 classes of 50 instances each, where each class refers to a type of iris plant.",
                           "cars" = "The data give the speed of cars and the distances taken to stop. Note that the data were recorded in the 1920s."),
         footnote2 = list("iris" = "Predicted attribute: class of iris plant.",
                           "cars" = "Data recorded in the 1920s"),
         footnote3 = list("iris" = "Source : R.A. Fisher",
                           "cars" = "Source : M. Ezekiel"),
         path = "C:/Users/AQEW8W/Downloads")
```
