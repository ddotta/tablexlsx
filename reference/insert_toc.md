# Insert Table of Contents to an existing workbook

This function takes an existing workbook and a list or vector of sheet
names, and creates a Table of Contents sheet with hyperlinks.

## Usage

``` r
insert_toc(wb, sheets)
```

## Arguments

- wb:

  Workbook object to modify

- sheets:

  A named list in the format list("dataframe" = "sheet_name")

## Value

Nothing, this function is called for its side effects

## See also

\[toxlsx()\]
