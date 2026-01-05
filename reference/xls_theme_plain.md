# Constructor function for a plain xls theme

This function is a wrapper around \[xls_theme()\] that creates an xls
theme for styling exported tables. It defines a simple theme whith no
special formatting. All its arguments must be \`openxlsx\` Style
objects.

## Usage

``` r
xls_theme_plain(
  title = openxlsx::createStyle(),
  col_header = openxlsx::createStyle(),
  character = openxlsx::createStyle(),
  footnote1 = openxlsx::createStyle(),
  footnote2 = openxlsx::createStyle(),
  footnote3 = openxlsx::createStyle(),
  mergedcell = openxlsx::createStyle(),
  ...
)
```

## Arguments

- title:

  Style for the title

- col_header:

  Style for the columns header

- character:

  Default style for data cells

- footnote1:

  Style for footnote1

- footnote2:

  Style for footnote2

- footnote3:

  Style for footnote3

- mergedcell:

  Style for merged cells

- ...:

  Other (named) custom styles

## Value

a named list of class xls_theme, whose elements are \`openxlsx\` Style
objects.

## See also

[`xls_theme()`](xls_theme.md),
[`xls_theme_default()`](xls_theme_default.md)

## Examples

``` r
# plain theme
xls_theme_plain()
#> $title
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> $col_header
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> $character
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> $footnote1
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> $footnote2
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> $footnote3
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> $mergedcell
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  
#> 
#> 
#> attr(,"class")
#> [1] "xls_theme"

# plain theme with title in bold
my_theme <- xls_theme_plain(title = openxlsx::createStyle(textDecoration = "bold"))

if (FALSE) { # \dontrun{
toxlsx(object = iris, path = tempdir(), theme = my_theme)
} # }
```
