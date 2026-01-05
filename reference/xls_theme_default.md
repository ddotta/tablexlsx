# Constructor function for the default xls theme

This function is a wrapper around \[xls_theme()\] that creates an xls
theme for styling exported tables. It defines a theme whith sensible
default formatting values. It also defines custom styles for "number",
"decimal" and "percent column types. All its arguments must be
\`openxlsx\` Style objects.

## Usage

``` r
xls_theme_default(
  title = openxlsx::createStyle(fontSize = 16, textDecoration = "bold"),
  footnote1 = openxlsx::createStyle(fontSize = 12),
  footnote2 = openxlsx::createStyle(fontSize = 12),
  footnote3 = openxlsx::createStyle(fontSize = 12),
  col_header = openxlsx::createStyle(fontSize = 12, textDecoration = "bold", border =
    c("top", "bottom", "left", "right"), borderStyle = "thin", wrapText = TRUE, halign =
    "center"),
  character = openxlsx::createStyle(fontSize = 12, border = c("top", "bottom", "left",
    "right"), borderStyle = "thin"),
  number = openxlsx::createStyle(fontSize = 12, numFmt = "### ### ### ##0", border =
    c("top", "bottom", "left", "right"), borderStyle = "thin"),
  decimal = openxlsx::createStyle(fontSize = 12, numFmt = "### ### ### ##0.0", border =
    c("top", "bottom", "left", "right"), borderStyle = "thin"),
  percent = openxlsx::createStyle(fontSize = 12, numFmt = "#0.0", border = c("top",
    "bottom", "left", "right"), borderStyle = "thin", halign = "center"),
  mergedcell = openxlsx::createStyle(fontSize = 12, border = c("top", "bottom", "left",
    "right"), borderStyle = "thin", wrapText = TRUE, valign = "center", halign =
    "center"),
  ...
)
```

## Arguments

- title:

  Style for the title

- footnote1:

  Style for footnote1

- footnote2:

  Style for footnote2

- footnote3:

  Style for footnote3

- col_header:

  Style for the columns header

- character:

  Default style for data cells

- number:

  Style for columns in number format

- decimal:

  Style for columns in decimal format

- percent:

  Style for columns in percent format

- mergedcell:

  Style for merged cells

- ...:

  Other (named) custom styles

## Value

a named list of class xls_theme, whose elements are \`openxlsx\` Style
objects.

## See also

[`xls_theme()`](xls_theme.md), [`xls_theme_plain()`](xls_theme_plain.md)

## Examples

``` r
# default theme
xls_theme_default()
#> $title
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 16 
#>  Font decoration: BOLD 
#>  
#> 
#> 
#> $col_header
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 12 
#>  Font decoration: BOLD 
#>  Cell borders: Top: thin, Bottom: thin, Left: thin, Right: thin 
#>  Cell border colours: #000000, #000000, #000000, #000000 
#>  Cell horz. align: center 
#>  wraptext: TRUE 
#> 
#> 
#> $character
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 12 
#>  Cell borders: Top: thin, Bottom: thin, Left: thin, Right: thin 
#>  Cell border colours: #000000, #000000, #000000, #000000 
#>  
#> 
#> 
#> $footnote1
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 12 
#>  
#> 
#> 
#> $footnote2
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 12 
#>  
#> 
#> 
#> $footnote3
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 12 
#>  
#> 
#> 
#> $mergedcell
#> A custom cell style. 
#> 
#>  Cell formatting: GENERAL 
#>  Font size: 12 
#>  Cell borders: Top: thin, Bottom: thin, Left: thin, Right: thin 
#>  Cell border colours: #000000, #000000, #000000, #000000 
#>  Cell horz. align: center 
#>  Cell vert. align: center 
#>  wraptext: TRUE 
#> 
#> 
#> $number
#> A custom cell style. 
#> 
#>  Cell formatting: "### ### ### ##0" 
#>  Font size: 12 
#>  Cell borders: Top: thin, Bottom: thin, Left: thin, Right: thin 
#>  Cell border colours: #000000, #000000, #000000, #000000 
#>  
#> 
#> 
#> $decimal
#> A custom cell style. 
#> 
#>  Cell formatting: "### ### ### ##0.0" 
#>  Font size: 12 
#>  Cell borders: Top: thin, Bottom: thin, Left: thin, Right: thin 
#>  Cell border colours: #000000, #000000, #000000, #000000 
#>  
#> 
#> 
#> $percent
#> A custom cell style. 
#> 
#>  Cell formatting: "#0.0" 
#>  Font size: 12 
#>  Cell borders: Top: thin, Bottom: thin, Left: thin, Right: thin 
#>  Cell border colours: #000000, #000000, #000000, #000000 
#>  Cell horz. align: center 
#>  
#> 
#> 
#> attr(,"class")
#> [1] "xls_theme"

# default theme with title in italic
my_theme <- xls_theme_default(title = openxlsx::createStyle(textDecoration = "italic"))

if (FALSE) { # \dontrun{
toxlsx(object = iris, path = tempdir(), theme = my_theme)
} # }
```
