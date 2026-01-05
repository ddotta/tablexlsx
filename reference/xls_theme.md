# Constructor function for xls themes

This function creates an xls theme for styling exported tables. All its
arguments must be \`openxlsx\` Style objects.

## Usage

``` r
xls_theme(
  title,
  col_header,
  character,
  footnote1,
  footnote2,
  footnote3,
  mergedcell,
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

[`xls_theme_plain()`](xls_theme_plain.md),
[`xls_theme_default()`](xls_theme_default.md)

## Examples

``` r
my_theme <- xls_theme(
  title = openxlsx::createStyle(),
  col_header = openxlsx::createStyle(),
  character = openxlsx::createStyle(),
  footnote1 = openxlsx::createStyle(),
  footnote2 = openxlsx::createStyle(),
  footnote3 = openxlsx::createStyle(),
  mergedcell = openxlsx::createStyle()
)

if (FALSE) { # \dontrun{
toxlsx(object = iris, path = tempdir(), theme = my_theme)
} # }
```
