# Convert R data frames to excel files

This function allows you to write R data frames given in the \`object\`
argument to excel files located in the \`path\` directory. The function
takes several arguments but the only two required are \`object\` and
\`path\`.  
See examples gallery :
\<https://ddotta.github.io/tablexlsx/articles/aa-examples.html\>

## Usage

``` r
toxlsx(
  object,
  path,
  tosheet = list(),
  title = list(),
  table_of_contents = FALSE,
  columnstyle = list(default = NULL),
  theme = xls_theme_default(),
  footnote1 = list(),
  footnote2 = list(),
  footnote3 = list(),
  mergecol = NULL,
  bygroup = list(),
  groupname = FALSE,
  filename = "Export",
  asTable = FALSE,
  automaticopen = FALSE
)
```

## Arguments

- object:

  data.frame to be converted to excel

- path:

  path to save excel file (either a directory name or a file name with
  full path)

- tosheet:

  list of sheet names for each element of object. If omitted, sheets are
  named by default "Sheet 1", "Sheet 2"...

- title:

  list of title for each element of object If omitted, title takes the
  name of the dataframe in \`object\`

- table_of_contents:

  boolean indicating if a table of contents sheet should be included in
  the final workbook

- columnstyle:

  list of style for columns of each element of object Only useful if you
  want to customise the style of each column \`

- theme:

  styling theme, a named list of \`openxlsx\` Styles

- footnote1:

  list of footnote1 for each element of object If omitted, no footnote1

- footnote2:

  list of footnote2 for each element of object If omitted, no footnote2

- footnote3:

  list of footnote3 for each element of object If omitted, no footnote3

- mergecol:

  list of character vectors that indicate the columns for which we want
  to merge the modalities

- bygroup:

  list of character vectors indicating the names of the columns by which
  to group

- groupname:

  list of booleans indicating whether the names of the grouping
  variables should be written

- filename:

  name for the excel file ("Export" by default). Ignored if \`path\` is
  a file name.

- asTable:

  logical indicating if data should be written as an Excel Table (FALSE
  by default)

- automaticopen:

  logical indicating if excel file should open automatically (FALSE by
  default)

## Value

an excel file

## Examples

``` r
# Simply export a data frame to an xlsx file
# For more examples, see examples gallery :
# https://ddotta.github.io/tablexlsx/articles/aa-examples.html
if (FALSE) { # \dontrun{
toxlsx(object = iris, path = tempdir())
} # }
```
