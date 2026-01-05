# Function that adds a data frame to an (existing) .xlsx workbook sheet

Function that adds a data frame to an (existing) .xlsx workbook sheet

## Usage

``` r
add_table(
  Table,
  WbTitle,
  SheetTitle,
  TableTitle,
  StartRow = 1,
  StartCol = 1,
  FormatList = setNames(rep(list(Theme[["character"]]), length(colnames(Table))),
    colnames(Table)),
  Theme = xls_theme_default(),
  HeightTableTitle = 2,
  TableFootnote1 = "",
  TableFootnote2 = "",
  TableFootnote3 = "",
  MergeCol = character(0),
  ByGroup = character(0),
  GroupName = FALSE,
  asTable = FALSE
)
```

## Arguments

- Table:

  : data frame to be exported to the workbook sheet

- WbTitle:

  : workbook

- SheetTitle:

  : string used for the sheet's name

- TableTitle:

  : string used for the data frame's title

- StartRow:

  : export start line number in the sheet (by default 1)

- StartCol:

  : export start column number in the sheet (by default 1)

- FormatList:

  : list that indicates the format of each column of the data frame

- Theme:

  : styling theme, a named list of \`openxlsx\` Styles

- HeightTableTitle:

  : multiplier (if needed) for the height of the title line (by default
  2)

- TableFootnote1:

  : string for TableFootnote1

- TableFootnote2:

  : string for TableFootnote2

- TableFootnote3:

  : string for TableFootnote3

- MergeCol:

  : character vector that indicates the columns for which to merge the
  modalities

- ByGroup:

  character vector indicating the name of the columns by which to group

- GroupName:

  boolean indicating whether the name of the grouping variable should be
  written

- asTable:

  logical indicating if data should be written as an Excel Table (FALSE
  by default)

## Value

excel wb object
