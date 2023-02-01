#' @name add_table
#'
#' @title Function that adds a data frame to an (existing) .xlsx workbook sheet
#'
#' @param Table : data frame to be exported to the workbook sheet
#' @param WbTitle : workbook
#' @param SheetTitle : string used for the sheet's name
#' @param TableTitle : string used for the data frame's title
#' @param StartRow : export start line number in the sheet (by default 1)
#' @param StartCol : export start column number in the sheet (by default 1)
#' @param FormatList : list that indicates the format of each column of the data frame
#' @param HeightTableTitle : multiplier (if needed) for the height of the title line (by default 2)
#' @param TableFootnote1 : string for TableFootnote1
#' @param TableFootnote2 : string for TableFootnote2
#' @param TableFootnote3 : string for TableFootnote3
#'
#' @return excel wb object
#'
#' @importFrom openxlsx addWorksheet setColWidths setRowHeights writeData addStyle writeDataTable
#' @export


add_table <- function(
    Table,
    WbTitle,
    SheetTitle,
    TableTitle,
    StartRow = 1,
    StartCol = 1,
    FormatList = list(),
    MergedCol = NULL,
    HeightTableTitle = 2,
    TableFootnote1 = list(),
    TableFootnote2 = list(),
    TableFootnote3 = list()) {

  # Assert parameters
  assert_class(Table, "data.frame")
  assert_class(WbTitle, "Workbook")
  assert_character1(SheetTitle)
  assert_character1(TableTitle)
  assert_numeric1(StartRow)
  assert_numeric1(StartCol)
  assert_class(FormatList, "list")
  assert_numeric1(HeightTableTitle)
  assert_character1(TableFootnote1)
  assert_character1(TableFootnote2)
  assert_character1(TableFootnote3)

  # If the sheet does not exist in the Excel file, we create it; otherwise, we invoke it
  if (!(SheetTitle %in% names(WbTitle))) {
    mysheet <- openxlsx::addWorksheet(WbTitle, SheetTitle)
  } else {
    mysheet <- SheetTitle
  }


  # Adjusting the size of columns and rows
  openxlsx::setColWidths(WbTitle, sheet = mysheet, cols = StartCol + 1, widths = 45)
  openxlsx::setColWidths(WbTitle, sheet = mysheet, cols = StartCol + 2, widths = 30)
  openxlsx::setColWidths(WbTitle, sheet = mysheet, cols = c(StartCol + 3:38), widths = 20)

  # Size of column headers
  openxlsx::setRowHeights(WbTitle, sheet = mysheet, rows = StartRow + 2, heights = 20 * HeightTableTitle)

  # Add title
  openxlsx::setRowHeights(WbTitle, sheet = mysheet, rows = StartRow, heights = 20)

  # Definition of column formats
  openxlsx::writeData(
    WbTitle,
    sheet = mysheet,
    x = TableTitle,
    startCol = StartCol,
    startRow = StartRow
  )

  openxlsx::addStyle(
    WbTitle,
    sheet = mysheet,
    cols = StartCol,
    rows = StartRow,
    style = style$title
  )


  # Add a table
  openxlsx::writeDataTable(
    wb = WbTitle,
    sheet = mysheet,
    x = Table,
    startRow = StartRow + 2,
    startCol = StartCol + 1,
    rowNames = FALSE,
    headerStyle = style$col_header
  )

  # Format of the table's columns
  sapply(seq_len(length(FormatList)), function(i) {
    openxlsx::addStyle(
      WbTitle,
      sheet = mysheet,
      cols = i + StartCol,
      rows = ((StartRow + 3):(StartRow + 2 + nrow(Table))),
      style = FormatList[[i]]
    )
  })

  # Add footnotes
  openxlsx::writeData(
    WbTitle,
    sheet = mysheet, x = TableFootnote1,
    startCol = StartCol, startRow = StartRow + nrow(Table) + 4
  )
  openxlsx::addStyle(
    WbTitle,
    sheet = mysheet,
    cols = StartCol, rows = StartRow + nrow(Table) + 4,
    style = style$footnote1
  )

  openxlsx::writeData(
    WbTitle,
    sheet = mysheet, x = TableFootnote2,
    startCol = StartCol, startRow = StartRow + nrow(Table) + 5
  )
  openxlsx::addStyle(
    WbTitle,
    sheet = mysheet,
    cols = StartCol, rows = StartRow + nrow(Table) + 5,
    style = style$footnote2
  )

  openxlsx::writeData(
    WbTitle,
    sheet = mysheet, x = TableFootnote3,
    startCol = StartCol, startRow = StartRow + nrow(Table) + 6
  )
  openxlsx::addStyle(
    WbTitle,
    sheet = mysheet,
    cols = StartCol, rows = StartRow + nrow(Table) + 6,
    style = style$footnote3
  )
}
