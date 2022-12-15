#' @title Function that adds a data frame or tibble to an (existing) .xlsx workbook sheet
#'
#' @param Table : data frame or tibble to be exported to the workbook sheet
#' @param WbTitle : workbook
#' @param SheetTitle : string used for the sheet's name
#' @param TableTitle : string used for the data frame/tibble's title
#' @param StartRow : export start line number in the sheet (by default 1)
#' @param StartCol : export start column number in the sheet (by default 1)
#' @param FormatList : list that indicates the format of each column of the data frame/tibble
#' @param HeightTableTitle : multiplier (if needed) for the height of the title line (by default 2)
#' @param Footnote1 : string for footnote1
#' @param Footnote2 : string for footnote2
#' @param Footnote3 : string for footnote1
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
    FormatList,
    HeightTableTitle = 2,
    Footnote1,
    Footnote2,
    Footnote3){


  # If the sheet does not exist in the Excel file, we create it; otherwise, we invoke it
  if (!(SheetTitle %in% names(WbTitle))) {
    mysheet <- openxlsx::addWorksheet(WbTitle, SheetTitle)
  } else {
    mysheet <- SheetTitle
  }


  # Adjusting the size of columns and rows
  openxlsx::setColWidths(WbTitle, sheet = mysheet, cols = StartCol + 1,widths = 45)
  openxlsx::setColWidths(WbTitle, sheet = mysheet, cols = StartCol + 2,widths = 30)
  openxlsx::setColWidths(WbTitle, sheet = mysheet, cols = c(StartCol + 3:38),widths = 20)

  # Size of column headers
  openxlsx::setRowHeights(WbTitle, sheet = mysheet, rows = StartRow + 2, heights = 20*HeightTableTitle)

  # Add title
  openxlsx::setRowHeights(WbTitle, sheet = mysheet, rows = StartRow, heights = 20)

  # Definition of column formats
  openxlsx::writeData(WbTitle,
                      sheet = mysheet,
                      x = TableTitle,
                      startCol = StartCol,
                      startRow = StartRow)

  openxlsx::addStyle(WbTitle,
                     sheet = mysheet,
                     cols = StartCol,
                     rows = StartRow,
                     style = style_title)


  ## Style's definitions
  # For title
  style_title <- openxlsx::createStyle(fontSize = 16, textDecoration = c("bold"))
  # For footnote1
  style_footnote1 <- openxlsx::createStyle(fontSize = 12)
  # For footnote2
  style_footnote2 <- openxlsx::createStyle(fontSize = 12, textDecoration = c("bold"))
  # For footnote3
  style_footnote3 <- openxlsx::createStyle(fontSize = 12, textDecoration = c("bold"))
  # For column headers
  style_col_header <- openxlsx::createStyle(fontSize = 12,
                                            textDecoration = c("bold"),
                                            border = c("top", "bottom", "left", "right"),
                                            borderStyle = "thin",
                                            wrapText = TRUE,
                                            halign = "center")



  # Add a table
  openxlsx::writeDataTable(wb = WbTitle,
                           sheet = mysheet,
                           x = Table,
                           startRow = StartRow + 2,
                           startCol = StartCol + 1,
                           rowNames = FALSE,
                           headerStyle = style_col_header)

  # Format of the table's columns
  assign_format <- function(i) {
    openxlsx::addStyle(WbTitle, sheet = mysheet,
                       cols = i + StartCol,
                       rows = ((StartRow + 3):(StartRow + 2 + nrow(Table))),
                       style = FormatList[[i]])
  }

  sapply(seq_len(length(FormatList)), assign_format)


  # Add footnotes
  openxlsx::writeData(WbTitle, sheet = mysheet, x = Footnote1,
                      startCol = StartCol, startRow = StartRow + nrow(Table) + 4)
  openxlsx::addStyle(WbTitle, sheet = mysheet,
                     cols = StartCol, rows = StartRow + nrow(Table) + 4,
                     style = style_footnote1)

  openxlsx::writeData(WbTitle, sheet = mysheet, x = Footnote2,
                      startCol = StartCol, startRow = StartRow + nrow(Table) + 5)
  openxlsx::addStyle(WbTitle, sheet = mysheet,
                     cols = StartCol, rows = StartRow + nrow(Table) + 5,
                     style = style_footnote2)

  openxlsx::writeData(WbTitle, sheet = mysheet, x = Footnote3,
                      startCol = StartCol, startRow = StartRow + nrow(Table) + 6)
  openxlsx::addStyle(WbTitle, sheet = mysheet,
                     cols = StartCol, rows = StartRow + nrow(Table) + 6,
                     style = style_footnote3)

}
