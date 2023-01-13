library(testthat)

test_that("add_table function throws error when passing invalid data type for Table parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- "invalid_data_type"
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle),
    error = TRUE
  )
})

test_that("add_table function throws error when passing invalid data type for WbTitle parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame()
  WbTitle <- data.frame()
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for SheetTitle parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame()
  WbTitle <- wb
  SheetTitle <- 123
  TableTitle <- "Test Table"

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for TableTitle parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame()
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- 123

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for StartRow parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame()
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- "invalid_data_type"

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle, StartRow),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for StartCol parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame()
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartCol <- "invalid_data_type"

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle, StartCol = StartCol),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for StartRow parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:3, y = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- "invalid_data_type"
  StartCol <- 1
  FormatList <- list()

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle, StartRow, StartCol, FormatList),
    error = TRUE)
})

test_that("add_table function throws error when passing invalid data type for StartCol parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:3, y = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- 1
  StartCol <- "invalid_data_type"
  FormatList <- list()

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle, StartRow, StartCol, FormatList),
    error = TRUE)
})

test_that("add_table function throws error when passing invalid data type for TableFootnote1 parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:3, y = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- 1
  StartCol <- 1
  FormatList <- list()
  TableFootnote1 <- 123

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle,
              StartRow, StartCol, FormatList, TableFootnote1 = TableFootnote1),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for TableFootnote2 parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:3, y = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- 1
  StartCol <- 1
  FormatList <- list()
  TableFootnote2 <- 123

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle,
              StartRow, StartCol, FormatList, TableFootnote2 = TableFootnote2),
    error = TRUE
    )
})

test_that("add_table function throws error when passing invalid data type for TableFootnote3 parameter", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:3, y = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- 1
  StartCol <- 1
  FormatList <- list()
  TableFootnote3 <- 123

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle, StartRow, StartCol,
              FormatList, TableFootnote3 = TableFootnote3),
    error = TRUE
    )
})

test_that("add_table function creates a new sheet in the workbook if the sheet name does not exist", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:10, y = 11:20)
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  FormatList <- list()

  result <- add_table(Table, WbTitle, SheetTitle, TableTitle, FormatList)

  expect_true(SheetTitle %in% openxlsx::getSheets(result))
})

test_that("add_table function adds data frame to the specified sheet in the workbook", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(A=c(1,2,3), B=c("a","b","c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"

  add_table(Table, WbTitle, SheetTitle, TableTitle)

  read_table <- openxlsx::read.xlsx(WbTitle, sheet = SheetTitle)
  expect_equal(Table, read_table)
})

test_that("add_table function sets correct column widths and row heights", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(col1 = c("a", "b", "c"), col2 = c(1, 2, 3))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  StartRow <- 1
  StartCol <- 1
  FormatList <- list("col1" = "string_format", "col2" = "numeric_format")
  HeightTableTitle <- 2
  TableFootnote1 <- "Table footnote 1"
  TableFootnote2 <- "Table footnote 2"
  TableFootnote3 <- "Table footnote 3"

  add_table(Table, WbTitle, SheetTitle, TableTitle, StartRow, StartCol, FormatList, HeightTableTitle, TableFootnote1, TableFootnote2, TableFootnote3)

  expect_equal(openxlsx::getColWidths(WbTitle, sheet = SheetTitle)[[1]], c(45, 30, 20, 20, 20, ...))
  expect_equal(openxlsx::getRowHeights(WbTitle, sheet = SheetTitle)[[1]], c(20, 20 * HeightTableTitle, 20, ...))
})

test_that("add_table function writes table title and footnotes in correct location", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(A = 1:5, B = 6:10)
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  TableFootnote1 <- "Footnote 1"
  TableFootnote2 <- "Footnote 2"
  TableFootnote3 <- "Footnote 3"

  add_table(Table, WbTitle, SheetTitle, TableTitle, TableFootnote1 = TableFootnote1, TableFootnote2 = TableFootnote2, TableFootnote3 = TableFootnote3)

  expect_equal(openxlsx::readCell(WbTitle, SheetTitle, row = 1, col = 1), TableTitle)
  expect_equal(openxlsx::readCell(WbTitle, SheetTitle, row = nrow(Table) + 3, col = 1), TableFootnote1)
  expect_equal(openxlsx::readCell(WbTitle, SheetTitle, row = nrow(Table) + 4, col = 1), TableFootnote2)
  expect_equal(openxlsx::readCell(WbTitle, SheetTitle, row = nrow(Table) + 5, col = 1), TableFootnote3)
})

test_that("add_table function applies correct formatting to columns and rows of table", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = c(1, 2, 3), y = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  FormatList <- list(openxlsx::createStyle(numFmt = "0.0"), openxlsx::createStyle(fontColor = "red"))

  add_table(Table, WbTitle, SheetTitle, TableTitle, FormatList = FormatList)
  ws <- openxlsx::getWorksheet(WbTitle, SheetTitle)

  expect_equal(openxlsx::getCellStyle(ws, row = 3, col = 1)$numFmt, "0.0")
  expect_equal(openxlsx::getCellStyle(ws, row = 3, col = 2)$fontColor, "red")
  expect_equal(openxlsx::getCellStyle(ws, row = 4, col = 1)$numFmt, "0.0")
  expect_equal(openxlsx::getCellStyle(ws, row = 4, col = 2)$fontColor, "red")
  expect_equal(openxlsx::getCellStyle(ws, row = 5, col = 1)$numFmt, "0.0")
  expect_equal(openxlsx::getCellStyle(ws, row = 5, col = 2)$fontColor, "red")
})

test_that("add_table function returns the excel workbook object", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(A = c(1,2,3), B = c(4,5,6))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"

  result <- add_table(Table, WbTitle, SheetTitle, TableTitle)

  expect_class(result, "Workbook")
})
