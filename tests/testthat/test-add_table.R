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
  StartRow <- "1"

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
  StartCol <- "1"

  expect_snapshot(
    add_table(Table, WbTitle, SheetTitle, TableTitle, StartCol = StartCol),
    error = TRUE
  )
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
      StartRow, StartCol, FormatList,
      TableFootnote1 = TableFootnote1
    ),
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
      StartRow, StartCol, FormatList,
      TableFootnote2 = TableFootnote2
    ),
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
      FormatList,
      TableFootnote3 = TableFootnote3
    ),
    error = TRUE
  )
})

test_that("add_table function creates a new sheet in the workbook if the sheet name does not exist", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(x = 1:10, y = 11:20)
  WbTitle <- wb
  SheetTitle <- "Sheet 1"
  TableTitle <- "Test Table"
  FormatList <- list()

  add_table(Table, WbTitle, SheetTitle, TableTitle, FormatList,
            StartRow=1, StartCol = 1,
            TableFootnote1 = "note1",TableFootnote2 = "note2",TableFootnote3 = "note3")

  openxlsx::saveWorkbook(
    WbTitle,
    file.path(
      tempdir(),
      "test.xlsx"
    ),
    overwrite = TRUE
  )

  expect_true(SheetTitle %in% openxlsx::getSheetNames(file.path(tempdir(),"test.xlsx")))
})

test_that("add_table function writes the same data frame in the workbook", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(A = c(1, 2, 3), B = c("a", "b", "c"))
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"
  FormatList <- list()

  add_table(Table, WbTitle, SheetTitle, TableTitle, FormatList,
            StartRow=1, StartCol = 1,
            TableFootnote1 = "note1",TableFootnote2 = "note2",TableFootnote3 = "note3")

  openxlsx::saveWorkbook(
    WbTitle,
    file.path(
      tempdir(),
      "test.xlsx"
    ),
    overwrite = TRUE
  )

  new_table <- openxlsx::read.xlsx(file.path(tempdir(),"test.xlsx"), sheet = SheetTitle, startRow = 2)[1:3,2:3]
  expect_equal(Table, new_table)
})

test_that("add_table function writes footnotes in correct locations", {
  wb <- openxlsx::createWorkbook()
  Table <- data.frame(A = 1:5, B = 6:10)
  WbTitle <- wb
  SheetTitle <- "Sheet1"
  TableTitle <- "Test Table"

  add_table(Table, WbTitle, SheetTitle, TableTitle, FormatList,
            StartRow=1, StartCol = 1,
            TableFootnote1 = "note1",TableFootnote2 = "note2",TableFootnote3 = "note3")

  openxlsx::saveWorkbook(
    WbTitle,
    file.path(
      tempdir(),
      "test.xlsx"
    ),
    overwrite = TRUE
  )

  expect_equal(openxlsx::read.xlsx(file.path(tempdir(),"test.xlsx"), SheetTitle, startRow = nrow(Table) + 3)[1,1], "note1")
  expect_equal(openxlsx::read.xlsx(file.path(tempdir(),"test.xlsx"), SheetTitle, startRow = nrow(Table) + 3)[2,1], "note2")
  expect_equal(openxlsx::read.xlsx(file.path(tempdir(),"test.xlsx"), SheetTitle, startRow = nrow(Table) + 3)[3,1], "note3")
})
