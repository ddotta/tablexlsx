test_that("insert_toc adds a sheet called Table of Contents", {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet1")
  openxlsx::addWorksheet(wb, "Sheet2")
  insert_toc(wb, sheets = list("iris" = "Sheet1", "mtcars" = "Sheet2"))
  expect_true("Table of Contents" %in% names(wb))
})

test_that("insert_toc returns invisibly NULL", {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "TestSheet")
  expect_invisible(
    insert_toc(wb, sheets = list("iris" = "Sheet1", "mtcars" = "Sheet2"))
  )
  expect_true("Table of Contents" %in% names(wb))
})

test_that("insert_toc does nothing when sheets is empty", {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "OnlySheet")
  expect_invisible(insert_toc(wb, sheets = list()))
  expect_false("Table of Contents" %in% names(wb))
})

test_that("insert_toc errors for invalid sheets", {
  wb <- openxlsx::createWorkbook()
  expect_error(
    insert_toc(wb, sheets = 123),
    "sheets must be a named list of worksheet names"
  )
})
