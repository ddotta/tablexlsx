test_that("toxlsx works correctly with a simple data frame", {
  iris %>% toxlsx(path = tempdir())
  new_table <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "Sheet 1",
                                   rows = 3:153, cols = 2:6)
  expect_equal(nrow(iris), nrow(new_table))
  expect_equal(ncol(iris), ncol(new_table))
  expect_true("Sheet 1" %in% openxlsx::getSheetNames(file.path(tempdir(),"Export.xlsx")))
})

test_that("toxlsx works correctly with a list", {
  list(iris,cars) |> toxlsx(path = tempdir())
  new_table_iris <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                        sheet = "Sheet 1",
                                        rows = 3:153, cols = 2:6)
  expect_equal(nrow(iris), nrow(new_table_iris))
  expect_equal(ncol(iris), ncol(new_table_iris))

  new_table_cars <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                        sheet = "Sheet 2",
                                        rows = 3:53, cols = 2:12)
  expect_equal(cars, new_table_cars)
  expect_equal(nrow(cars), nrow(new_table_cars))
  expect_equal(ncol(cars), ncol(new_table_cars))
  expect_true("Sheet 1" %in% openxlsx::getSheetNames(file.path(tempdir(),"Export.xlsx")))
  expect_true("Sheet 2" %in% openxlsx::getSheetNames(file.path(tempdir(),"Export.xlsx")))
})


test_that("toxlsx works correctly with several data frames in a same sheet", {
  list(iris,cars) %>%
    toxlsx(tosheet = list("iris" = "mydata",
                          "cars" = "mydata"),
           footnote1 = list("iris" = "The data set contains 3 classes of 50 instances each, where each class refers to a type of iris plant.",
                            "cars" = "The data give the speed of cars and the distances taken to stop. Note that the data were recorded in the 1920s."),
           footnote2 = list("iris" = "Predicted attribute: class of iris plant.",
                            "cars" = "Data recorded in the 1920s"),
           footnote3 = list("iris" = "Source : R.A. Fisher",
                            "cars" = "Source : M. Ezekiel"),
           path = tempdir())

  new_table_iris <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "mydata",
                                   rows = 3:153, cols = 2:6)
  expect_equal(nrow(iris), nrow(new_table_iris))
  expect_equal(ncol(iris), ncol(new_table_iris))

  new_table_cars <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "mydata",
                                   rows = 163:213, cols = 2:3)
  expect_equal(cars, new_table_cars)
  expect_equal(nrow(cars), nrow(new_table_cars))
  expect_equal(ncol(cars), ncol(new_table_cars))
})

test_that("toxlsx works correctly with a lot of specifications", {

    list(iris,cars) %>%
    toxlsx(tosheet = list("iris" = "first",
                          "cars" = "second"),
           title = list("iris" = "Table iris",
                        "cars" = "Table cars"),
           # The `columnstyle` argument is optional in toxlsx().
           # It is used only if you want to specify the format of each column
           columnstyle = list("iris" = list("c1" = "decimal",
                                            "c2" = "decimal",
                                            "c3" = "number",
                                            "c4" = "number",
                                            "c5" = "character"),
                              "cars" =  list("c1" = "number",
                                             "c2" = "number")),
           footnote1 = list("iris" = "The data set contains 3 classes of 50 instances each, where each class refers to a type of iris plant.",
                            "cars" = "The data give the speed of cars and the distances taken to stop. Note that the data were recorded in the 1920s."),
           footnote2 = list("iris" = "Predicted attribute: class of iris plant.",
                            "cars" = "Data recorded in the 1920s"),
           footnote3 = list("iris" = "Source : R.A. Fisher",
                            "cars" = "Source : M. Ezekiel"),
           filename = "Results",
           path = tempdir())

  new_table_iris <- openxlsx::read.xlsx(file.path(tempdir(),"Results.xlsx"),
                                        sheet = "first",
                                        rows = 3:153, cols = 2:6)
  expect_equal(nrow(iris), nrow(new_table_iris))
  expect_equal(ncol(iris), ncol(new_table_iris))

  new_table_cars <- openxlsx::read.xlsx(file.path(tempdir(),"Results.xlsx"),
                                        sheet = "second",
                                        rows = 3:53, cols = 2:3)
  expect_equal(cars, new_table_cars)
  expect_equal(nrow(cars), nrow(new_table_cars))
  expect_equal(ncol(cars), ncol(new_table_cars))
})

