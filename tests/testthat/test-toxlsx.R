# Creation of data frames
x <- iris
x$Species <- as.character(x$Species)
y <- cars

df1 <- data.frame(
  group = c("dupont","dupont","arnold","arnold"),
  name = c("toto","tata","tutu","tete"),
  volume = c(10,8,12,15)
)
df2 <- data.frame(
  country = c(rep("france",4),rep("england",4)),
  group = c(rep("dupont",2),rep("martin",2),
            rep("arnold",2),rep("harry",2)),
  name = c("toto","tata","tutu","tete",
           "momo","mama","mumu","meme"),
  volume = c(10,8,12,15,
             14,10,5,12)
)

ext_iris <- iris %>% head()
ext_iris$Species <- as.character(ext_iris$Species)
ext_cars <- cars %>% head()

tb1 <- data.frame(tables = c(rep("iris",5),rep("cars",2)),
                  var = c(names(iris),names(cars)))
tb2 <- data.frame(tables = c("iris","cars","cars"),
                  rownumber = c(150,50,32))
# Tests
test_that("toxlsx works correctly with a simple data frame", {
  x %>% toxlsx(path = tempdir())
  new_table <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "Sheet 1",
                                   rows = 3:153, cols = 2:6)
  expect_equal(x, new_table)
  expect_equal(nrow(x), nrow(new_table))
  expect_equal(ncol(x), ncol(new_table))
  expect_true("Sheet 1" %in% openxlsx::getSheetNames(file.path(tempdir(),"Export.xlsx")))
})

test_that("toxlsx works correctly with a list", {
  list(x,y) |> toxlsx(path = tempdir())
  new_table_x <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "Sheet 1",
                                   rows = 3:153, cols = 2:6)
  expect_equal(x, new_table_x)
  expect_equal(nrow(x), nrow(new_table_x))
  expect_equal(ncol(x), ncol(new_table_x))
  new_table_y <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "Sheet 2",
                                   rows = 3:53, cols = 2:12)
  expect_equal(y, new_table_y)
  expect_equal(nrow(y), nrow(new_table_y))
  expect_equal(ncol(y), ncol(new_table_y))
  expect_true("Sheet 1" %in% openxlsx::getSheetNames(file.path(tempdir(),"Export.xlsx")))
  expect_true("Sheet 2" %in% openxlsx::getSheetNames(file.path(tempdir(),"Export.xlsx")))
})

test_that("toxlsx works correctly with mergecol argument", {
  df1 %>% toxlsx(path =  tempdir(), mergecol = "group")

  new_table <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "Sheet 1",
                                   rows = 3:7, cols = 2:4)
  expect_equal(df1, new_table)
  expect_equal(nrow(df1), nrow(new_table))
  expect_equal(ncol(df1), ncol(new_table))
})

test_that("toxlsx works correctly while merging several columns", {
  df2 %>% toxlsx(path = tempdir(), mergecol = c("country","group"))

  new_table <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "Sheet 1",
                                   rows = 3:11, cols = 2:5)
  expect_equal(df2, new_table)
  expect_equal(nrow(df2), nrow(new_table))
  expect_equal(ncol(df2), ncol(new_table))
})

test_that("toxlsx works correctly with several data frames in a same sheet", {
  list(tb1,tb2) %>%
    toxlsx(tosheet = list("tb1" = "mydata",
                          "tb2" = "mydata"),
           footnote1 = list("tb1" = "The data set contains 3 classes of 50 instances each, where each class refers to a type of iris plant.",
                            "tb2" = "The data give the speed of cars and the distances taken to stop. Note that the data were recorded in the 1920s."),
           footnote2 = list("tb1" = "Predicted attribute: class of iris plant.",
                            "tb2" = "Data recorded in the 1920s"),
           footnote3 = list("tb1" = "Source : R.A. Fisher",
                            "tb2" = "Source : M. Ezekiel"),
           path = tempdir())

  new_table_tb1 <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "mydata",
                                   rows = 3:10, cols = 2:3)
  expect_equal(tb1, new_table_tb1)
  expect_equal(nrow(tb1), nrow(new_table_tb1))
  expect_equal(ncol(tb1), ncol(new_table_tb1))

  new_table_tb2 <- openxlsx::read.xlsx(file.path(tempdir(),"Export.xlsx"),
                                   sheet = "mydata",
                                   rows = 20:23, cols = 2:3)
  expect_equal(tb2, new_table_tb2)
  expect_equal(nrow(tb2), nrow(new_table_tb2))
  expect_equal(ncol(tb2), ncol(new_table_tb2))
})

test_that("toxlsx works correctly with a lot of specifications", {
  list(iris,cars) %>%
    toxlsx(tosheet = list("ext_iris" = "first",
                          "ext_cars" = "second"),
           title = list("ext_iris" = "Head of iris",
                        "ext_cars" = "Head of cars"),
           # The `columnstyle` argument is optional in toxlsx().
           # It is used only if you want to specify the format of each column
           columnstyle = list("ext_iris" = list("c1" = "decimal",
                                            "c2" = "decimal",
                                            "c3" = "number",
                                            "c4" = "number",
                                            "c5" = "character"),
                              "ext_cars" =  list("c1" = "number",
                                             "c2" = "number")),
           footnote1 = list("ext_iris" = "The data set contains 3 classes of 50 instances each, where each class refers to a type of iris plant.",
                            "ext_cars" = "The data give the speed of cars and the distances taken to stop. Note that the data were recorded in the 1920s."),
           footnote2 = list("ext_iris" = "Predicted attribute: class of iris plant.",
                            "ext_cars" = "Data recorded in the 1920s"),
           footnote3 = list("ext_iris" = "Source : R.A. Fisher",
                            "ext_cars" = "Source : M. Ezekiel"),
           filename = "Results",
           path = tempdir())

  new_table_iris <- openxlsx::read.xlsx(file.path(tempdir(),"Results.xlsx"),
                                        sheet = "first",
                                        rows = 3:9, cols = 2:6)
  expect_equal(ext_iris, new_table_iris)
  expect_equal(nrow(ext_iris), nrow(new_table_iris))
  expect_equal(ncol(ext_iris), ncol(new_table_iris))

  new_table_cars <- openxlsx::read.xlsx(file.path(tempdir(),"Results.xlsx"),
                                        sheet = "second",
                                        rows = 3:9, cols = 2:3)
  expect_equal(ext_cars, new_table_cars)
  expect_equal(nrow(ext_cars), nrow(new_table_cars))
  expect_equal(ncol(ext_cars), ncol(new_table_cars))
})

