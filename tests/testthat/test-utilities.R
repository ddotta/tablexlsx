test_that("calcstartrow returns correct result", {
  expect_equal(calcstartrow(1), 1)
  expect_equal(calcstartrow(2), 11)
  expect_equal(calcstartrow(3), 21)
})

test_that("calcskippedrow returns correct result", {
  mydf <- list(iris,cars,mtcars)
  expect_equal(calcskippedrow(mylist=mydf,x=1), 0)
  expect_equal(calcskippedrow(mylist=mydf,x=2), 150)
  expect_equal(calcskippedrow(mylist=mydf,x=3), 200)
  expect_equal(calcskippedrow(mylist=mydf,x=4), 232)
})

test_that("get_indices_of_identical_elements returns correct result", {
  myvector <- c("momo","momo","momo","mumu","mumu")
  expect_equal(length(get_indices_of_identical_elements(myvector)), 2)
  expect_equal(get_indices_of_identical_elements(myvector)[1], "1:3")
  expect_equal(get_indices_of_identical_elements(myvector)[2], "4:5")
})

test_that("get_indices_from_vector returns correct result", {
  myvector <- c("momo","momo","momo","mumu","mumu")
  expect_equal(length(get_indices_from_vector(myvector)), 1)
  expect_equal(get_indices_from_vector(myvector), "1:5")
})

test_that("convert_range_string returns correct result", {
  expect_equal(length(convert_range_string("1:3")), 3)
  expect_equal(convert_range_string("1:3"), c(1,2,3))
})
