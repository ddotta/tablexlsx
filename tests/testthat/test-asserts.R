test_that("assert_class function works as expected", {
  # Test 1: Input is of correct class for numeric input
  x <- 1
  assert_class(x, "numeric")
  expect_true(TRUE)

  # Test 2: Input is of correct class for character input
  x <- "1"
  assert_class(x, "character", TRUE)
  expect_true(TRUE)

  # Test 3: Input is of correct class for list input
  x <- list()
  assert_class(x, "list", TRUE)
  expect_true(TRUE)

  # Test 4: Input is of correct class for data.frame input
  x <- data.frame()
  assert_class(x, "data.frame", TRUE)
  expect_true(TRUE)

  # Test 5: assert_class returns an error
  x <- list()
  expect_snapshot(
    assert_class(x, "matrix", TRUE),
    error = TRUE)
})

test_that("assert_character1 function works as expected", {
  # Test 1: Input is of correct class for character input
  assert_character1("1")
  expect_true(TRUE)

  # Test 2: assert_character1 returns an error with character vector and several elements
  expect_snapshot(
    assert_character1(c("1","2")),
    error = TRUE)

  # Test 3: assert_character1 returns an error with numeric input
  expect_snapshot(
    assert_character1(1),
    error = TRUE)
})

test_that("assert_numeric1 function works as expected", {
  # Test 1: Input is of correct class for numeric input
  assert_numeric1(1)
  expect_true(TRUE)

  # Test 2: Input is of correct class for numeric input and scalar as true
  assert_numeric1(1, scalar = TRUE)
  expect_true(TRUE)

  # Test 3: assert_numeric1 returns an error with numeric vector and several elements
  expect_snapshot(
    assert_numeric1(c(1,2)),
    error = TRUE)

  # Test 4: assert_numeric1 returns an error with character input
  expect_snapshot(
    assert_numeric1("1"),
    error = TRUE)

  # Test 5: assert_numeric1 returns an error with numeric vector and several elements and scalar as true
  expect_snapshot(
    assert_numeric1(c(1,2),scalar = TRUE),
    error = TRUE)
})

test_that("assert_named_list function works as expected", {
  # Test 1: Input is of correct class for a named list
  x <- list("elem1" = "a")
  assert_named_list(x)
  expect_true(TRUE)

  # Test 2: assert_named_list returns an error with an unnamed list
  x <- list("a")
  expect_snapshot(
    assert_named_list(x),
    error = TRUE)
})

test_that("assert_named_list_in_list function works as expected", {
  # Test 1: Input is of correct class for a named list
  x <- list(first = list("elem1" = "a"))
  assert_named_list_in_list(x)
  expect_true(TRUE)

  # Test 2: assert_named_list_in_list returns an error with a simple named list
  x <- list("elem1" = "a")
  expect_snapshot(
    assert_named_list_in_list(x),
    error = TRUE)
})

test_that("assert_grouped function works as expected", {
  # Test 1: assert_grouped returns TRUE with ungrouped data frame
  assert_grouped(iris)
  expect_true(TRUE)

  # Test 2: assert_named_list returns an error with an unnamed list
  assert_grouped(list(iris,cars))
  expect_true(TRUE)
})
