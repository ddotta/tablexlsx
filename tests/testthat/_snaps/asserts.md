# assert_class function works as expected

    Code
      assert_class(x, "matrix", TRUE)
    Error <simpleError>
      x must be of class matrix or null

# assert_character1 function works as expected

    Code
      assert_character1(c("1", "2"))
    Error <simpleError>
      c12 must be a character vector of length 1L

---

    Code
      assert_character1(1)
    Error <simpleError>
      1 must be a character vector of length 1L

# assert_numeric1 function works as expected

    Code
      assert_numeric1(c(1, 2))
    Error <simpleError>
      c must be a numeric vector of length 1L1 must be a numeric vector of length 1L2 must be a numeric vector of length 1L

---

    Code
      assert_numeric1("1")
    Error <simpleError>
      1 must be a numeric vector of length 1L

---

    Code
      assert_numeric1(c(1, 2), scalar = TRUE)
    Error <simpleError>
      c must be a single number1 must be a single number2 must be a single number

# assert_named_list function works as expected

    Code
      assert_named_list(x)
    Error <simpleError>
      x must be a named list

# assert_named_list_in_list function works as expected

    Code
      assert_named_list_in_list(x)
    Error <simpleError>
      x must be a list composed of one or more lists of which all elements must be named

