# Assertions for parameter validates (from openxlsx package)
# These should be used at the beginning of functions to stop execution early

#' @name assert_class
#'
#' @param x Object
#' @param class class to test
#' @param or_null boolean. By default FALSE
#'
#' @noRd
assert_class <- function(x, class, or_null = FALSE) {
  sx <- as.character(substitute(x))
  ok <- inherits(x, class)

  if (or_null) {
    ok <- ok | is.null(x)
    class <- c(class, "null")
  }

  if (!ok) {
    msg <- sprintf("%s must be of class %s", sx, paste(class, collapse = " or "))
    stop(msg, call. = FALSE)
  }
}

#' @name assert_character1
#'
#' @param x Object
#' @param scalar boolean. By default FALSE
#'
#' @noRd
assert_character1 <- function(x, scalar = FALSE) {
  ok <- is.character(x) && length(x) == 1L

  if (scalar) {
    ok <- ok & nchar(x) == 1L
  }

  if (!ok) {
    stop(substitute(x), " must be a character vector of length 1L", call. = FALSE)
  }
}

#' @name assert_numeric1
#'
#' @param x Object
#' @param scalar boolean. By default FALSE
#'
#' @noRd
assert_numeric1 <- function(x, scalar = FALSE) {
  msg <- paste0(substitute(x), " must be a ")
  ok <- is.numeric(x) & length(x) == 1L

  if (scalar) {
    ok <- ok && nchar(x) == 1L
    msg <- paste0(msg, "single number")
  } else {
    msg <- paste0(msg, "numeric vector of length 1L")
  }

  if (!ok) {
    stop(msg, call. = FALSE)
  }
}

#' @name assert_named_list
#'
#' @param x Object
#'
#'
#' @noRd
assert_named_list <- function(x) {
  ok <- !is.null(names(x))

  if (!ok) {
    stop(substitute(x), " must be a named list", call. = FALSE)
  }
}

#' @name assert_named_list_in_list
#'
#' @param x Object
#'
#' @noRd
assert_named_list_in_list <- function(x) {
  ok <- sapply(x, function(y) {
    !("" %in% allNames(y))
  })

  if (!all(ok)) {
    stop(substitute(x), " must be a list composed of one or more lists of which all elements must be named", call. = FALSE)
  }
}

#' @name assert_named_list
#'
#' @param x Object
#'
#' @noRd
assert_grouped <- function(x) {
  if (is.data.frame(x)) {
    ok <- isFALSE("groups" %in% names(attributes(x)))
  } else {
    ok <- !any(sapply(x, function(df) {
      "groups" %in% names(attributes(df))
    }))
  }

  if (!ok) {
    stop(substitute(x), " must not be grouped", call. = FALSE)
  }
}

#' @name assert_xls_theme
#'
#' @param x Object
#'
#' @noRd
assert_xls_theme <- function(x) {
  if (!all(vapply(x, FUN = inherits, FUN.VALUE = logical(1L), "Style"))) {
    stop(substitute(x), " must be a list of elements of class Style", call. = FALSE)
  }
  necessary_elements <- c("title", "col_header", "character", "footnote1", "footnote2", "footnote3", "mergedcell")
  missing_elements <- setdiff(necessary_elements, names(x))
  if (length(missing_elements) > 0) {
    stop(substitute(x), " must contain styles for elements ", paste(missing_elements, collapse=" "), call. = FALSE)
  }
}
