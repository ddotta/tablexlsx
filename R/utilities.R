#' @name calcstartrow
#'
#' @title Utility function that returns start row to correctly separate df written in a same sheet
#'
#' @param x a number
#'
#' @examples
#' calcstartrow(1) # return 1
#' calcstartrow(2) # return 11
#' calcstartrow(3) # return 21
#'
#' @keywords internal
calcstartrow <- function(x) {

  if (x==1) {
    res <- 1
  } else {
    res <- (x*10) - 9
  }
  return(res)

}

#' @name calcskippedrow
#'
#' @title Utility function that returns the number of rows that need to be skipped
#' to correctly account for previous data frames already present in the sheet
#'
#' @param mylist a list of data frames
#' @param x a number
#'
#' @examples
#' mydf <- list(iris,cars,mtcars)
#' calcskippedrow(mylist=mydf,x=1) # return 0
#' calcskippedrow(mylist=mydf,x=2) # return 150
#' calcskippedrow(mylist=mydf,x=3) # return 200 (150 rows of iris + 50 of cars)
#' calcskippedrow(mylist=mydf,x=4) # return 232
#'
#' @keywords internal
calcskippedrow <- function(mylist,x) {

  res <- 0
  if (x>1) {
    for (i in 1:(x-1)) {
      res <- res + nrow(mylist[[i]])
    }
  }
  return(res)

}

#' @name get_indices_of_identical_elements
#'
#' @title Utility function that takes a vector as input and returns
#' the indices of the elements in the vector that are identical.
#'
#' @param vector a vector
#'
#' @examples
#' myvector <- c("momo","momo","momo","mumu","mumu")
#' get_indices_of_identical_elements(myvector)
#' #' # Output:
#' # [1] "1:3" "4:5"
#'
#' @keywords internal
get_indices_of_identical_elements <- function(vector) {
  res <- c()
  start <- 1
  for (i in 2:length(vector)) {
    if (vector[i] != vector[i-1]) {
      end <- i - 1
      res <- c(res, paste(start, end, sep = ":"))
      start <- i
    }
  }
  end <- length(vector)
  res <- c(res, paste(start, end, sep = ":"))
  res
}

#' @name get_indices_from_vector
#'
#' @title Utility function that takes a vector as input and returns
#' the indices of the first and last element.
#'
#' @param vector a vector
#'
#' @examples
#' myvector <- c("momo","momo","momo","mumu","mumu")
#' get_indices_from_vector(myvector)
#' #' # Output:
#' # [1] "1:5"
#'
#' @keywords internal
get_indices_from_vector <- function(vector) {
  res <- c(paste0("1:",as.character(length(vector))))
  res
}

#' @name convert_range_string
#'
#' @title Utility function that takes converts a string representing
#' a range of numbers (e.g. "1:3") into a vector of numbers
#'
#' @param range_string A string representing a range of numbers
#'
#' @return A vector of numbers
#'
#' @examples
#' convert_range_string("1:3")
#' #' # Output:
#' # [1] 1 2 3
#'
#' @keywords internal
convert_range_string <- function(range_string) {
  range <- as.numeric(unlist(strsplit(range_string, ":")))
  seq(range[1], range[2])
}
