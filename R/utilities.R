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
#' @noRd
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
#' calcskippedrow(maliste=mydf,x=1) # return 0
#' calcskippedrow(maliste=mydf,x=2) # return 150
#' calcskippedrow(maliste=mydf,x=3) # return 200 (150 rows of iris + 50 of cars)
#' calcskippedrow(maliste=mydf,x=4) # return 232
#'
#' @noRd
calcskippedrow <- function(mylist,x) {

  res <- 0
  if (x>1) {
    for (i in 1:(x-1)) {
      res <- res + nrow(mylist[[i]])
    }
  }
  return(res)

}
