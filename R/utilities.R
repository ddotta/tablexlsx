#' @name calcstartrow
#'
#' @title Utility function that returns start row for writing df
#'
#' @description This function returns start row for writing df in the
#' case where several df are in the same sheet (see `toxlsx()`)
#'
#' @param x file format
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
