#' @name .onLoad
#' @import utils
#' @noRd
.onLoad <- function(libname, pkgname) {
  # set global variables in order to avoid CHECK notes
  utils::globalVariables(c("style"))

  invisible()
}
