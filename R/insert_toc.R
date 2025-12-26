#' @name insert_toc
#'
#' @title Insert Table of Contents to an existing workbook
#'
#' @description This function takes an existing workbook and a list or vector
#'   of sheet names, and creates a Table of Contents sheet with hyperlinks.
#'
#' @param wb Workbook object to modify
#' @param sheets A named list in the format list("dataframe" = "sheet_name")
#'
#' @seealso [toxlsx()]
#'
#' @return Nothing, this function is called for its side effects
#' @keywords internal
insert_toc <- function(wb, sheets) {
  # This should never throw due to it being sent from toxlsx
  if (!is.list(sheets)) {
    stop("sheets must be a named list of worksheet names.")
  }

  # Flatten named list
  sheets <- unlist(sheets, use.names = FALSE)

  if (!length(sheets)) return(invisible(NULL))

  # Fill workbook
  toc_sheet <- "Table of Contents"
  openxlsx::addWorksheet(wb, toc_sheet)

  links <- vapply(
    sheets,
    \(s) openxlsx::makeHyperlinkString(sheet = s, row = 1, col = 1, text = s),
    character(1)
  )

  openxlsx::writeData(
    wb,
    sheet = toc_sheet,
    startRow = 1,
    startCol = 1,
    x = "Sheet Index"
  )

  openxlsx::addStyle(
    wb,
    sheet = toc_sheet,
    style = xls_theme_default()$title,
    rows = c(1, seq_along(links) + 1),
    cols = 1,
    gridExpand = FALSE
  )

  openxlsx::setColWidths(
    wb,
    sheet = toc_sheet,
    cols = 1,
    widths = "auto"
  )

  for (i in seq_along(links)) {
    openxlsx::writeFormula(
      wb,
      sheet = toc_sheet,
      startRow = i + 1,
      startCol = 1,
      x = links[i]
    )
  }

  invisible(NULL)
}
