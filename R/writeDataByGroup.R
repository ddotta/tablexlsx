#' WriteData by groups
#'
#' @param wb A Workbook object
#' @param sheet The worksheet to write to
#' @param x A data.frame
#' @param group Character vector indicating the name of the columns by which to group
#' @param groupsep Character vector that separates modalities if multiple columns in group
#' @param groupname Boolean indicating whether the name of the grouping variable should be written
#' @param groupequal Character vector that separates the name and the modalities of grouping variable (if groupname = TRUE)
#' @param startRow The starting row to write to
#' @param ... Additional arguments passed to openxlsx::writeData()
#'
#' @return A Workbook in which data has been written by group
#'
writeDataByGroup <- function(
    wb,
    sheet,
    x,
    group = ".group",
    groupsep = " / ",
    groupname = FALSE,
    groupequal = " = ",
    startRow = 1,
    ...
) {
  table_groups <- split(x[, !names(x) %in% group], x[group], sep = groupsep)

  s <- startRow
  openxlsx::writeData(wb, sheet, x[0, !names(x) %in% group], startRow = s, ...)
  s <- s + 1
  for (g in names(table_groups)) {
    groupindic <- if (groupname) paste0(paste(group, collapse = groupsep), groupequal, g) else g
    openxlsx::writeData(wb, sheet, groupindic, startRow = s, ...)
    openxlsx::writeData(wb, sheet, table_groups[[g]], startRow = s+1, colNames = FALSE, ...)
    s <- s + 1 + nrow(table_groups[[g]])
  }

  return(wb)
}
