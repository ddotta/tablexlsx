#' @name xls_theme
#'
#' @title Constructor function for xls themes
#'
#' @description
#' This function creates an xls theme for styling exported tables.
#' All its arguments must be `openxlsx` Style objects.
#'
#' @param title Style for the title
#' @param col_header Style for the columns header
#' @param character Default style for data cells
#' @param footnote1 Style for footnote1
#' @param footnote2 Style for footnote2
#' @param footnote3 Style for footnote3
#' @param mergedcell Style for merged cells
#' @param ... Other (named) custom styles
#'
#' @return a named list of class xls_theme, whose elements are `openxlsx` Style objects.
#' @export
#'
#' @seealso \code{\link[tablexlsx:xls_theme_plain]{xls_theme_plain()}},
#' \code{\link[tablexlsx:xls_theme_default]{xls_theme_default()}}
#'
#' @examples
#' my_theme <- xls_theme(
#'   title = openxlsx::createStyle(),
#'   col_header = openxlsx::createStyle(),
#'   character = openxlsx::createStyle(),
#'   footnote1 = openxlsx::createStyle(),
#'   footnote2 = openxlsx::createStyle(),
#'   footnote3 = openxlsx::createStyle(),
#'   mergedcell = openxlsx::createStyle()
#' )
#'
#' \dontrun{
#' toxlsx(object = iris, path = tempdir(), theme = my_theme)
#' }
xls_theme <- function(title,
                      col_header,
                      character,
                      footnote1,
                      footnote2,
                      footnote3,
                      mergedcell,
                      ...) {
  theme <- list(title = title,
                col_header = col_header,
                character = character,
                footnote1 = footnote1,
                footnote2 = footnote2,
                footnote3 = footnote3,
                mergedcell = mergedcell,
                ...)
  class(theme) <- "xls_theme"
  assert_xls_theme(theme)
  theme
}

#' @name xls_theme_plain
#'
#' @title Constructor function for a plain xls theme
#'
#' @description
#' This function is a wrapper around [xls_theme()] that creates an xls theme for styling exported tables.
#' It defines a simple theme whith no special formatting.
#' All its arguments must be `openxlsx` Style objects.
#'
#' @param title Style for the title
#' @param col_header Style for the columns header
#' @param character Default style for data cells
#' @param footnote1 Style for footnote1
#' @param footnote2 Style for footnote2
#' @param footnote3 Style for footnote3
#' @param mergedcell Style for merged cells
#' @param ... Other (named) custom styles
#'
#' @return a named list of class xls_theme, whose elements are `openxlsx` Style objects.
#' @export
#'
#' @seealso \code{\link[tablexlsx:xls_theme]{xls_theme()}},
#' \code{\link[tablexlsx:xls_theme_default]{xls_theme_default()}}
#'
#' @examples
#' # plain theme
#' xls_theme_plain()
#'
#' # plain theme with title in bold
#' my_theme <- xls_theme_plain(title = openxlsx::createStyle(textDecoration = "bold"))
#'
#' \dontrun{
#' toxlsx(object = iris, path = tempdir(), theme = my_theme)
#' }
xls_theme_plain = function(
    title = openxlsx::createStyle(),
    col_header = openxlsx::createStyle(),
    character = openxlsx::createStyle(),
    footnote1 = openxlsx::createStyle(),
    footnote2 = openxlsx::createStyle(),
    footnote3 = openxlsx::createStyle(),
    mergedcell = openxlsx::createStyle(),
    ...
) {
  xls_theme(
    title = title,
    col_header = col_header,
    character = character,
    footnote1 = footnote1,
    footnote2 = footnote2,
    footnote3 = footnote3,
    mergedcell = mergedcell,
    ...
  )
}

#' @name xls_theme_default
#'
#' @title Constructor function for the default xls theme
#'
#' @description
#' This function is a wrapper around [xls_theme()] that creates an xls theme for styling exported tables.
#' It defines a theme whith sensible default formatting values.
#' It also defines custom styles for "number", "decimal" and "percent column types.
#' All its arguments must be `openxlsx` Style objects.
#'
#' @param title Style for the title
#' @param col_header Style for the columns header
#' @param character Default style for data cells
#' @param number Style for columns in number format
#' @param decimal Style for columns in decimal format
#' @param percent Style for columns in percent format
#' @param footnote1 Style for footnote1
#' @param footnote2 Style for footnote2
#' @param footnote3 Style for footnote3
#' @param mergedcell Style for merged cells
#' @param ... Other (named) custom styles
#'
#' @return a named list of class xls_theme, whose elements are `openxlsx` Style objects.
#' @export
#'
#' @seealso \code{\link[tablexlsx:xls_theme]{xls_theme()}},
#' \code{\link[tablexlsx:xls_theme_plain]{xls_theme_plain()}}
#'
#' @examples
#' # default theme
#' xls_theme_default()
#'
#' # default theme with title in italic
#' my_theme <- xls_theme_default(title = openxlsx::createStyle(textDecoration = "italic"))
#'
#' \dontrun{
#' toxlsx(object = iris, path = tempdir(), theme = my_theme)
#' }
xls_theme_default = function(
    title = openxlsx::createStyle(fontSize = 16, textDecoration = "bold"),
    # For footnote1
    footnote1 = openxlsx::createStyle(fontSize = 12),
    # For footnote2
    footnote2 = openxlsx::createStyle(fontSize = 12, textDecoration = "bold"),
    # For footnote3
    footnote3 = openxlsx::createStyle(fontSize = 12, textDecoration = "bold"),
    # For column headers
    col_header = openxlsx::createStyle(
      fontSize = 12,
      textDecoration = "bold",
      border = c("top", "bottom", "left", "right"),
      borderStyle = "thin",
      wrapText = TRUE,
      halign = "center"
    ),
    # For cells in character format
    character = openxlsx::createStyle(
      fontSize = 12,
      border = c("top", "bottom", "left", "right"),
      borderStyle = "thin"
    ),
    # For cells in number format (with thousands separator)
    number = openxlsx::createStyle(
      fontSize = 12,
      numFmt = "### ### ### ##0",
      border = c("top", "bottom", "left", "right"),
      borderStyle = "thin"
    ),
    # For cells in number format with decimals (and with thousands separator)
    decimal = openxlsx::createStyle(
      fontSize = 12,
      numFmt = "### ### ### ##0.0",
      border = c("top", "bottom", "left", "right"),
      borderStyle = "thin"
    ),
    # For cells in percentage format (centered)
    percent = openxlsx::createStyle(
      fontSize = 12,
      numFmt = "#0.0",
      border = c("top", "bottom", "left", "right"),
      borderStyle = "thin",
      halign = "center"
    ),
    mergedcell = openxlsx::createStyle(
      fontSize = 12,
      border = c("top", "bottom", "left", "right"),
      borderStyle = "thin",
      wrapText = TRUE,
      valign = "center",
      halign = "center"
    ),
    ...
  ) {
  xls_theme_plain(
    title = title,
    col_header = col_header,
    character = character,
    number = number,
    decimal = decimal,
    percent = percent,
    footnote1 = footnote1,
    footnote2 = footnote2,
    footnote3 = footnote3,
    mergedcell = mergedcell,
    ...
  )
}
