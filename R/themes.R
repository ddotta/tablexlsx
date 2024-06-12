xls_theme <- function(title,
                      col_header,
                      character,
                      footnote1,
                      footnote2,
                      footnote3,
                      mergedcell,
                      ...) {
  theme <- structure(list(title = title,
                col_header = col_header,
                character = character,
                footnote1 = footnote1,
                footnote2 = footnote2,
                footnote3 = footnote3,
                mergedcell = mergedcell,
                ...),
                class = "xls_theme")
  assert_xls_theme(theme)
  return(theme)
}

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

xls_theme_default = function(
    title = openxlsx::createStyle(fontSize = 16, textDecoration = c("bold")),
    # For footnote1
    footnote1 = openxlsx::createStyle(fontSize = 12),
    # For footnote2
    footnote2 = openxlsx::createStyle(fontSize = 12, textDecoration = "bold"),
    # For footnote3
    footnote3 = openxlsx::createStyle(fontSize = 12, textDecoration = "bold"),
    # For column headers
    col_header = openxlsx::createStyle(
      fontSize = 12,
      textDecoration = c("bold"),
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
