################################################################%#
#### Code to create the rds file `style.rda in `inst/extdata`####
##############################################################%#

## Style's definitions
style <- list(
  # For title
  title = openxlsx::createStyle(fontSize = 16, textDecoration = c("bold")),
  # For footnote1
  footnote1 = openxlsx::createStyle(fontSize = 12),
  # For footnote2
  footnote2 = openxlsx::createStyle(fontSize = 12, textDecoration = c("bold")),
  # For footnote3
  footnote3 = openxlsx::createStyle(fontSize = 12, textDecoration = c("bold")),
  # For column headers
  col_header = openxlsx::createStyle(fontSize = 12,
                                            textDecoration = c("bold"),
                                            border = c("top", "bottom", "left", "right"),
                                            borderStyle = "thin",
                                            wrapText = TRUE,
                                            halign = "center"),

  # For cells in character format
  character = openxlsx::createStyle(fontSize = 12,
                                    border = c("top", "bottom", "left", "right"),
                                    borderStyle = "thin"),

  # For cells in number format (with thousands separator)
  number = openxlsx::createStyle(fontSize = 12,
                                 numFmt = "### ### ### ##0",
                                 border = c("top", "bottom", "left", "right"),
                                 borderStyle = "thin"),

  # For cells in number format with decimals (and with thousands separator)
  decimal = openxlsx::createStyle(fontSize = 12,
                                  numFmt = "### ### ### ##0.0",
                                  border = c("top", "bottom", "left", "right"),
                                  borderStyle = "thin"),

  # For cells in percentage format (centered)
  percent = openxlsx::createStyle(fontSize = 12,
                                        numFmt = "#0.0",
                                        border = c("top", "bottom", "left", "right"),
                                        borderStyle = "thin",
                                        halign = "center")

)


usethis::use_data(style, overwrite = TRUE)
