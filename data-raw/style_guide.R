######################################################################%#
#### Code to create the rds file `style_guide.rda in `inst/extdata`####
####################################################################%#

## Style's definitions
style_guide <- list(
  # For title
  style_title = openxlsx::createStyle(fontSize = 16, textDecoration = c("bold")),
  # For footnote1
  style_footnote1 = openxlsx::createStyle(fontSize = 12),
  # For footnote2
  style_footnote2 = openxlsx::createStyle(fontSize = 12, textDecoration = c("bold")),
  # For footnote3
  style_footnote3 = openxlsx::createStyle(fontSize = 12, textDecoration = c("bold")),
  # For column headers
  style_col_header = openxlsx::createStyle(fontSize = 12,
                                            textDecoration = c("bold"),
                                            border = c("top", "bottom", "left", "right"),
                                            borderStyle = "thin",
                                            wrapText = TRUE,
                                            halign = "center"),

  # For column headers
  style_col_header = openxlsx::createStyle(fontSize = 12,
                                              textDecoration = c("bold"),
                                              border = c("top", "bottom", "left", "right"),
                                              borderStyle = "thin",
                                              wrapText=TRUE,
                                              halign = "center"),

  # For cells in character format
  char = openxlsx::createStyle(fontSize = 12,
                                  border = c("top", "bottom", "left", "right"),
                                  borderStyle = "thin"),

  # For cells in number format (with thousands separator)
  number = openxlsx::createStyle(fontSize = 12,
                                      numFmt = "### ### ### ##0",
                                      border = c("top", "bottom", "left", "right"),
                                      borderStyle = "thin"),

  # For cells in percentage format (centered)
  percent = openxlsx::createStyle(fontSize = 12,
                                        numFmt = "#0.0",
                                        border = c("top", "bottom", "left", "right"),
                                        borderStyle = "thin",
                                        halign = "center")

)


usethis::use_data(style_guide, overwrite = TRUE)
