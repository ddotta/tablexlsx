toxlsx <- function(object,
                   tosheet = list(),
                   title = list(),
                   columnstyle = list(),
                   footnote1 = list(),
                   footnote2 = list(),
                   footnote3 = list(),
                   path,
                   automaticopen = TRUE) {


  # check if object is a data frame or a list
  assert_class(object, c("data.frame","list"))
  # check if tosheet is a list
  assert_class(tosheet, "list")
  # check if title is a list
  assert_class(title, "list")
  # check if columnstyle is a list
  assert_class(columnstyle, "list")
  # check if columnstyle is a named list
  assert_named_list(columnstyle)
  # check if footnote1 is a list
  assert_class(footnote1, "list")
  # check if footnote2 is a list
  assert_class(footnote2, "list")
  # check if footnote3 is a list
  assert_class(footnote3, "list")

  sc <- sys.calls()
  caller <- sc[[length(sc) - 1]]
  output_char <- lapply(as.list(caller), deparse)[[2]]
  if (grepl("list",output_char)) {
    # Remove all whitespace from output_char
    output_char <- gsub(" ", "", output_char, fixed = TRUE)
    # Extract string within parenthesis
    matches <- gsub("[\\(\\)]",
                    "",
                    regmatches(output_char,gregexpr("\\(.*?\\)",output_char))[[1]])
    # Separate the elements of the extracted string by a comma
    output_name <- unlist(strsplit(matches, ","))
  } else {
    output_name <- output_char
  }

  # Initialize an empty list for output
  # and name the elements of the output list with output_name
  output <- setNames(vector("list",length = length(output_name)),
                     output_name)

  for (df in output_name) {
    output[[df]][["sheet"]] <- tosheet[[df]]
    output[[df]][["title"]] <- title[[df]]
    output[[df]][["column"]] <- columnstyle[[df]]
    output[[df]][["footnote1"]] <- footnote1[[df]]
    output[[df]][["footnote2"]] <- footnote2[[df]]
    output[[df]][["footnote3"]] <- footnote3[[df]]
  }

  # Creation empty workbook
  wb <- openxlsx::createWorkbook()

  # Fill workbook
  for (df in output_name) {

    # Initialize empty named list to format columns
    ColumnList <- setNames(vector("list",length = length(output[[df]][["column"]])),
                           names(output[[df]][["column"]]))

    # Fill ColumnList
    for (i in 1:length(output[[df]][["column"]])) {
      ColumnList[[i]] <- style[[output[[df]][["column"]][[paste0("c",i)]]]]
    }

    add_table(
      Table = get(df),
      WbTitle = wb,
      SheetTitle = output[[df]][["sheet"]],
      TableTitle = output[[df]][["title"]],
      StartRow = 1,
      StartCol = 1,
      FormatList = ColumnList,
      HeightTableTitle = 2,
      TableFootnote1 = output[[df]][["footnote1"]],
      TableFootnote2 = output[[df]][["footnote2"]],
      TableFootnote3 = output[[df]][["footnote3"]])

  }

  openxlsx::saveWorkbook(wb,
                         file.path(
                           path,
                           "Export.xlsx"
                         ),
                         overwrite = TRUE)

  if (isTRUE(automaticopen)) {
    openxlsx::openXL(file = file.path(
      path,
      "Export.xlsx"
    ))
  }

}
