toxlsx <- function(object,
                   tosheet,
                   title,
                   columnstyle,
                   path,
                   automaticopen = TRUE) {


  # check if object is a data frame or a list
  assert_class(object, c("data.frame","list"))
  # check if tosheet is a list
  assert_class(tosheet, "list")
  # check if columnstyle is a list
  assert_class(columnstyle, "list")
  # check if columnstyle is a named list
  assert_named_list(columnstyle)

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
  }

  # Creation empty workbook
  wb <- openxlsx::createWorkbook()

  # Fill workbook
  for (df in output_name) {

    ColumnList <- setNames(vector("list",length = length(output[[df]][["column"]])),
                           names(output[[df]][["column"]]))

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
      Footnote1 = "titi",
      Footnote2 = "tata",
      Footnote3 = "tutu")

  }

  openxlsx::saveWorkbook(wb,
                         file.path(
                           path,
                           "export.xlsx"
                         ),
                         overwrite = TRUE)

  if (isTRUE(automaticopen)) {
    openxlsx::openXL(file = file.path(
      path,
      "export.xlsx"
    ))
  }

}
