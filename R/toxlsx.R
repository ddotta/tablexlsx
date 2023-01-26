#' @name toxlsx
#'
#' @title Convert R data frames to excel files
#'
#' @description This is a function called "toxlsx" that takes in several parameters
#' including an "object" (which must be a data frame or list),
#' "tosheet", "title", "columnstyle", "footnote1", "footnote2", and "footnote3"
#' (all of which must be lists), and "path" and "automaticopen".
#' The function starts by performing several assertions to ensure that the
#' input parameters are of the correct type and format. It then creates an
#' empty list called "output" and names its elements using the "output_name"
#' variable. The function then loops through each element in "output_name"
#' and assigns values to several elements within the "output" list.
#' The function then creates an empty workbook and fills it with the data
#' from the "output" list. The function also includes some additional logic
#' for formatting columns in the workbook.
#'
#' @param object data.frame or list to be converted to excel
#' @param tosheet list of sheet names for each element of object
#' @param title list of title for each element of object
#' @param columnstyle list of style for each element of object
#' @param footnote1 list of footnote1 for each element of object
#' @param footnote2 list of footnote2 for each element of object
#' @param footnote3 list of footnote3 for each element of object
#' @param path path to save excel file
#' @param automaticopen logical indicating if excel file should open automatically
#'
#' @return an excel file
#'
toxlsx <- function(object,
                   tosheet = list(),
                   title = list(),
                   columnstyle = list("default" = NULL),
                   footnote1 = list(),
                   footnote2 = list(),
                   footnote3 = list(),
                   path,
                   automaticopen = TRUE) {

  # check if object is a data frame or a list
  assert_class(object, c("data.frame", "list"))
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

  # Get the function call of the last function call
  sc <- sys.calls()
  caller <- sc[[length(sc) - 1]]
  # Convert the function call to a character vector
  output_char <- lapply(as.list(caller), deparse)[[2]]

  # Check if the output_char contains the word "list"
  if (grepl("list", output_char)) {
    # Remove all whitespace from output_char
    output_char <- gsub(" ", "", output_char, fixed = TRUE)
    # Extract string within parenthesis
    matches <- gsub(
      "[\\(\\)]",
      "",
      regmatches(output_char, gregexpr("\\(.*?\\)", output_char))[[1]]
    )
    # Separate the elements of the extracted string by a comma
    output_name <- unlist(strsplit(matches, ","))
  } else {
    output_name <- output_char
  }

  # Initialize an empty list for output
  # and name the elements of the output list with output_name
  output <- setNames(
    vector("list", length = length(output_name)),
    output_name
  )

  # Initialize also an empty list for Sheetslist
  Sheetslist <- output

  # Loop through each element in output_name
  for (df in output_name) {
    output[[df]][["sheet"]] <-
      if (length(tosheet) == 0) {
        # If tosheet is not provided, use "Sheet #" as the sheet name
        paste0("Sheet ", as.character(which(output_name == df)))
      } else {
        # Else use df as the sheet name
        tosheet[[df]]
      }
    Sheetslist[which(output_name == df)] <- output[[df]][["sheet"]]
    output[[df]][["title"]] <- if (length(title) == 0) df else title[[df]]
    output[[df]][["column"]] <- if (paste(names(columnstyle), collapse = "") %in% c("default")) list() else columnstyle[[df]]
    output[[df]][["footnote1"]] <- if (length(footnote1) == 0) "" else footnote1[[df]]
    output[[df]][["footnote2"]] <- if (length(footnote2) == 0) "" else footnote2[[df]]
    output[[df]][["footnote3"]] <- if (length(footnote3) == 0) "" else footnote3[[df]]
  }

  # Creation empty workbook
  wb <- openxlsx::createWorkbook()

  # Fill workbook
  for (df in output_name) {
    # If argument columnstyle is not filled in the function
    if (paste(names(columnstyle), collapse = "") %in% c("default")) {
      # Initialize empty named list to format columns (ColumnList)
      ColumnList <- as.list(setNames(
        rep("character", length(names(get(df)))),
        paste0("c", 1:length(names(get(df))))
      ))

      # Fill ColumnList
      for (i in 1:length(ColumnList)) {
        ColumnList[[i]] <- style[[ColumnList[[paste0("c", i)]]]]
      }
    # Else if argument columnstyle is filled in the function
    } else {
      # Initialize empty named list to format columns (ColumnList)
      ColumnList <- setNames(
        vector("list", length = length(output[[df]][["column"]])),
        names(output[[df]][["column"]])
      )

      # Fill ColumnList
      for (i in 1:length(output[[df]][["column"]])) {
        ColumnList[[i]] <- style[[output[[df]][["column"]][[paste0("c", i)]]]]
      }
    }

    # If at least two sheets are filled in tosheet argument
    if (length(tosheet) > 1) {
      # We split this tosheet list according to each distint element
      listsplitted <- split(tosheet, match(tosheet, unique(unlist(tosheet))))
      # Each element of listsplitted is named with the name of each processed sheet
      listsplitted <- setNames(listsplitted,
                               unique(unlist(tosheet)))
      # We create namecurrentsheet as the name where df must be written
      namecurrentsheet <- names(listsplitted[[tosheet[[df]]]])
      # We create multipledfinsheet of the same length as listsplitted
      #  made up of booleans that indicate whether several df are in the same sheet
      multipledfinsheet <- lapply(listsplitted, function(x) length(x)>1)
    }

    # Useful print to debug
    # print(output_name)
    # print(listsplitted)
    # print(multipledfinsheet[[tosheet[[df]]]])
    # print(calcstartrow(which(namecurrentsheet == df)))

    # Use add_table() function to add each df in workbook
    add_table(
      Table = get(df),
      WbTitle = wb,
      SheetTitle = output[[df]][["sheet"]],
      TableTitle = output[[df]][["title"]],
      StartRow =
        # If the "tosheet" argument is not filled in
        # or if only one sheet is filled in "tosheet",
        # then we necessarily start writing from line 1
        if (length(tosheet) == 0 | length(tosheet) == 1) {
          1
          # Else if at least two sheets are filled in "tosheet" argument
          # and each df must be in different sheets
        } else if (length(tosheet) > 1 & isFALSE(multipledfinsheet[[tosheet[[df]]]])) {
          1
          # Else if at least two sheets are filled in "tosheet" argument
          # and at least two df must be in a same sheet
        } else if (length(tosheet) > 1 & isTRUE(multipledfinsheet[[tosheet[[df]]]])) {
          # StartRow is equal to 1 for first df
          # StartRow is equal to 11 + nrow(first df) for second df
          # StartRow is equal to 21 + nrow(first df) + nrow(second df) for third df
          calcstartrow(which(namecurrentsheet == df)) + calcskippedrow(mylist = object, x = which(namecurrentsheet == df))
        },
      StartCol = 1,
      FormatList = ColumnList,
      HeightTableTitle = 2,
      TableFootnote1 = output[[df]][["footnote1"]],
      TableFootnote2 = output[[df]][["footnote2"]],
      TableFootnote3 = output[[df]][["footnote3"]]
    )
  }

  print(file.path(
    path,
    "Export.xlsx"
  ))

  # Save workbook
  openxlsx::saveWorkbook(
    wb,
    file.path(
      path,
      "Export.xlsx"
    ),
    overwrite = TRUE
  )

  # Open workbook automatically if automaticopen is TRUE
  if (isTRUE(automaticopen)) {
    openxlsx::openXL(file = file.path(
      path,
      "Export.xlsx"
    ))
  }
}
