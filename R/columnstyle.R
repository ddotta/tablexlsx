columnstyle <- function(object,ListStyle) {

  # check if ListStyle is a list
  assert_class(ListStyle, "list")
  # check if ListStyle is a named list
  assert_named_list(ListStyle)
  # check if ListStyle is a list composed of lists of which all elements are named
  assert_named_list_in_list(ListStyle)

  return(ListStyle)

}
