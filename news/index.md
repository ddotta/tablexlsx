# Changelog

## tablexlsx 1.1.1

This release includes :

- Added functionallity for `table_of_contents` argument to be passed
  onto `toxlsx` [\#33](https://github.com/ddotta/tablexlsx/issues/33)

## tablexlsx 1.1.0

CRAN release: 2024-10-16

- new method for styling tables: [`toxlsx()`](../reference/toxlsx.md)
  now accepts a `theme` argument, which has to be supplied as an object
  returned by [`xls_theme()`](../reference/xls_theme.md) functions. Some
  themes are provided by default:
  [`xls_theme_default()`](../reference/xls_theme_default.md) and
  [`xls_theme_plain()`](../reference/xls_theme_plain.md). The default
  theme has been slightly changed.
  [\#40](https://github.com/ddotta/tablexlsx/issues/40)
- provide meaningful error message if merge cols don’t exist
  [\#20](https://github.com/ddotta/tablexlsx/issues/20)
- `path` can now be supplied as a file name with full path instead of a
  directory name [\#29](https://github.com/ddotta/tablexlsx/issues/29)
- (fix) [`toxlsx()`](../reference/toxlsx.md) no longer fails when the
  `object` argument is the result of a computation
  [\#18](https://github.com/ddotta/tablexlsx/issues/18)

## tablexlsx 1.0.0

CRAN release: 2024-06-06

This release includes :

- [`add_table()`](../reference/add_table.md) and
  [`toxlsx()`](../reference/toxlsx.md) now accept a `bygroup` argument
  that splits the table into groups before writing to the sheet
  [\#23](https://github.com/ddotta/tablexlsx/issues/23)
- most arguments to [`add_table()`](../reference/add_table.md) and
  [`toxlsx()`](../reference/toxlsx.md) can now be passed as atomic
  vectors. If the first argument is a single `data.frame`, the behavior
  is the same as for a lenght-one list. If the first argument is a list
  of `data.frame`s, those arguments are recycled in order to match the
  length of the list. This change applies to the arguments `tosheet`,
  `title`, `footnoteX`, `mergecol`, `bygroup`, `groupname`
  [\#19](https://github.com/ddotta/tablexlsx/issues/19)
- when a list is passed to [`toxlsx()`](../reference/toxlsx.md), the
  `mergecol` argument can now be specified for each data.frame of the
  list [\#21](https://github.com/ddotta/tablexlsx/issues/21)
- `asTable` argument is now set to FALSE in functions
  [`add_table()`](../reference/add_table.md) and
  [`toxlsx()`](../reference/toxlsx.md)
  [\#24](https://github.com/ddotta/tablexlsx/issues/24)
- Slight changes to API in order to solve
  [\#16](https://github.com/ddotta/tablexlsx/issues/16) and
  [\#21](https://github.com/ddotta/tablexlsx/issues/21)

## tablexlsx 0.2.1

This release includes :

- Use of {cli} for alert messages to users in case of success

## tablexlsx 0.2.0

This release includes :

- Add a test to check if the data frames to be exported are grouped or
  not [\#12](https://github.com/ddotta/tablexlsx/issues/12)
- Improves unit tests for function [`toxlsx()`](../reference/toxlsx.md)
- `@importFrom` statements are now grouped in a single
  `package-tablexlsx` file for easier maintenance

## tablexlsx 0.1.0

CRAN release: 2023-02-14

- Setting up pkgdown website with first vignette
- Added [`toxlsx()`](../reference/toxlsx.md) function to convert data
  frames to excel files
- Added [`add_table()`](../reference/add_table.md) function to add data
  frames in workbook
- Added a `NEWS.md` file to track changes to the package.
