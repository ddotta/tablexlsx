# tablexlsx 1.1.0

* new method for styling tables: `toxlsx()` now accepts a `theme` argument, which has to be supplied as an object returned by `xls_theme()` functions. Some themes are provided by default: `xls_theme_default()` and `xls_theme_plain()`. The default theme has been slightly changed. #40
* provide meaningful error message if merge cols don't exist #20
* `path` can now be supplied as a file name with full path instead of a directory name #29
* (fix) `toxlsx()` no longer fails when the `object` argument is the result of a computation #18

# tablexlsx 1.0.0

This release includes :


* `add_table()` and `toxlsx()` now accept a `bygroup` argument that splits the table into groups before writing to the sheet #23
* most arguments to `add_table()` and `toxlsx()` can now be passed as atomic vectors. If the first argument is a single `data.frame`, the behavior is the same as for a lenght-one list. If the first argument is a list of `data.frame`s, those arguments are recycled in order to match the length of the list. This change applies to the arguments `tosheet`, `title`, `footnoteX`, `mergecol`, `bygroup`, `groupname` #19
* when a list is passed to `toxlsx()`, the `mergecol` argument can now be specified for each data.frame of the list #21
* `asTable` argument is now set to FALSE in functions `add_table()` and `toxlsx()` #24
* Slight changes to API in order to solve #16 and #21 

# tablexlsx 0.2.1

This release includes :

* Use of {cli} for alert messages to users in case of success

# tablexlsx 0.2.0

This release includes :

* Add a test to check if the data frames to be exported are grouped or not #12
* Improves unit tests for function `toxlsx()`
* `@importFrom` statements are now grouped in a single `package-tablexlsx` file
for easier maintenance

# tablexlsx 0.1.0

* Setting up pkgdown website with first vignette
* Added `toxlsx()` function to convert data frames to excel files
* Added `add_table()` function to add data frames in workbook
* Added a `NEWS.md` file to track changes to the package.
