# tablexlsx (WIP)

This release includes :

* `asTable` argument is now set to FALSE in function `add_table()` and `toxlsx()` #24
* Slight changes to API in order to solve #19, #21 and #26 

The main changes are that `mergecol` can now be used with multiple data frames, and that we can only supply character arguments instead of lists in the following cases:
1. There is only one data.frame
2. The value of the argument is to be the same for all data frames (e.g. same footnote)

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
