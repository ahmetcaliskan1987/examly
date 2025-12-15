## Test environments
* local Windows 11 install, R 4.4.2
* win-builder (devel and release)

## Reverse dependencies
There are no reverse dependencies.

## Description of changes
This is a minor update (0.2.0) including UI enhancements and statistical refinements:
* Added UI helpers: 'Select All' for items and 'Distribute Weights' button.
* Refined Statistics: Separated discrimination metrics (Point-biserial vs. UL27) and added interpretation guidelines based on literature.
* Fixed a bug in parsing pasted answer keys with irregular spacing.
* Updated DESCRIPTION and NEWS.md.

### R CMD check results

We have checked the package locally using `devtools::check()` (0 errors, 0 warnings, 0 note) and on win-builder (`R-devel`) using `devtools::check_win_devel()` (0 errors, 0 warnings, 0 note).

Thank you for your time.
