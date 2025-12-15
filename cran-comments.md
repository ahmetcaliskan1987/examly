## Resubmission

This is a resubmission for `examly` version 0.2.0.

## Description of changes
This is a minor update including UI enhancements and statistical refinements:
* Added UI helpers: 'Select All' for items and 'Distribute Weights' button.
* Refined Statistics: Separated discrimination metrics (Point-biserial vs. UL27) and added interpretation guidelines based on literature.
* Fixed a bug in parsing pasted answer keys with irregular spacing.
* Updated DESCRIPTION and NEWS.md.

### R CMD check results

We have checked the package locally using `devtools::check()` (0 errors, 0 warnings, 0 note) and on win-builder (`R-devel`) using `devtools::check_win_devel()` (0 errors, 0 warnings, 1 note).

The NOTE on the win-builder check was:
* `checking CRAN incoming feasibility ... Days since last update: 1`
    (We understand this is an informational note for the reviewer, as this is a rapid resubmission to fix the critical bug.)

Thank you for your time.
