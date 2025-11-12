## Resubmission: examly 0.1.2

This is a resubmission for `examly` version 0.1.2.

### Reason for Update

The previous version on CRAN (v0.1.1) contained a critical bug. This bug prevented several key analysis outputs from being rendered or displayed correctly in the Shiny interface.

This update (v0.1.2) specifically fixes this bug. The key features that were previously blocked (and are now working correctly) include:
* Summary tables for the most correct, incorrect, and skipped questions.
* The count of students scoring 50 or above.
* The bar chart visualising student counts within specific score ranges.

This update is essential for the package to function as intended.

### R CMD check results

I have checked the package locally using `devtools::check()` (0 errors, 0 warnings, 1 note) and on win-builder (`R-devel`) using `devtools::check_win_devel()` (0 errors, 0 warnings, 1 note).

The NOTE on my local check was:
* `unable to verify current time`
    (This is a known benign note on my local build platform.)

The NOTE on the win-builder check was:
* `checking CRAN incoming feasibility ... Days since last update: 1`
    (I understand this is an informational note for the reviewer, as this is a rapid resubmission to fix the critical bug.)

Thank you for your time.
