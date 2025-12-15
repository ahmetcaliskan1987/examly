# examly 0.2.0

* **New Feature:** Added "Select All" / "Deselect All" links for faster item type classification in the Shiny interface.
* **New Feature:** Added a "Distribute Weights" button to automatically equalize item weights out of 100.
* **Enhancement:** Split discrimination statistics into two distinct columns: Correlation-based ($r_{jx}$) and Upper/Lower 27% difference ($D$).
* **Enhancement:** Added literature-based interpretation tables and color-coded comments for item difficulty and discrimination values.
* **Bug Fix:** Improved the text parsing logic for pasted answer keys to better handle spaces and separators.

# examly 0.1.2

* Fixed a critical bug (present in v0.1.1) that prevented key analysis outputs from rendering in the Shiny interface. This fix restores the following components:
    * Summary tables (most correct, incorrect, and skipped items).
    * Student count for scores 50 and above.
    * Bar chart for student score ranges.

# examly 0.1.1

* Initial release to CRAN.
