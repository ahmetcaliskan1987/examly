
<!-- README.md is generated from README.Rmd. Please edit that file -->

# examly

<!-- badges: start -->

<!-- badges: end -->

The goal of examly is to provide a Shiny-based toolkit for comprehensive
item/test analysis. It is designed to handle multiple-choice,
true-false, and open-ended questions from datasets in 1-0 or other
formats, delivering key analyses such as difficulty, discrimination,
response-option analysis, reports.

## Installation

Install the package directly from GitHub using the `devtools` package:

``` r
# First install devtools if not already installed:
install.packages("devtools")

# Then install examly from GitHub:
devtools::install_github("ahmetcaliskan1987/examly")
```

## Usage

After installation, load the library and launch the Shiny application:

``` r
library(examly)
run_app()
```
