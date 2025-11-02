#' Load translation dictionary
#'
#' Finds `<lang>.json` under the installed package's `i18n/` folder and,
#' if not found (during development), falls back to `inst/i18n/<lang>.json`.
#'
#' @param lang Character scalar language code. Currently `"tr"` or `"en"`.
#' @return A named list (key -> string) parsed from the JSON file.
#' @examples
#' \dontrun{
#' d <- i18n_load("tr")
#' i18n_t(d, "ui.title", "Başlık")
#' }
#' @export
#' @importFrom jsonlite fromJSON
i18n_load <- function(lang = "tr") {
  stopifnot(is.character(lang), length(lang) == 1L)
  if (!lang %in% c("tr", "en")) {
    stop("Unsupported language: ", lang)
  }
  fn <- paste0(lang, ".json")
  # installed package path
  path <- system.file("shinyapp", "i18n", fn, package = "SMART")
  # dev fallback
  if (!nzchar(path)) path <- file.path("inst", "shinyapp", "i18n", fn)
  if (!file.exists(path)) {
    stop("i18n not found: ", fn, " (looked in ", path, ")")
  }
  jsonlite::fromJSON(path, simplifyVector = TRUE)
}

#' Translate a UI/message key
#'
#' Returns the value for `key` from a dictionary produced by [i18n_load()].
#' If the key is missing, returns `default` when provided, otherwise the key itself.
#'
#' @param dict Named list produced by [i18n_load()].
#' @param key Character scalar; lookup key.
#' @param default Optional fallback value if the key is not present.
#' @return Character scalar.
#' @examples
#' \dontrun{
#' d <- i18n_load("en")
#' i18n_t(d, "buttons.download", "Download")
#' }
#' @export
i18n_t <- function(dict, key, default = NULL) {
  if (!is.list(dict)) stop("`dict` must be a list returned by i18n_load()")
  if (!is.character(key) || length(key) != 1L) stop("`key` must be a character scalar")
  val <- dict[[key]]
  if (!is.null(val)) return(val)
  if (!is.null(default)) return(default)
  key
}
