#' @title Fetch and Tidy Data
#' @description Fetches Access tables and performs some basic cleaning of white space and NA strings
#'
#' @param tbl_name Name of data table
#' @param connection Database connection object
#'
#' @return A tibble of tidy data
#'
fetch_and_tidy <- function(tbl_name, connection) {
  tidy_data <- tibble::as_tibble(RODBC::sqlFetch(connection, tbl_name, as.is = TRUE)) %>%
    dplyr::mutate(dplyr::across(where(is.character), function(x) {
      x %>%
        utf8::utf8_encode() %>%  # Encode text as UTF-8 - this prevents a lot of parsing issues in R
        trimws(whitespace = "[\\h\\v]") %>%  # Trim leading and trailing whitespace
        dplyr::na_if("")  # Replace empty strings with NA
    }))
  return(tidy_data)
}
