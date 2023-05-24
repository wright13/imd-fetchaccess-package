#' @title Get Access Tables
#'
#' @description This function assumes that you are using the Access Metadata Generator in your Access database, and using the fetch_from_access function. This function connects to the user designated access database and extracts the tables needed.
#'
#' @param db_path Path to your Access database
#' @param data_prefix Prefix used in your Access database to indicate data export tables and/or queries
#' @param lookup_prefix Prefix used in your Access database to indicate lookup tables
#'
#' @return A nested list containing three lists of tibbles: data, lookups, and metadata.
#' @export
#' @import RODBC
#' @import dplyr
#' @importFrom magrittr %>%
#'

get_access_tables <- function(db_path, data_prefix, lookup_prefix){

  # Load datasets for use. Pulls directly from Access database back end
  connection <- RODBC::odbcConnectAccess2007(db_path)

  metadata_prefix <- c("tsys_", "qsys_")  # Prefixes of metadata queries/tables
  data_search_string <- paste0("(^", data_prefix, ".*)", collapse = "|")  # Regex to match data table names
  lookup_search_string <- paste0("(^", lookup_prefix, ".*)", collapse = "|")  # Regex to match lookup table names
  metadata_search_string <- paste0("(^", metadata_prefix, ".*)", collapse = "|")  # Regex to match metadata table names
  # Regular expression to match all table names
  table_search_string <- paste0(c(data_search_string, lookup_search_string, metadata_search_string), collapse = "|")

  # Get names of tables to import and omit those we don't
  tables <- RODBC::sqlTables(connection) %>%
    dplyr::filter(TABLE_TYPE %in% c("TABLE", "VIEW"),
                  grepl(table_search_string, TABLE_NAME),
                  !(TABLE_NAME %in% tables_to_omit))

  data_tables <- tables$TABLE_NAME[grepl(data_search_string, tables$TABLE_NAME)]
  lookup_tables <- tables$TABLE_NAME[grepl(lookup_search_string, tables$TABLE_NAME)]
  metadata_tables <- tables$TABLE_NAME[grepl(metadata_search_string, tables$TABLE_NAME)]

  # Import data and rename tables without prefixes
  data <- sapply(data_tables, fetch_and_tidy, connection = connection)
  names(data) <- stringr::str_remove(data_tables, paste0("(", data_prefix, ")", collapse = "|"))

  lookups <- sapply(lookup_tables, fetch_and_tidy, connection = connection)

  metadata <- sapply(metadata_tables, fetch_and_tidy, connection = connection)
  names(metadata) <- stringr::str_remove(metadata_tables, paste0("(", metadata_prefix, ")", collapse = "|"))

  # Close db connection
  RODBC::odbcCloseAll()

  tables <- list(data = data,
                 lookups = lookups,
                 metadata = metadata)

  return(tables)
}
