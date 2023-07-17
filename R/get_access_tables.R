#' @title Get Access Tables
#'
#' @description This function assumes that you are using the Access Metadata Generator in your Access database, and using the fetch_from_access function. It also assumes that the user has 64bit Access, 64 bit R installed, and the database version is newer than 2007 and not password protected. It will not work with 32 bit MS Access or R.  This function connects to the user designated access database and extracts the export data, lookup, and metadata tables needed.
#'
#' @param db_path Path to your Access database
#' @param data_prefix Prefix used in your Access database to indicate data export tables and/or queries
#' @param lookup_prefix Prefix used in your Access database to indicate lookup tables
#' @param tables_to_omit a list of any tables to omit from the import
#'
#' @return A nested list containing three lists of tibbles: data, lookups, and metadata.
#' @export
#'

get_access_tables <- function(db_path, data_prefix, lookup_prefix, tables_to_omit){

  # Checks that the database path is valid
  if(file.exists(db_path)){
    # opens DB connection
    connection <- RODBC::odbcConnectAccess2007(db_path)
  } else{
    cli::cli_abort(c(
      x = "Failed to find database at '{db_path}'.",
      "!" = "Please make sure that the file exists in the folder",
      "!" = "and/or that the variable for the database name has exactly the same spelling as the file."))
    }

  # Load datasets for use. Pulls directly from Access database back end-
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
