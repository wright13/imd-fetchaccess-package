#' Fetch data from Access database
#'
#' This function assumes that you are using the Access Metadata Generator in your Access database.
#'
#' @param db_path Path to your Access database
#' @param data_prefix Prefix used in your Access database to indicate data export tables and/or queries
#' @param lookup_prefix Prefix used in your Access database to indicate lookup tables
#' @param tables_to_omit Character vector of table names that match the data and/or lookup prefixes but should not be included in the data import
#' @param custom_wrangler Optional - function that takes arguments `data`, `lookups`, and `metadata`. `data` and `lookups` are lists whose names and content correspond to the data and lookup tables in the database. Names do not include prefixes. `metadata` contains a tibble of field-level metadata called `MetadataAttributes`. See qsys_MetadataAttributes in the Access database for the contents of this tibble. This function should perform any necessary data wrangling specific to your dataset and return a named list containing `data`, `lookups`, and `metadata` with contents modified as needed. Do not remove or add tibbles in `data` or `lookups` and do not modify their names. If you add, remove, or rename columns in a tibble in `data`, you must modify the contents of `metadata` accordingly. Do not modify the structure or column names of `metadata`. The structure and column names of `lookups` should also be left as-is. Typically the only necessary modification to `lookups` will be to filter overly large species lists to only include taxa that appear in the data.
#' @param save_to_files Save data and data dictionaries to files on hard drive?
#' @inheritParams writeToFiles
#'
#' @return A nested list containing three lists of tibbles: data, lookups, and metadata.
#' @export
#'
#'
db_path <-"C:\\Users\\EEdson\\OneDrive - DOI\\Projects_InProgress\\SNPL\\GOGA_SNPL\\UpdatingFetchAccess\\SnowyPlovers_BE_20230522_AccessExportTool.accdb"
custom_wrangler = wrangle_snpl
data_prefix = "qExport"
lookup_prefix = "tlu"
tables_to_omit = c()

fetchFromAccess <- function(db_path, data_prefix = "qExport", lookup_prefix = "tlu", tables_to_omit = c(), custom_wrangler, save_to_files = FALSE, data_dir = here::here("data", "final"), dictionary_dir = here::here("data", "dictionary")){
  connection <- RODBC::odbcConnectAccess2007(db_path)  # Load datasets for use. Pulls directly from Access database back end

  metadata_prefix <- c("tsys_", "qsys_")  # Prefixes of metadata queries/tables
  data_search_string <- paste0("(^", data_prefix, ".*)", collapse = "|")  # Regex to match data table names
  lookup_search_string <- paste0("(^", lookup_prefix, ".*)", collapse = "|")  # Regex to match lookup table names
  metadata_search_string <- paste0("(^", metadata_prefix, ".*)", collapse = "|")  # Regex to match metadata table names
  table_search_string <- paste0(c(data_search_string, lookup_search_string, metadata_search_string), collapse = "|")  # Regular expression to match all table names

  # Get names of tables to import
  tables <- RODBC::sqlTables(connection) %>%
    dplyr::filter(TABLE_TYPE %in% c("TABLE", "VIEW"),
                  grepl(table_search_string, TABLE_NAME),
                  !(TABLE_NAME %in% tables_to_omit)) # Tables with tlu, tsys_, or qExport prefix that we want to omit from the exported data.
  data_tables <- tables$TABLE_NAME[grepl(data_search_string, tables$TABLE_NAME)]
  lookup_tables <- tables$TABLE_NAME[grepl(lookup_search_string, tables$TABLE_NAME)]
  metadata_tables <- tables$TABLE_NAME[grepl(metadata_search_string, tables$TABLE_NAME)]

  # Import data and rename tables without prefixes
  data <- sapply(data_tables, fetchAndTidy, connection = connection)
  names(data) <- stringr::str_remove(data_tables, paste0("(", data_prefix, ")", collapse = "|"))
  lookups <- sapply(lookup_tables, fetchAndTidy, connection = connection)
  # don't remove the prefix as we need these
  #names(lookups) <- stringr::str_remove(lookup_tables, paste0("(", lookup_prefix, ")", collapse = "|"))
  metadata <- sapply(metadata_tables, fetchAndTidy, connection = connection)
  names(metadata) <- stringr::str_remove(metadata_tables, paste0("(", metadata_prefix, ")", collapse = "|"))
  RODBC::odbcCloseAll()  # Close db connection

  # Do custom data wrangling
  if (!missing(custom_wrangler)) {
    new_tables <- custom_wrangler(data, lookups, metadata)
    data <- new_tables$data
    lookups <- new_tables$lookups
    metadata <- new_tables$metadata
  }

  # --- Make data dictionaries ---
  # Tables dictionary
  tables_dict <- metadata$MetadataQueries %>%
    dplyr::mutate(tableName = stringr::str_remove(tableName, "qExport"),
                  fileName = paste0(tableName, ".csv")) %>%
    dplyr::select(tableName, fileName, tableDescription)

  # Fields dictionary
  fields_dict <- metadata$EDIT_metadataAttributeInfo %>%
    dplyr::mutate(tableName = stringr::str_remove(tableName, "qExport"),
                  class = dplyr::case_when(readonlyClass %in% c("Short Text", "Long Text", "Memo", "Text", "Yes/No", "Hyperlink") ~ "character",
                                           readonlyClass %in% c("Number", "Large Number", "Byte", "Integer", "Long Integer", "Single", "Double", "Replication ID", "Decimal", "AutoNumber", "Currency") ~ "numeric",
                                           readonlyClass == "Categorical" ~ "categorical",
                                           readonlyClass %in% c("Date/Time", "Date/Time Extended") ~ "Date",
                                    TRUE ~ "unknown")) %>%
    dplyr::select(tableName, attributeName, "attributeDefinition" = readonlyDescription , class, unit, dateTimeFormatString, missingValueCode, missingValueCodeExplanation)

  # Categories dictionary
  # lookup_fields <- metadata$MetadataRelationFields %>%
  #   dplyr::filter(tableName %in% fields_dict$sourceTable & foreignKeyName %in% fields_dict$sourceField & !is.na(foreignDescriptionName)) %>%
  #   dplyr::inner_join(fields_dict, by = c("tableName" = "sourceTable", "foreignKeyName" = "sourceField")) %>%
  #   dplyr::select(attributeName, lookupName, lookupAttributeName) %>%
  #   dplyr::inner_join(metadata$editMetadataLookupDefs, by = "lookupName") %>%
  #   unique()
  #lookup_fields <-metadata$MetadataLookupDefs

  categories_dict <- tibble::tibble()
  for (i in 1:nrow(metadata$MetadataLookupDefs)) {
    lkup <- metadata$MetadataLookupDefs[[i, "lookupTableName"]]
    lkup_code <- metadata$MetadataLookupDefs[[i, "lookupCodeField"]]
    lkup_def <- metadata$MetadataLookupDefs[[i, "LookupDescriptionField"]]
    temp <- tibble::tibble(attributeName = metadata$MetadataLookupDefs[[i, "attributeName"]],
                           code = lookups[[lkup]][[lkup_code]],
                           definition = lookups[[lkup]][[lkup_def]])
    categories_dict <- dplyr::bind_rows(categories_dict, temp)
  }

  # Put everything in a list
  all_tables <- list(data = data,
                     lookups = lookups,
                     metadata = list(tables = tables_dict,
                                     fields = fields_dict,
                                     categories = categories_dict))

  if (save_to_files) {
    writeToFiles(all_tables, data_dir, dictionary_dir)
  }

  return(all_tables)
}



#' Write data and data dictionaries to files
#'
#' @param all_tables Output of `fetchFromAccess()`
#' @param data_dir Folder to store data csv's in
#' @param dictionary_dir Folder to store data dictionaries in
#'
writeToFiles <- function(all_tables, data_dir = here::here("data", "final"), dictionary_dir = here::here("data", "dictionary")) {

  data <- all_tables$data
  tables_dict <- all_tables$metadata$tables
  fields_dict <- all_tables$metadata$fields
  categories_dict <- all_tables$metadata$categories

  # write datasets to csv
  if (!dir.exists(data_dir)) {
    dir.create(data_dir, recursive = TRUE)
  }
  lapply(names(data), function(tbl_name) {
    readr::write_csv(data[[tbl_name]],
                     here::here(data_dir, paste0(tbl_name, ".csv")),
                     na = "")
  })

  # Write dictionaries to file
  if (!dir.exists(dictionary_dir)) {
    dir.create(dictionary_dir, recursive = TRUE)
  }
  readr::write_tsv(tables_dict, here::here(dictionary_dir, "data_dictionary_tables.txt"), na = "", append = FALSE)
  readr::write_tsv(fields_dict, here::here(dictionary_dir, "data_dictionary_attributes.txt"), na = "", append = FALSE)
  readr::write_tsv(categories_dict, here::here(dictionary_dir, "data_dictionary_categories.txt"), na = "", append = FALSE)
}

#' Fetch and tidy data
#'
#' @param tbl_name Name of data table
#' @param connection Database connection object
#'
#' @return A tibble of tidy data
#'
fetchAndTidy <- function(tbl_name, connection) {
  tidy_data <- tibble::as_tibble(RODBC::sqlFetch(connection, tbl_name, as.is = TRUE)) %>%
    dplyr::mutate(dplyr::across(where(is.character), function(x) {
      x %>%
        utf8::utf8_encode() %>%  # Encode text as UTF-8 - this prevents a lot of parsing issues in R
        trimws(whitespace = "[\\h\\v]") %>%  # Trim leading and trailing whitespace
        dplyr::na_if("")  # Replace empty strings with NA
    }))
return(tidy_data)
}
