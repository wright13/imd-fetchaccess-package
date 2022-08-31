#' Fetch data from Access database
#'
#' This function assumes that you are using the Access Metadata Generator in your Access database.
#'
#' @param db_path Path to your Access database
#' @param data_prefix Character vector of prefix(es) used in your Access database to indicate data export tables and/or queries.Be sure to include special characters like underscores (e.g. 'tbl_'). If you are using the `data_regex` argument, be sure to include prefixes for all matched tables.
#' @param data_regex Regular expression to match names of data tables. You can ignore this if your data table prefix(es) are specific to the tables you want to read in. If you only want to read in a subset of the tables specified by `data_prefix`, use this argument to specify only the tables you want. `data_prefix` is still required, as it is used to clean up table names.
#' @param lookup_prefix Character vector of prefix(es) used in your Access database to indicate lookup tables. Be sure to include special characters like underscores (e.g. 'tlu_'). If you are using the `lookup_regex` argument, be sure to include prefixes for all matched tables.
#' @param lookup_regex Regular expression to match names of lookup tables. You can ignore this if your lookup table prefix(es) are specific to the lookups you want to read in. If you only want to read in a subset of the lookup tables specified by `lookup_prefix`, use this argument to specify only the tables you want. This argument is especially useful if you have one or two tables labeled as 'data' tables (e.g. tbl_Sites) that act as lookups in some cases. `lookup_prefix` is still required, as it is used to clean up lookup names.
#' @param add_r_classes Include R classes in addition to EML classes?
#' @param custom_wrangler Optional - function that takes arguments `data`, `lookups`, and `metadata`. `data` and `lookups` are lists whose names and content correspond to the data and lookup tables in the database. Names do not include prefixes. `metadata` contains a tibble of field-level metadata called `MetadataAttributes`. See qsys_MetadataAttributes in the Access database for the contents of this tibble. This function should perform any necessary data wrangling specific to your dataset and return a named list containing `data`, `lookups`, and `metadata` with contents modified as needed. Do not remove or add tibbles in `data` or `lookups` and do not modify their names. If you add, remove, or rename columns in a tibble in `data`, you must modify the contents of `metadata` accordingly. Do not modify the structure or column names of `metadata`. The structure and column names of `lookups` should also be left as-is. Typically the only necessary modification to `lookups` will be to filter overly large species lists to only include taxa that appear in the data.
#' @param save_to_files Save data and data dictionaries to files on hard drive?
#' @param ... Options to pass to [writeToFiles()]
#' @inheritParams RODBC::sqlQuery
#'
#' @details See [RODBC::sqlFetch()] and [RODBC::sqlQuery()] for more information on the `as.is` argument.
#'
#' @return A nested list containing three lists of tibbles: data, lookups, and metadata.
#' @export
#'
fetchFromAccess <- function(db_path,
                            data_prefix = "qExport",
                            data_regex = paste0("(^", data_prefix, ".*)", collapse = "|"),
                            lookup_prefix = "tlu",
                            lookup_regex = paste0("(^", lookup_prefix, ".*)", collapse = "|"),
                            as.is = FALSE, add_r_classes = FALSE,
                            custom_wrangler,
                            save_to_files = FALSE, ...){
  connection <- RODBC::odbcConnectAccess2007(db_path)  # Load datasets for use. Pulls directly from Access database back end

  metadata_prefix <- c("tsys_", "qsys_")  # Prefixes of metadata queries/tables
  metadata_regex <- paste0("(^", metadata_prefix, ".*)", collapse = "|")  # Regex to match metadata table names
  table_search_string <- paste0(c(data_regex, lookup_regex, metadata_regex), collapse = "|")  # Regular expression to match all table names

  # Get names of tables to import
  tables <- RODBC::sqlTables(connection) %>%
    dplyr::filter(TABLE_TYPE %in% c("TABLE", "VIEW"),
                  grepl(table_search_string, TABLE_NAME))
  data_tables <- tables$TABLE_NAME[grepl(data_regex, tables$TABLE_NAME)]
  lookup_tables <- tables$TABLE_NAME[grepl(lookup_regex, tables$TABLE_NAME)]
  metadata_tables <- tables$TABLE_NAME[grepl(metadata_regex, tables$TABLE_NAME)]

  # Import data and rename tables without prefixes
  data <- sapply(data_tables, fetchAndTidy, connection = connection, as.is = as.is)
  names(data) <- stringr::str_remove(data_tables, paste0("(", data_prefix, ")", collapse = "|"))
  lookups <- sapply(lookup_tables, fetchAndTidy, connection = connection, as.is = as.is)
  names(lookups) <- stringr::str_remove(lookup_tables, paste0("(", lookup_prefix, ")", collapse = "|"))
  metadata <- sapply(metadata_tables, fetchAndTidy, connection = connection, as.is = as.is)
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
    dplyr::mutate(tableName = stringr::str_remove(tableName, data_prefix),
                  fileName = paste0(tableName, ".csv")) %>%
    dplyr::select(tableName, fileName, tableDescription)

  # Fields dictionary
  fields_dict <- metadata$MetadataAttributes %>%
    dplyr::mutate(tableName = stringr::str_remove(tableName, data_prefix),
                  class = dplyr::case_when(class %in% c("Short Text", "Long Text", "Memo", "Text", "Yes/No", "Hyperlink") ~ "character",
                                    class %in% c("Number", "Large Number", "Byte", "Integer", "Long Integer", "Single", "Double", "Replication ID", "Decimal", "AutoNumber", "Currency") ~ "numeric",
                                    class %in% c("Date/Time", "Date/Time Extended") ~ "Date",
                                    TRUE ~ "unknown")) %>%
    dplyr::select(tableName, attributeName, attributeDefinition, class, unit, dateTimeFormatString, missingValueCode, missingValueCodeExplanation, sourceField, sourceTable)

  # Categories dictionary

  # Get lookup table relationship info
  lookup_rels <- metadata$MetadataRelationFields %>%
    dplyr::filter(!is.na(foreignDescriptionName),
                  (lookupName %in% fields_dict$sourceTable) | ((tableName %in% fields_dict$sourceTable) & (foreignKeyName %in% fields_dict$sourceField)))

  # Case 1: lookup table is present in query and column in export query is pulling from lookup table
  lookup_fields_1 <- fields_dict %>%
    dplyr::select(sourceField, sourceTable, attributeName) %>%
    dplyr::inner_join(lookup_rels, by = c("sourceTable" = "lookupName")) %>%
    dplyr::select(sourceTable, sourceField, definitionColumnName = foreignDescriptionName, attributeName) %>%
    unique()

  # Case 2: there is a related lookup table, but the export query does not contain the lookup table and instead just pulls from the foreign key column.
  lookup_fields_2 <- fields_dict %>%
    dplyr::select(sourceField, sourceTable, attributeName) %>%
    dplyr::inner_join(lookup_rels, by = c("sourceTable" = "tableName", "sourceField" = "foreignKeyName")) %>%
    dplyr::select(sourceTable = lookupName, sourceField = lookupAttributeName, definitionColumnName = foreignDescriptionName, attributeName) %>%
    unique()

  lookup_fields <- rbind(lookup_fields_1, lookup_fields_2) %>% unique()

  categories_dict <- mapply(function(lkup, lkup_code, lkup_def, attr) {
    lkup <- stringr::str_remove(lkup, paste0("(", lookup_prefix, ")", collapse = "|"))
    df <- tibble::tibble(attributeName = attr,
                         code = as.character(lookups[[lkup]][[lkup_code]]),
                         definition = as.character(lookups[[lkup]][[lkup_def]]))
    return(df)
  }, lookup_fields$sourceTable, lookup_fields$sourceField, lookup_fields$definitionColumnName, lookup_fields$attributeName, SIMPLIFY = FALSE)

  categories_dict <- do.call(rbind, categories_dict) %>% unique()

  if (add_r_classes) {
    fields_dict <- getRClass(fields_dict, data)
  }

  # Put everything in a list
  all_tables <- list(data = data,
                     lookups = lookups,
                     metadata = list(tables = tables_dict,
                                     fields = fields_dict,
                                     categories = categories_dict))

  if (save_to_files) {
    writeToFiles(all_tables, ...)
  }

  return(all_tables)
}

#' Write data and data dictionaries to files
#'
#' You shouldn't need to call this function directly unless you are using it to write a data export function for another R package. If you are using this package on its own, you will usually want to call `fetchFromAccess(save_to_files = TRUE)`.
#'
#' @param all_tables Output of `fetchFromAccess()`
#' @param data_dir Folder to store data csv's in
#' @param dictionary_dir Folder to store data dictionaries in
#' @param lookup_dir Optional folder to store lookup tables in. If left as `NA`, lookups won't be exported.
#'
#' @export
#'
writeToFiles <- function(all_tables, data_dir = here::here("data", "final"), dictionary_dir = here::here("data", "dictionary"), metadata_filenames = c(tables = "data_dictionary_tables.txt",
                                                                                                                                                       attributes = "data_dictionary_attributes.txt",
                                                                                                                                                       categories = "data_dictionary_categories.txt"),
                         lookup_dir = NA, verbose = FALSE) {

  data <- all_tables$data
  lookups <- all_tables$lookups
  tables_dict <- all_tables$metadata$tables
  fields_dict <- all_tables$metadata$fields
  categories_dict <- all_tables$metadata$categories

  # Write data to csv
  if (!dir.exists(data_dir)) {
    dir.create(data_dir, recursive = TRUE)
  }
  if (verbose) {message(paste0("Writing data to ", data_dir, "..."))}
  lapply(names(data), function(tbl_name) {
    if (verbose) {message(paste0("\t\t", tbl_name, ".csv"))}
    readr::write_csv(data[[tbl_name]],
                     here::here(data_dir, paste0(tbl_name, ".csv")),
                     na = "")
  })

  # Write lookups to csv
  if (!is.na(lookup_dir)) {
    if (!dir.exists(lookup_dir)) {
      dir.create(lookup_dir, recursive = TRUE)
    }
    if (verbose) {message(paste0("\nWriting lookups to ", lookup_dir, "..."))}
    lapply(names(lookups), function(tbl_name) {
      if (verbose) {message(paste0("\t\t", tbl_name, ".csv"))}
      readr::write_csv(lookups[[tbl_name]],
                       here::here(lookup_dir, paste0(tbl_name, ".csv")),
                       na = "")
    })
  }

  # Write dictionaries to file
  if (!dir.exists(dictionary_dir)) {
    dir.create(dictionary_dir, recursive = TRUE)
  }
  if (verbose) {message(paste0("\nWriting metadata to ", dictionary_dir, "..."))}
  if (verbose) {message(paste0("\t\t", metadata_filenames["tables"]))}
  readr::write_tsv(tables_dict, here::here(dictionary_dir, metadata_filenames["tables"]), na = "", append = FALSE)
  if (verbose) {message(paste0("\t\t", metadata_filenames["attributes"]))}
  readr::write_tsv(fields_dict, here::here(dictionary_dir, metadata_filenames["attributes"]), na = "", append = FALSE)
  if (verbose) {message(paste0("\t\t", metadata_filenames["categories"]))}
  readr::write_tsv(categories_dict, here::here(dictionary_dir, metadata_filenames["categories"]), na = "", append = FALSE)
}

#' Fetch and tidy data
#'
#' @param tbl_name Name of data table
#' @param connection Database connection object
#'
#' @return A tibble of tidy data
#'
fetchAndTidy <- function(tbl_name, connection, as.is) {
  tidy_data <- tibble::as_tibble(RODBC::sqlFetch(connection, tbl_name, as.is = as.is, stringsAsFactors = FALSE)) %>%
    dplyr::mutate(dplyr::across(where(is.character), function(x) {
      x %>%
        utf8::utf8_encode() %>%  # Encode text as UTF-8 - this prevents a lot of parsing issues in R
        trimws(whitespace = "[\\h\\v]") %>%  # Trim leading and trailing whitespace
        dplyr::na_if("")  # Replace empty strings with NA
    }))

  return(tidy_data)
}

#' Helper function to get primary R class of each data column
#'
#' @param fields Fields data dictionary, as returned by [fetchFromAccess]
#' @param data List of data tables, as returned by [fetchFromAccess]
#'
#' @return `fields` with an additional rClass column
#'
getRClass <- function(fields, data) {

  sapply(fields$tableName, function(table_name) {
    types <- sapply(data[[table_name]], function(col) {class(col)[1]})
    fields[fields$tableName == table_name, "rClass"] <<- types[fields[fields$tableName == table_name,]$attributeName]
  })

  return(fields)
}

#' Generate column spec from data dictionary
#'
#' Given a fields data dictionary, create a list of column specifications that can be used in [readr::read_csv()] or [vroom::vroom()]
#'
#' @param fields Fields data dictionary, as returned by [fetchFromAccess]
#'
#' @return A list of lists
#' @export
#'
makeColSpec <- function(fields) {
  fields %<>%
    dplyr::mutate(colObject = dplyr::case_when(rClass == "character" ~ "c",
                                               rClass == "logical" ~ "l",
                                               rClass == "integer" ~ "i",
                                               rClass == "numeric" ~ "d",
                                               rClass == "Date" ~ "D",
                                               rClass == "POSIXct" ~ "T",
                                               rClass == "POSIXlt" ~ "T",
                                               rClass == "factor" ~ "f",
                                               TRUE ~ "?")) %>%
    split(~tableName)

  col_spec <- lapply(fields, function(table) {
    spec <- split(table$colObject, table$attributeName)
    return(spec)
  })

  names(col_spec) <- names(fields)

  return(col_spec)
}
