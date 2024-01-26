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
  metadata <- sapply(metadata_tables, fetchAndTidy, connection = connection, as.is = TRUE)
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
    dplyr::mutate(tableName = stringr::str_remove(tableName, paste0("(", data_prefix, ")", collapse = "|")),
                  fileName = paste0(tableName, ".csv")) %>%
    dplyr::select(tableName, fileName, tableDescription)

  # Fields dictionary
  fields_dict <- metadata$editMetadataAttributeInfo %>%
    dplyr::mutate(tableName = stringr::str_remove(tableName, paste0("(", data_prefix, ")", collapse = "|")),
                  class = dplyr::case_when(
                    !is.na(lookup) ~ "categorical",
                    class %in% c("Short Text", "Long Text", "Memo", "Text", "Yes/No", "Hyperlink", "GUID") ~ "character",
                    class %in% c("Number", "Large Number", "Byte", "Integer", "Long Integer", "Single", "Double", "Replication ID", "Decimal", "AutoNumber", "Currency") ~ "numeric",
                    class %in% c("Date/Time", "Date/Time Extended") ~ "Date",
                    TRUE ~ "unknown"),
                  lookup = stringr::str_remove(lookup, paste0("(", lookup_prefix, ")", collapse = "|"))) %>%
    dplyr::select(tableName, attributeName, attributeDefinition, class, unit, dateTimeFormatString, missingValueCode, missingValueCodeExplanation, lookup)

  # Categories dictionary

  # Get attributes that are associated with a lookup
  category_attrs <- dplyr::select(fields_dict, attributeName, lookup) %>%
    dplyr::filter(!is.na(lookup)) %>%
    unique()

  # Get lookups that have code and definition fields defined
  lookup_defs <- metadata$editMetadataLookupDefs %>%
    dplyr::filter(!is.na(keyColumnName) & !is.na(definitionColumnName)) %>%
    dplyr::mutate(lookupName = stringr::str_remove(lookupName, paste0("(", lookup_prefix, ")", collapse = "|")))

  # Throw an error if there are attributes with a lookup table not defined in the list of lookup tables and cols, or vice versa
  if (any(!(category_attrs$lookup %in% lookup_defs$lookupName)) || any(!(lookup_defs$lookupName %in% category_attrs$lookup))) {
    if (any(!(category_attrs$lookup %in% lookup_defs$lookupName))) {
      msg_1 <- paste0("Information missing from tsys_editMetadataLookupDefs for ", paste0(category_attrs$lookup[!(category_attrs$lookup %in% lookup_defs$lookupName)], collapse = ", "), ".")
    } else {
      msg_1 <- NULL
    }

    if (any(!(lookup_defs$lookupName %in% category_attrs$lookup))) {
      msg_2 <- paste0(paste0(lookup_defs$lookupName[!(lookup_defs$lookupName %in% category_attrs$lookup)], collapse = ", "), " not present in lookup column of tsys_editMetadataAttributeInfo.")
    } else {
      msg_2 <- NULL
    }

    msg <- paste0(c(msg_1, msg_2), collapse = "\n")
    stop(msg)
  }

  if (nrow(category_attrs) > 0) {
    categories_dict <- lapply(1:nrow(category_attrs), function(row_num) {
      lkup_name <- category_attrs$lookup[row_num]
      key_col <- lookup_defs$keyColumnName[lookup_defs$lookupName == lkup_name]
      desc_col <- lookup_defs$definitionColumnName[lookup_defs$lookupName == lkup_name]
      categories <- dplyr::select(lookups[[lkup_name]], dplyr::all_of(c(key_col, desc_col))) %>%
        dplyr::mutate(attributeName = category_attrs$attributeName[row_num]) %>%
        dplyr::rename(code = key_col, definition = desc_col) %>%
        dplyr::select(attributeName, code, definition)  # Make sure cols are in the right order

      return(categories)
    })
    categories_dict <- do.call(rbind, categories_dict) %>% unique()
  } else {
    warning("No categorical variables (columns associated with a lookup table) found. Make sure that you have filled in all the lookup and attribute info in the Access tool.")
    categories_dict <- tibble::tibble(attributeName = character(0), code = character(0), definition = character(0))  # make empty tibble if no categories info exists
  }

  if (add_r_classes) {
    fields_dict <- getRClass(fields_dict, data)
  }

  # Set correct column types based on metadata
  # Create column type spec
  col_spec <- fetchaccess::makeColSpec(fields_dict)
  data <- sapply(names(data), function(tbl_name) {
    col_types <- col_spec[[tbl_name]]
    int_cols <- names(col_types)[col_types == "i"]
    char_cols <- names(col_types)[col_types == "c"]
    num_cols <- names(col_types)[col_types == "n"]
    double_cols <- names(col_types)[col_types == "d"]
    logical_cols <- names(col_types)[col_types == "l"]
    factor_cols <- names(col_types)[col_types == "f"]
    date_cols <- names(col_types)[col_types == "D"]
    datetime_cols <- names(col_types)[col_types == "T"]
    time_cols <- names(col_types)[col_types == "t"]

    tbl <- data[[tbl_name]]
    tbl <- dplyr::mutate(tbl, dplyr::across(int_cols, as.integer),
                         dplyr::across(char_cols, as.character),
                         dplyr::across(num_cols, as.numeric),
                         dplyr::across(double_cols, as.double),
                         dplyr::across(logical_cols, as.logical),
                         dplyr::across(factor_cols, as.factor))

    # Format dates and times
    # Change this to an apply fxn eventually
    for (col in c(date_cols, datetime_cols, time_cols)) {
      format_string <- fields_dict$dateTimeFormatString[fields_dict$tableName == tbl_name & fields_dict$attributeName == col]
      if (!is.na(format_string)) {
        format <- QCkit::convert_datetime_format(format_string, convert_z = TRUE)
        tbl <- tbl %>%
          dplyr::mutate(dplyr::across(col, ~ as.POSIXct(.x, format = format))) %>%
          dplyr::mutate(dplyr::across(col, ~ if(!grepl("[Yy]|M|[Dd]", format_string)) {hms::as_hms(.x)} else {.x})) %>%
          dplyr::mutate(dplyr::across(col, ~ if(!grepl("[Hh].*m.*s*", format_string)) {lubridate::as_date(.x)} else {.x}))
      } else {
        stop(paste("No datetime format provided for column", col, "in table", tbl_name))
      }
    }

    return(tbl)
  }, simplify = FALSE, USE.NAMES = TRUE)

  # Do this again to update column classes after fixing dates and times
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
#' @param dictionary_filenames Named list with names `c("tables", "attributes", "categories")` indicating what to name the tables, attributes, and categories data dictionaries. You are encouraged to keep the default names unless you have a good reason to change them.
#' @param verbose Output feedback to console?
#'
#' @export
#'
writeToFiles <- function(all_tables, data_dir = here::here("data", "final"), dictionary_dir = here::here("data", "dictionary"), dictionary_filenames = c(tables = "data_dictionary_tables.txt",
                                                                                                                                                       attributes = "data_dictionary_attributes.txt",
                                                                                                                                                       categories = "data_dictionary_categories.txt"),
                         lookup_dir = NA, verbose = FALSE) {

  data <- all_tables$data
  lookups <- all_tables$lookups
  tables_dict <- all_tables$metadata$tables
  fields_dict <- all_tables$metadata$fields
  categories_dict <- all_tables$metadata$categories
  col_spec <- makeColSpec(fields_dict)

  # Write data to csv
  if (!dir.exists(data_dir)) {
    dir.create(data_dir, recursive = TRUE)
  }
  if (verbose) {message(paste0("Writing data to ", data_dir, "..."))}
  lapply(names(data), function(tbl_name) {
    if (verbose) {message(paste0("\t\t", tbl_name, ".csv"))}
    tbl <- data[[tbl_name]]
    # Convert dates & times back to character before writing so that they stay in correct format
    col_types <- col_spec[[tbl_name]]
    date_cols <- names(col_types)[col_types == "D"]
    datetime_cols <- names(col_types)[col_types == "T"]
    time_cols <- names(col_types)[col_types == "t"]
    for (col in c(date_cols, datetime_cols, time_cols)) {
      format_string <- fields_dict$dateTimeFormatString[fields_dict$tableName == tbl_name & fields_dict$attributeName == col]
      if (is.na(format_string)) {
        stop(paste("No datetime format provided for column", col, "in table", tbl_name))
      }
      tbl <- tbl %>%
        dplyr::mutate(across(col, ~format(.x, format = QCkit::convert_datetime_format(format_string, convert_z = FALSE))))
    }
    readr::write_csv(tbl,
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
  if (verbose) {message(paste0("\t\t", dictionary_filenames["tables"]))}
  readr::write_tsv(tables_dict, here::here(dictionary_dir, dictionary_filenames["tables"]), na = "", append = FALSE)
  if (verbose) {message(paste0("\t\t", dictionary_filenames["attributes"]))}
  readr::write_tsv(fields_dict, here::here(dictionary_dir, dictionary_filenames["attributes"]), na = "", append = FALSE)
  if (verbose) {message(paste0("\t\t", dictionary_filenames["categories"]))}
  readr::write_tsv(categories_dict, here::here(dictionary_dir, dictionary_filenames["categories"]), na = "", append = FALSE)
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

  names(fields$rClass) <- NULL  # get rid of nonexistent names for this vector

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
                                               rClass == "hms" ~ "t",
                                               rClass == "difftime" ~ "t",
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
