#' @title Create Data Dictionaries
#'
#' @description This function assumes that you are using the Access Metadata Generator in your Access databaseand using the fetch_from_access function.  This function takes the lists of data, looup, and metdata tables and performs data wrangling to get into right format for eml
#'
#' @param db_path Path to your Access database
#' @param data_prefix Prefix used in your Access database to indicate data export tables and/or queries
#' @param lookup_prefix Prefix used in your Access database to indicate lookup tables
#' @param tables_to_omit Character vector of table names that match the data
#' @param custom_wrangler Optional - function that takes arguments `data`, `lookups`, and `metadata`. `data` and `lookups` are lists whose names and content correspond to the data and lookup tables in the database. Names do not include prefixes. `metadata` contains a tibble of field-level metadata called `MetadataAttributes`. See qsys_MetadataAttributes in the Access database for the contents of this tibble. This function should perform any necessary data wrangling specific to your dataset and return a named list containing `data`, `lookups`, and `metadata` with contents modified as needed. Do not remove or add tibbles in `data` or `lookups` and do not modify their names. If you add, remove, or rename columns in a tibble in `data`, you must modify the contents of `metadata` accordingly. Do not modify the structure or column names of `metadata`. The structure and column names of `lookups` should also be left as-is. Typically the only necessary modification to `lookups` will be to filter overly large species lists to only include taxa that appear in the data.
#' @param save_to_files Should the function save data and data dictionaries to files on hard drive?
#' @inheritParams writeToFiles
#'
#' @return A nested list containing three lists of tibbles: data, lookups, and metadata.
#' @export
#'

create_data_dictionaries<- function(data, lookups, metadata){

  ##---- Tables dictionary ----##
  tables_dict <- metadata$MetadataQueries %>%
    dplyr::mutate(tableName = stringr::str_remove(tableName, "qExport"),
                  fileName = paste0(tableName, ".csv")) %>%
    dplyr::select(tableName, fileName, tableDescription)

  ##---- Fields dictionary----##
  fields_dict <- metadata$EDIT_metadataAttributeInfo %>%
    dplyr::mutate(tableName = stringr::str_remove(tableName, "qExport"),
                  class = dplyr::case_when(
                    readonlyClass %in% c("Short Text", "Long Text", "Memo", "Text",
                                         "Yes/No", "Hyperlink") ~ "character",
                    readonlyClass %in% c("Number", "Large Number", "Byte", "Integer",
                                         "Long Integer", "Single", "Double", "Replication ID",
                                         "Decimal", "AutoNumber", "Currency") ~ "numeric",
                    readonlyClass == "Categorical" ~ "categorical",
                    readonlyClass %in% c("Date/Time", "Date/Time Extended") ~ "Date",
                    TRUE ~ "unknown")) %>%
    dplyr::select(tableName,
                  attributeName,
                  "attributeDefinition" = readonlyDescription,
                  class,
                  unit,
                  dateTimeFormatString,
                  missingValueCode,
                  missingValueCodeExplanation)

  ##---- Categories dictionary ----##
  categories_dict <- tibble::tibble()

  for (i in 1:nrow(metadata$MetadataLookupDefs)) {
    lkup <- metadata$MetadataLookupDefs[[i, "lookupTableName"]]
    lkup_code <- metadata$MetadataLookupDefs[[i, "lookupCodeField"]]
    lkup_def <- metadata$MetadataLookupDefs[[i, "lookupDescriptionField"]]
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

  return(all_tables)
}
