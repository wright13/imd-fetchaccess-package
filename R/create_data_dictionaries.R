#' @title Create Data Dictionaries
#'
#' @description This function assumes that you are using the Access Metadata Generator in your Access database and using the fetch_from_access function.  This function takes the lists of data, lookup, and metadata tables and performs data wrangling to get into right format for eml
#'
#' @param data list containing data tables
#' @param lookups list containing lookup tables
#' @param metadata list containing metdata tables
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
