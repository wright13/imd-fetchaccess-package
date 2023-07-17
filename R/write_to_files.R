#' @title Write To Files
#' @description Write data and data dictionaries to files
#'
#' @param all_tables Output of `fetchFromAccess()` and `create_data_dictionaries()`
#' @param data_dir Folder to store data csv's in
#' @param dictionary_dir Folder to store data dictionaries in
#'
write_to_files <- function(all_tables, data_dir = here::here("data", "final"), dictionary_dir = here::here("data", "dictionary")) {

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
