#' @title Fetch data from Access database
#'
#' @description This function assumes that you are using the Access Metadata Generator in your Access database.It is the main function to get the tables from the user defined access database, perform data wrangling, and produce the dicts needed to create eml metadata using the NPS eml tools as part of a data package. This function should perform any necessary data wrangling specific to your dataset and return a named list containing `data`, `lookups`, and `metadata` with contents modified as needed. Do not remove or add tibbles in `data` or `lookups` and do not modify their names. If you add, remove, or rename columns in a tibble in `data`, you must modify the contents of `metadata` accordingly. Do not modify the structure or column names of `metadata`. The structure and column names of `lookups` should also be left as-is. Typically the only necessary modification to `lookups` will be to filter overly large species lists to only include taxa that appear in the data.
#'
#' @param db_path Path to your Access database
#' @param data_prefix Prefix used in your Access database to indicate data export tables and/or queries
#' @param lookup_prefix Prefix used in your Access database to indicate lookup tables
#' @param tables_to_omit Character vector of table names that match the data
#' @param custom_wrangler Optional - function that takes arguments `data`, `lookups`, and `metadata`. All are lists whose names and content correspond to the export data, lookup, and metadata tables in the database after using the database tool.
#' @param save_to_files Should the function save data and data dictionaries to files on hard drive?
#' @inheritParams writeToFiles
#'
#' @return A nested list containing three lists of tibbles: data, lookups, and metadata.
#' @export
#'
#'
# db_path <-"C:\\Users\\EEdson\\OneDrive - DOI\\Projects_InProgress\\SNPL\\GOGA_SNPL\\UpdatingFetchAccess\\SnowyPlovers_BE_20230522_AccessExportTool.accdb"
# custom_wrangler = NA
# data_prefix = "qExport"
# lookup_prefix = "tlu"
# tables_to_omit = c()

fetch_from_access <- function(db_path, data_prefix = "qExport", lookup_prefix = "tlu", tables_to_omit = c(), custom_wrangler, save_to_files = FALSE, data_dir = here::here("data", "final"), dictionary_dir = here::here("data", "dictionary")){

  #get Access tables
  tables <- get_access_tables(db_path, data_prefix,lookup_prefix)


  # Do custom data wrangling (for SNPL need to fix wrangler to use table list)
  if (!missing(custom_wrangler)) {

    new_tables <- custom_wrangler(tables)
    data <- new_tables$data
    lookups <- new_tables$lookups
    metadata <- new_tables$metadata

  } else{
    # just parse out the returned list into 3 tibbles
    data <- tables$data
    lookups <- tables$lookups
    metadata <- tables$metadata
  }

# test metadata tables for missing values etc. before making data dictionaries)
  if(my_counter<-test_metadata(metadata) ==0){

    #create data dictionaries
    all_tables <-create_data_dictionaries(data, lookups, metadata)

    if (save_to_files) {
      writeToFiles(all_tables, data_dir, dictionary_dir)
    }

    return(all_tables)

  } else{

    print("You MUST fix these errors before proceding")

  }


}

