#' @title Test Metadata
#' @description function to test that thr metadata tables created in Access as part of the tool, have no missing values. Missing values at this stage would mean that the EML processes would fail later on
#'
#'
#' @param metadata list of metadata tables created as part of the fetch_from_access function
#'
#' @return a numeric value based on the number of tests the input data fails
#' @export
#'
#' @examples
#'
test_metadata<- function(metadata){
  counter<- as.numeric(0)

  if(sum(is.na(metadata$MetadataQueries$tableDescription)) >0){
    cli::cli_alert_danger("WARNING!! Not all export tables have a TABLE DESCRIPTION.")
    cli::cli_alert_info("  You can fix this in the Access database by making sure the export queries have a descrpition entered into their object properties.")
    counter <- counter +1
  }

  if(sum(is.na(metadata$EDIT_metadataAttributeInfo$readonlyClass)) >0){
    cli::cli_alert_danger("WARNING!! Not all field attributes have a CLASS.")
    cli::cli_alert_info("  You can fix this in the Access database by making sure the field readonlyClass in tsys_EDIT_metadataAttributeInfo does not have any missing values.")
    counter <- counter +1
  }

  if(sum(is.na(metadata$EDIT_metadataAttributeInfo$readonlyDescription)) >0){
    cli::cli_alert_danger("WARNING!! Not all field attributes have a DESCRIPTION.")
    cli::cli_alert_info("  You can fix this in the Access database by making sure the field readonlyDescription in tsys_EDIT_metadataAttributeInfo does not have any missing values.")
    counter <- counter +1
  }

  if(sum(is.na(metadata$MetadataLookupDefs$lookupCodeField)) >0){
    cli::cli_alert_danger("WARNING!! Not all used lookup tables have their CODE FIELD specified.")
    cli::cli_alert_info("  You can fix this in the Access database by making sure the field lookupCodeField in tsys_EDIT_metadataSourceFields is filled in correctly for every categorical field.")
    cli::cli_alert_info("  You then need to delete the table tsys_MetadataLookUpDefs and click the CreateMetadata button again")
    counter <- counter +1
  }

  if(sum(is.na(metadata$MetadataLookupDefs$LookupDescriptionField)) >0){
    cli::cli_alert_danger("WARNING!! Not all used lookup tables have their DESCRIPTION FIELD specified.")
    cli::cli_alert_info("  You can fix this in the Access database by making sure the field lookupDescriptionField in tsys_EDIT_metadataSourceFields is filled in correctly for every categorical field.")
    cli::cli_alert_info("  You then need to delete the table tsys_MetadataLookUpDefs and click the CreateMetadata button again")
    counter <- counter +1
  }

  return(counter)
}
