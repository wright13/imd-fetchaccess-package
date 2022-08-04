#' Get the metadata generator template database
#'
#' Copies the metadata generator template database to a location of your choice.
#'
#' @param destination_folder Where to save the template database
#' @param overwrite_existing If the template database already exists in the destination folder, should it be overwritten?
#'
#' @export
#'
getTemplateDatabase <- function(destination_folder, overwrite_existing = FALSE) {
  destination_folder <- normalizePath(destination_folder)
  template_loc <- system.file("MetadataGeneratorTemplate.accdb", package = "fetchaccess")
  file.copy(template_loc, destination_folder, overwrite = overwrite_existing)
  message(paste("Copied", template_loc, "to", destination_folder))
}
