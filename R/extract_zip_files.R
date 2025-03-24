#' Extract ZIP Files in a Given Folder
#' @param folder_path The directory containing ZIP files
#' @return No return value, called for side effects.
#' @export
extract_zip_files <- function(folder_path) {
  zip_files <- fs::dir_ls(folder_path, glob = "*.zip")
  for (zip_file in zip_files) {
    zip_basename <- tools::file_path_sans_ext(basename(zip_file))
    output_dir <- file.path(folder_path, zip_basename)
    if (!dir.exists(output_dir)) dir.create(output_dir)

    utils::unzip(zip_file, exdir = output_dir)
    message("Extracted: ", zip_file)
  }
}
