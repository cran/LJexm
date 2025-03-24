#' Run the Full Process: Extract, Convert, and Merge PDFs
#' @param folder_path The directory containing files for processing. If NULL, the user is prompted to select a directory interactively.
#' @return No return value, called for side effects.
#' @export

run_app <- function(folder_path = NULL) {
  if (.Platform$OS.type != "windows") {
    stop("Error: LJexm is only supported on Windows.")
  }

  if (is.null(folder_path)) {
    if (interactive() && requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
      folder_path <- rstudioapi::selectDirectory()
    } else {
      folder_path <- utils::choose.dir()
    }
  }

  if (!is.null(folder_path) && dir.exists(folder_path)) {
    message("Processing folder: ", folder_path)
    extract_zip_files(folder_path)
    convert_files_to_pdf(folder_path)
    merge_pdfs(folder_path)
    message("All tasks completed successfully!")
  } else {
    message("Invalid folder selection or selection was canceled.")
  }
}

