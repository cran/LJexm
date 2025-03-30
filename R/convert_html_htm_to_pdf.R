#' Convert HTML/HTM Files to PDF
#' Automatically installs 'PhantomJS' if missing.
#' @param folder_path The directory where ZIP files are extracted.
#' @return No return value, called for side effects.
#' @export
convert_html_to_pdf <- function(folder_path) {
  # Ensure webshot package is available
  if (!requireNamespace("webshot", quietly = TRUE)) {
    stop("The 'webshot' package is required but not installed.")
  }

  # Check and Install PhantomJS if Needed
  if (!webshot::is_phantomjs_installed()) {
    message("PhantomJS is not installed. Installing now...")
    webshot::install_phantomjs()
    message("PhantomJS installation completed.")
  }

  phantomjs_args <- c("--webdriver=no", "--ignore-ssl-errors=true", "--ssl-protocol=any")

  # Get all subdirectories inside folder_path
  subfolders <- fs::dir_ls(folder_path, type = "directory")

  if (length(subfolders) == 0) {
    message("No extracted folders found. Skipping HTML conversion.")
    return()
  }

  for (subfolder in subfolders) {
    html_files <- fs::dir_ls(subfolder, regexp = "(?i)\\.(html|htm)$", type = "file")

    if (length(html_files) == 0) {
      message("No HTML or HTM files found in ", subfolder, ". Skipping.")
      next
    }

    for (html_file in html_files) {
      pdf_output <- file.path(subfolder, paste0(tools::file_path_sans_ext(basename(html_file)), ".pdf"))
      webshot::webshot(html_file, pdf_output)
      if (file.exists(pdf_output)) {
        message("Converted: ", html_file, " -> ", pdf_output)
      } else {
        message("Failed to convert: ", html_file)
      }
    }
  }
}

